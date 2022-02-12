Attribute VB_Name = "modWaveOut"
' //
' // modWaveOut.bas - audio playback module
' // by The trick, 2022
' //

Option Explicit

Private Const CALLBACK_WINDOW_CLASS As String = "VbDebugVst_AudioCallbackWndClass"
Private Const NUM_OF_BUFFERS        As Long = 2
Private Const BUFFER_SIZE_SAMPLES   As Long = &H1000

Private Type tBuffer
    iData() As Integer
    tHeader As WAVEHDR
End Type

Private m_bIsInitialized    As Boolean
Private m_bIsPlayed         As Boolean
Private m_hWndCallback      As Handle
Private m_hWaveOut          As Handle
Private m_tBuffers()        As tBuffer

' // Start playback (issues NewDataRequired)
Public Function StartPlayback() As Boolean
    Dim lIndex  As Long
    Dim lErr    As Long
    
    If m_bIsPlayed Then
        StartPlayback = True
        Exit Function
    End If
    
    For lIndex = 0 To 1
    
        If Not NewDataRequired(m_tBuffers(lIndex).iData, BUFFER_SIZE_SAMPLES) Then
            memset m_tBuffers(lIndex).iData(0), Len(m_tBuffers(lIndex).iData(0)) * BUFFER_SIZE_SAMPLES * 2, 0
        End If
        
        lErr = waveOutWrite(m_hWaveOut, m_tBuffers(lIndex).tHeader, Len(m_tBuffers(lIndex)))
        If lErr Then
            Log "waveOutWrite failed 0x" & Hex$(lErr)
            Exit Function
        End If
        
    Next
    
    m_bIsPlayed = True
    StartPlayback = True
    
End Function

' // Stop playback
Public Function StopPlayback() As Boolean
    Dim lErr    As Long
    
    lErr = waveOutReset(m_hWaveOut)
    
    StopPlayback = lErr = 0
    
    If lErr Then
        Log "waveOutReset failed 0x" & Hex$(lErr)
    End If
    
    m_bIsPlayed = False
    
End Function

' // Initialize playback
Public Function InitializeAudioOut( _
                ByVal lSampleRate As Long) As Boolean
    Dim tWndClass                       As WNDCLASSEX
    Dim tWavFmt                         As WAVEFORMATEX
    Dim lErr                            As Long
    Dim lIndex                          As Long
    Dim bPrepState(NUM_OF_BUFFERS - 1)  As Boolean
    
    ReDim m_tBuffers(NUM_OF_BUFFERS - 1)
    
    If m_bIsInitialized Then
        InitializeAudioOut = True
        Exit Function
    End If
    
    With tWndClass
        .cbSize = LenB(tWndClass)
        .hInstance = App.hInstance
        .lpszClassName = StrPtr(CALLBACK_WINDOW_CLASS)
        .lpfnWndProc = FAR_PROC(AddressOf WndProcCallback)
    End With
    
    If RegisterClassEx(tWndClass) = 0 Then
        Log "RegisterClassEx failed 0x" & Hex$(GetLastError)
        Exit Function
    End If
    
    m_hWndCallback = CreateWindowEx(0, CALLBACK_WINDOW_CLASS, vbNullString, 0, 0, 0, 0, 0, HWND_MESSAGE, 0, App.hInstance, ByVal NULL_PTR)
    If m_hWndCallback = 0 Then
        Log "CreateWindowEx failed 0x" & Hex$(GetLastError)
        GoTo CleanUp
    End If
    
    With tWavFmt
    
        .cbSize = 0
        .wFormatTag = WAVE_FORMAT_PCM
        .wBitsPerSample = 16
        .nSamplesPerSec = lSampleRate
        .nChannels = 2
        .nBlockAlign = .nChannels * .wBitsPerSample \ 8
        .nAvgBytesPerSec = .nSamplesPerSec * .nBlockAlign
        
    End With
    
    lErr = waveOutOpen(m_hWaveOut, WAVE_MAPPER, tWavFmt, m_hWndCallback, 0, CALLBACK_WINDOW)
    If lErr Then
        Log "waveOutOpen failed 0x" & Hex$(lErr)
        GoTo CleanUp
    End If
    
    For lIndex = 0 To UBound(m_tBuffers)
        
        With m_tBuffers(lIndex)
            
            ReDim .iData(BUFFER_SIZE_SAMPLES * 2 - 1)
            .tHeader.lpData = VarPtr(.iData(0))
            .tHeader.dwBufferLength = BUFFER_SIZE_SAMPLES * 2 * Len(.iData(0))
            .tHeader.dwUser = lIndex
            
            lErr = waveOutPrepareHeader(m_hWaveOut, .tHeader, Len(.tHeader))
            
            bPrepState(lIndex) = lErr = 0
            
            If Not bPrepState(lIndex) Then
                Log "waveOutPrepareHeader failed 0x" & Hex$(lErr)
                GoTo CleanUp
            End If
            
        End With

    Next
    
    InitializeAudioOut = True
    m_bIsInitialized = True
    
CleanUp:
    
    If Not InitializeAudioOut Then
        
        If m_hWaveOut Then
        
            For lIndex = 0 To UBound(m_tBuffers)
                If bPrepState(lIndex) Then
                    waveOutUnprepareHeader m_hWaveOut, m_tBuffers(lIndex).tHeader, Len(m_tBuffers(lIndex).tHeader)
                End If
            Next
            
            waveOutClose m_hWaveOut
            
        End If
        
        m_hWaveOut = 0
        
        If m_hWndCallback Then
            DestroyWindow m_hWndCallback
        End If
        
        m_hWndCallback = 0
        
        UnregisterClass CALLBACK_WINDOW_CLASS, App.hInstance
        
    End If
                        
End Function

' // Uninitialize sound
Public Sub UninitializeAudioOut()
    Dim lIndex  As Long
    
    If Not m_bIsInitialized Then
        Exit Sub
    End If
    
    m_bIsInitialized = False
    m_bIsPlayed = False
    
    waveOutReset m_hWaveOut

    For lIndex = 0 To UBound(m_tBuffers)
        waveOutUnprepareHeader m_hWaveOut, m_tBuffers(lIndex).tHeader, Len(m_tBuffers(lIndex).tHeader)
    Next
    
    waveOutClose m_hWaveOut
    DestroyWindow m_hWndCallback
    UnregisterClass CALLBACK_WINDOW_CLASS, App.hInstance
    
    m_hWaveOut = 0
    m_hWndCallback = 0
    
End Sub

Private Function NewDataRequired( _
                 ByRef iSamples() As Integer, _
                 ByVal lLength As Long) As Boolean
    Dim fBufferL()  As Single
    Dim fBufferR()  As Single
    Dim lIndex      As Long
    Dim fValue      As Single

    If Not RequestSamples(fBufferL, fBufferR, BUFFER_SIZE_SAMPLES) Then
        Exit Function
    End If
    
    For lIndex = 0 To BUFFER_SIZE_SAMPLES - 1
        
        fValue = fBufferL(lIndex)
        
        If fValue > 1 Then
            fValue = 1
        ElseIf fValue < -1 Then
            fValue = -1
        End If
        
        iSamples(lIndex * 2) = fValue * 32767
        
        fValue = fBufferR(lIndex)
        
        If fValue > 1 Then
            fValue = 1
        ElseIf fValue < -1 Then
            fValue = -1
        End If
        
        iSamples(lIndex * 2 + 1) = fValue * 32767
        
    Next
    
    NewDataRequired = True
    
End Function

Private Function WndProcCallback( _
                 ByVal hWnd As Handle, _
                 ByVal lMsg As Long, _
                 ByVal wParam As PTR, _
                 ByVal lParam As PTR) As PTR
    Dim tHeader As WAVEHDR
    
    Select Case lMsg
    Case MM_WOM_DONE
        
        If Not m_bIsPlayed Then
            Exit Function
        End If
        
        memcpy tHeader, ByVal lParam, Len(tHeader)
        
        If Not NewDataRequired(m_tBuffers(tHeader.dwUser).iData, BUFFER_SIZE_SAMPLES * 2) Then
            memset m_tBuffers(tHeader.dwUser).iData(0), BUFFER_SIZE_SAMPLES * 2 * Len(m_tBuffers(tHeader.dwUser).iData(0)), 0
        End If
        
        waveOutWrite m_hWaveOut, m_tBuffers(tHeader.dwUser).tHeader, Len(m_tBuffers(tHeader.dwUser).tHeader)
        
    Case Else
        WndProcCallback = DefWindowProc(hWnd, lMsg, wParam, ByVal lParam)
    End Select
    
End Function

Private Function FAR_PROC( _
                 ByVal pfn As PTR) As PTR
    FAR_PROC = pfn
End Function

