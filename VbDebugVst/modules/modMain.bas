Attribute VB_Name = "modMain"
' //
' // modMain.bas - the main module of VbDebugVst executable
' // By The trick, 2022
' //
' // This is very-very simple DAW specially developed to debug VST plugins in VB6
' // Much features isn't supported except for the basic ones other things are very crude
' //
' // The external plugin works in remote process to avoid deadlocks during debugging
' // so when you use editor features the window live outside MDI window. You can use show button to show it.
' // It requires external native-dll (VbDebugVstDll.dll) written in VB6 (in dll folder).
' //

Option Explicit

Public Const NUM_OF_BARS            As Long = 8
Public Const EVENTS_QUANTIZATION    As Long = 96    ' // PPQ
Public Const NUM_OF_EVENT_TICKS     As Long = EVENTS_QUANTIZATION * 4 * NUM_OF_BARS
Public Const NUM_OF_MIDI_EVENTS     As Long = 5

Public Type tKey
    lValue          As Long
    dVelocity       As Double
    dPos            As Double
    dLength         As Double
End Type

Public Type tPattern
    lLengthPerBeats As Long
    lNumOfKeys      As Long
    tKeys()         As tKey
End Type

Public Type tAudioFile
    dLenPerBars     As Double
    lSampleRate     As Long
    lNumOfSamples   As Long
    bIsMono         As Boolean
    fSamples()      As Single
End Type

Public Type tEvents
    sName           As String
    bCanBeAutomated As Boolean
    bInit           As Boolean
    dEvents()       As Double
End Type

Private Enum eBarEventType
    BET_INITIALIZE = 0
    BET_START_AUDIO = 1
    BET_STOP_AUDIO = 2
    BET_END_SONG = 3
    BET_START_SONG = 4
    BET_PLAYING = 5
End Enum

Private Type tBarEvent
    dPos            As Double
    lIndex          As Long
    eType           As eBarEventType
End Type

Public Type tSong
    tPattern                        As tPattern
    tAudioFile                      As tAudioFile
    bMidiTrack(NUM_OF_BARS - 1)     As Boolean
    bAudioTrack(NUM_OF_BARS - 1)    As Boolean
    lNumOfEvents                    As Long
    tEvents()                       As tEvents
End Type

Public g_tSong           As tSong

Private m_dTempo         As Double
Private m_lBlockSize     As Long
Private m_lUniqueID      As Long
Private m_lSampleRate    As Long
Private m_sProgID        As String
Private m_bVstConnected  As Boolean
Private m_sEffectName    As String
Private m_sVendorString  As String
Private m_sProductString As String
Private m_lNumOfPrograms As Long
Private m_lNumOfParams   As Long
Private m_lNumOfInputs   As Long
Private m_lNumOfOutputs  As Long
Private m_lVstVersion    As Long
Private m_cVstPlugin     As IVBVstEffect_dbg
Private m_bSupportMidi   As Boolean
Private m_bIsChunk       As Boolean
Private m_dPlaybackPos   As Double
Private m_bUseEditor     As Boolean

Public Sub Main()

    If Not InitializeAll Then
        UninitializeAll
    Else
        frmMain.Show
    End If

End Sub

' // Destroy current instance
Public Sub DestroyPlugin()
    Dim hr  As Long
    
    If m_cVstPlugin Is Nothing Then
        Exit Sub
    End If
    
    hr = m_cVstPlugin.Suspend()
    If hr < 0 Then
        Log "Suspend failed 0x" & Hex$(hr)
    End If
    
    Set m_cVstPlugin = Nothing
    
    hr = DestroyPluginRemote()
    If hr < 0 Then
        Log "DestroyPluginRemote failed 0x" & Hex$(hr)
    End If
    
    Log "Plugin has been destroyed 0x" & Hex$(hr)
    
    m_bVstConnected = False
    
    Set frmVSTSite.Plugin = Nothing
    
End Sub

' // Create new plugin
Public Function InitializePlugin() As Boolean
    Dim cPlugin As IVBVstEffect_dbg
    Dim lIndex  As Long
    Dim hr      As Long
    
    m_bVstConnected = False
    
    If Len(m_sProgID) = 0 Then
        Exit Function
    End If
    
    Set cPlugin = CreatePluginRemote(m_sProgID)
    
    If cPlugin Is Nothing Then
        Exit Function
    End If
    
    hr = cPlugin.VstVersion(m_lVstVersion)
    If hr < 0 Then
        Log "VstVersion failed 0x" & Hex$(hr)
        Exit Function
    End If
    
    Log "VstVersion: " & CStr(m_lVstVersion)
    
    hr = cPlugin.SupportsVSTEvents(m_bSupportMidi)
    If hr < 0 Then
        Log "SupportsVSTEvents failed 0x" & Hex$(hr)
        Exit Function
    End If
    
    Log "SupportsVSTEvents: " & CStr(m_bSupportMidi)
    
    hr = cPlugin.SampleRate_put(m_lSampleRate)
    If hr < 0 Then
        Log "SampleRate_put failed 0x" & Hex$(hr)
        Exit Function
    End If
    
    hr = cPlugin.ProgramsAreChunks(m_bIsChunk)
    If hr < 0 Then
        Log "ProgramsAreChunks failed 0x" & Hex$(hr)
        Exit Function
    End If
    
    Log "ProgramsAreChunks: " & CStr(m_bIsChunk)
    
    hr = cPlugin.BlockSize(m_lBlockSize)
    If hr < 0 Then
        Log "BlockSize failed 0x" & Hex$(hr)
        Exit Function
    End If
    
    hr = cPlugin.EffectName(m_sEffectName)
    If hr < 0 Then
        Log "EffectName failed 0x" & Hex$(hr)
        Exit Function
    End If
    
    Log "EffectName: '" & m_sEffectName & "'"
    
    hr = cPlugin.UniqueID(m_lUniqueID)
    If hr < 0 Then
        Log "UniqueId failed 0x" & Hex$(hr)
        Exit Function
    End If
    
    Log "UniqueId: " & CStr(m_lUniqueID)
    
    hr = cPlugin.VendorString(m_sVendorString)
    If hr < 0 Then
        Log "VendorString failed 0x" & Hex$(hr)
        Exit Function
    End If
    
    Log "VendorString: '" & m_sVendorString & "'"
    
    hr = cPlugin.ProductString(m_sProductString)
    If hr < 0 Then
        Log "ProductString failed 0x" & Hex$(hr)
        Exit Function
    End If
    
    Log "ProductString: '" & m_sProductString & "'"
    
    hr = cPlugin.NumOfPrograms(m_lNumOfPrograms)
    If hr < 0 Then
        Log "NumOfPrograms failed 0x" & Hex$(hr)
        Exit Function
    End If
    
    Log "NumOfPrograms: '" & m_lNumOfPrograms & "'"
    
    hr = cPlugin.NumOfParam(m_lNumOfParams)
    If hr < 0 Then
        Log "NumOfParam failed 0x" & Hex$(hr)
        Exit Function
    End If
    
    Log "NumOfParams: '" & m_lNumOfParams & "'"
    
    hr = cPlugin.NumOfInputs(m_lNumOfInputs)
    If hr < 0 Then
        Log "NumOfInputs failed 0x" & Hex$(hr)
        Exit Function
    End If
    
    Log "NumOfInputs: '" & m_lNumOfInputs & "'"
    
    hr = cPlugin.NumOfOutputs(m_lNumOfOutputs)
    If hr < 0 Then
        Log "NumOfOutputs failed 0x" & Hex$(hr)
        Exit Function
    End If
    
    Log "NumOfOutputs: '" & m_lNumOfOutputs & "'"
    
    If m_lNumOfInputs <> 2 Or m_lNumOfOutputs <> 2 Then
        Log "Unsupported number of inputs/outputs (only stereo). Inputs: " & CStr(m_lNumOfInputs) & _
             ", Outputs: " & CStr(m_lNumOfOutputs)
        Exit Function
    End If
    
    If m_lNumOfParams > 0 Then
        
        For lIndex = NUM_OF_MIDI_EVENTS To g_tSong.lNumOfEvents - 1
            Unload frmMain.mnuEventItem(lIndex)
        Next
        
        ReDim Preserve g_tSong.tEvents(m_lNumOfParams + NUM_OF_MIDI_EVENTS - 1)

        For lIndex = 0 To m_lNumOfParams - 1
            
            With g_tSong.tEvents(NUM_OF_MIDI_EVENTS + lIndex)
            
                hr = cPlugin.ParamName(lIndex, .sName)
                If hr < 0 Then
                    Log "ParamName failed 0x" & Hex$(hr)
                    Exit Function
                End If
                
                Load frmMain.mnuEventItem(NUM_OF_MIDI_EVENTS + lIndex)
                frmMain.mnuEventItem(NUM_OF_MIDI_EVENTS + lIndex).Caption = .sName
                
                If m_lVstVersion >= 2 Then
                
                    hr = cPlugin.CanParameterBeAutomated(lIndex, .bCanBeAutomated)
                    If hr < 0 Then
                        Log "CanParameterBeAutomated failed 0x" & Hex$(hr)
                        Exit Function
                    End If
                    
                Else
                    .bCanBeAutomated = True
                End If
                
            End With
            
        Next
        
        g_tSong.lNumOfEvents = m_lNumOfParams + NUM_OF_MIDI_EVENTS
        
    End If
    
    frmSongEdit.UpdatePluginInfo
    
    hr = cPlugin.Resume()
    If hr < 0 Then
        Log "Resume failed 0x" & Hex$(hr)
        Exit Function
    End If

    Set m_cVstPlugin = cPlugin
    m_bVstConnected = True
    UpdateAutomationParameters m_dPlaybackPos, True
    
    Set frmVSTSite.Plugin = cPlugin
    frmVSTSite.ShowPlugin
    
    InitializePlugin = True
    
End Function

' // Initialize all
Public Function InitializeAll() As Boolean
    Dim sRet    As String
    Dim lSize   As Long
    
    sRet = Space$(40)
    lSize = GetPrivateProfileString("Debug options", "ProgID", vbNullString, sRet, Len(sRet), App.Path & "\config.ini")
    
    If lSize = 0 Then
        Log "GetPrivateProfileString failed 0x" & Hex$(GetLastError)
    Else
        m_sProgID = Mid$(sRet, 1, lSize)
    End If
    
    lSize = GetPrivateProfileString("Debug options", "UseEditor", vbNullString, sRet, Len(sRet), App.Path & "\config.ini")
    
    If lSize = 0 Then
        Log "GetPrivateProfileString failed 0x" & Hex$(GetLastError)
    Else
        m_bUseEditor = Left$(sRet, 1) <> "0"
    End If
    
    If Not InitializeRemote(frmMain.hWnd) Then
        Log "InitializeRemote failed"
        Exit Function
    End If

    If Not InitializeAudioOut(44100) Then
        Log "InitializeAudioOut failed"
        Exit Function
    End If
    
    m_lSampleRate = 44100
    m_lBlockSize = &H400
    m_dTempo = 120
    
    InitSong
    
    InitializeAll = True
    
End Function

Public Sub UninitializeAll()
    UninitializeAudioOut
    UninitializeRemote
End Sub

Private Sub InitSong()
    Dim lIndex  As Long
    
    With g_tSong
    
        .lNumOfEvents = NUM_OF_MIDI_EVENTS
        ReDim .tEvents(.lNumOfEvents - 1)
        
        .tEvents(0).sName = "Pitch MIDI"
        .tEvents(1).sName = "Modulation wheel CC#01"
        .tEvents(2).sName = "Sustain Pedal CC#64"
        .tEvents(3).sName = "Channel pressure MIDI"
        .tEvents(4).sName = "Volume CC#07"
        
        memset .bAudioTrack(0), NUM_OF_BARS * Len(.bAudioTrack(0)), 0
        memset .bMidiTrack(0), NUM_OF_BARS * Len(.bMidiTrack(0)), 0
        .tAudioFile.lNumOfSamples = 0
        Erase .tAudioFile.fSamples
        .tPattern.lLengthPerBeats = 4
        .tPattern.lNumOfKeys = 0
        Erase .tPattern.tKeys
        
        For lIndex = 0 To .lNumOfEvents - 1
            
            If lIndex Then
                Load frmMain.mnuEventItem(lIndex)
            End If
            
            frmMain.mnuEventItem(lIndex).Caption = .tEvents(lIndex).sName
            .tEvents(lIndex).bCanBeAutomated = True
            
        Next
        
    End With
    
End Sub

' // Load state to plugin
Public Function LoadState( _
                ByRef bData() As Byte, _
                ByVal lSize As Long) As Long
    Dim lParamIndex As Long
    Dim fValue      As Single
    Dim hr          As Long
    Dim bRet        As Boolean
    
    If Not m_bVstConnected Then
        LoadState = E_FAIL
        Exit Function
    End If
                
    If m_bIsChunk Then
        
        hr = m_cVstPlugin.SetStateChunk(False, VarPtr(bData(0)), lSize, bRet)
        
        If hr < 0 Then
            
            Log "SetStateChunk failed 0x" & Hex$(hr)
            LoadState = hr
            Exit Function
            
        ElseIf Not bRet Then
        
            Log "SetStateChunk returns false"
            LoadState = E_FAIL
            Exit Function
        
        End If
        
    Else
        
        If lSize <> m_lNumOfParams * 4 Then
        
            Log "Invalid data format"
            LoadState = E_FAIL
            Exit Function
            
        End If
        
        For lParamIndex = 0 To m_lNumOfParams - 1
            
            GetMem4 bData(lParamIndex * 4), fValue
            
            If fValue > 1 Then
                fValue = 1
            ElseIf fValue < 0 Then
                fValue = 0
            End If
            
            hr = m_cVstPlugin.ParamValue_put(lParamIndex, fValue)
            
            If hr < 0 Then
                
                Log "ParamValue_put failed 0x" & Hex$(hr)
                LoadState = hr
                Exit Function
                
            End If
            
            frmVSTSite.ParameterChanged lParamIndex, fValue
            
        Next
        
    End If
    
End Function

' // Save plugin state
Public Function SaveState( _
                ByRef bOut() As Byte, _
                ByRef lSize As Long) As Long
    Dim lParamIndex As Long
    Dim fValue      As Single
    Dim pData       As PTR
    Dim hr          As Long
    
    If Not m_bVstConnected Then
        SaveState = E_FAIL
        Exit Function
    End If
    
    If m_bIsChunk Then
        
        hr = m_cVstPlugin.GetStateChunk(False, pData, lSize)
        
        If hr < 0 Then
        
            Log "GetStateChunk failed 0x" & Hex$(hr)
            SaveState = hr
            Exit Function
            
        End If
        
        If lSize > 0 Then
            
            ReDim bOut(lSize - 1)
            
            memcpy bOut(0), ByVal pData, lSize
            
        End If
        
    Else
        
        lSize = m_lNumOfParams * 4
        
        If lSize > 0 Then
        
            ReDim bOut(lSize - 1)
    
            For lParamIndex = 0 To m_lNumOfParams - 1
                
                hr = m_cVstPlugin.ParamValue_get(lParamIndex, fValue)
                
                If hr < 0 Then
                    
                    Log "ParamValue_get failed 0x" & Hex$(hr)
                    SaveState = hr
                    Exit Function
                    
                End If
                
                GetMem4 fValue, bOut(lParamIndex * 4)
                
            Next
            
        End If
        
    End If
    
End Function

Public Property Get UseEditor() As Boolean
    UseEditor = m_bUseEditor
End Property
Public Property Let UseEditor( _
                    ByVal bValue As Boolean)
                    
    If bValue = m_bUseEditor Then
        Exit Property
    End If
    
    m_bUseEditor = bValue
    
    If WritePrivateProfileString("Debug options", "UseEditor", bValue And 1, App.Path & "\config.ini") = 0 Then
        Log "WritePrivateProfileString failed 0x" & Hex$(GetLastError)
    End If
    
End Property

Public Property Get Tempo() As Double
    Tempo = m_dTempo
End Property
Public Property Let Tempo( _
                    ByVal dValue As Double)
                    
    m_dTempo = dValue
    
    If g_tSong.tAudioFile.lNumOfSamples Then
        g_tSong.tAudioFile.dLenPerBars = dValue * g_tSong.tAudioFile.lNumOfSamples / (m_lSampleRate * 240)
    End If
    
End Property

Public Property Get BlockSize() As Long
    BlockSize = m_lBlockSize
End Property
Public Property Let BlockSize( _
                    ByVal lValue As Long)
    Dim hr  As Long
    
    m_lBlockSize = lValue
    
    If m_bVstConnected Then
        hr = m_cVstPlugin.BlockSize(lValue)
        If hr < 0 Then
            Log "BlockSize failed 0x" & Hex$(hr)
        End If
    End If
    
End Property

Public Property Get SampleRate() As Long
    SampleRate = m_lSampleRate
End Property
Public Property Let SampleRate( _
                    ByVal lValue As Long)
    Dim hr  As Long
    
    If m_lSampleRate = lValue Then
        Exit Property
    End If
    
    StopPlayback
    UninitializeAudioOut
    Resample lValue
    
    If Not InitializeAudioOut(lValue) Then
        Log "InitializeAudioOut failed"
        Exit Property
    End If
    
    m_lSampleRate = lValue
    
    If m_bVstConnected Then
        hr = m_cVstPlugin.SampleRate_put(lValue)
        If hr < 0 Then
            Log "SampleRate_put failed 0x" & Hex$(hr)
        End If
    End If
    
End Property

Public Property Get ProgID() As String
    ProgID = m_sProgID
End Property
Public Property Let ProgID( _
                    ByRef sValue As String)
                    
    If StrComp(sValue, m_sProgID, vbTextCompare) = 0 Then
        Exit Property
    End If
    
    DestroyPlugin
    
    m_sProgID = sValue
    
    UninitializeRemote
    InitializeRemote frmMain.hWnd
    
    If WritePrivateProfileString("Debug options", "ProgID", m_sProgID, App.Path & "\config.ini") = 0 Then
        Log "WritePrivateProfileString failed 0x" & Hex$(GetLastError)
    End If
    
End Property

Public Property Get VstConnected() As Boolean
    VstConnected = m_bVstConnected
End Property
Public Property Get ProgramsAreChunks() As Boolean
    ProgramsAreChunks = m_bIsChunk
End Property
Public Property Get EffectName() As String
    EffectName = m_sEffectName
End Property
Public Property Get VendorString() As String
    VendorString = m_sVendorString
End Property
Public Property Get ProductString() As String
    ProductString = m_sProductString
End Property
Public Property Get NumOfPrograms() As Long
    NumOfPrograms = m_lNumOfPrograms
End Property
Public Property Get NumOfParams() As Long
    NumOfParams = m_lNumOfParams
End Property
Public Property Get NumOfInputs() As Long
    NumOfInputs = m_lNumOfInputs
End Property
Public Property Get NumOfOutputs() As Long
    NumOfOutputs = m_lNumOfOutputs
End Property
Public Property Get VstVersion() As Long
    VstVersion = m_lVstVersion
End Property
Public Property Get UniqueID() As Long
    UniqueID = m_lUniqueID
End Property
Public Property Get VstPlugin() As IVBVstEffect_dbg
    Set VstPlugin = m_cVstPlugin
End Property
Public Property Get SupportMidi() As Boolean
    SupportMidi = m_bSupportMidi
End Property
Public Property Get PlaybackPos() As Double
    PlaybackPos = m_dPlaybackPos
End Property
Public Property Let PlaybackPos( _
                    ByVal dValue As Double)
    m_dPlaybackPos = dValue
    
    ' // Update parameters
    UpdateAutomationParameters dValue, True
    
    ' // Reset all notes
    ResetNotes
    
End Property

' // Collect midi events in period
Public Function CollectEventsInTimeslice( _
                ByVal dStartPos As Double, _
                ByVal dLength As Double) As Long
    Dim lParamIndex As Long
    Dim lStartIndex As Long
    Dim lEndIndex   As Long
    Dim lTickIndex  As Long
    Dim bAddEvent   As Boolean
    Dim lEventIndex As Long
    Dim dDeltaMul   As Double
    Dim lMidiEvent  As Long
    Dim dEndPos     As Double
    Dim lBarIndex   As Long
    Dim dPatOffset  As Double
    Dim dPatEnd     As Double
    Dim lKeyIndex   As Long
    Dim dEventPos   As Double
    Dim dCurPos     As Double
    
    If Not m_bSupportMidi Then
        Exit Function
    End If
    
    lStartIndex = Int(dStartPos * 4 * EVENTS_QUANTIZATION)
    lEndIndex = lStartIndex + Int(dLength * 4 * EVENTS_QUANTIZATION) - 1
    
    If lStartIndex >= NUM_OF_EVENT_TICKS Then
        lStartIndex = NUM_OF_EVENT_TICKS - 1
    ElseIf lStartIndex < 0 Then
        lStartIndex = 0
    End If
    
    If lEndIndex >= NUM_OF_EVENT_TICKS Then
        lEndIndex = NUM_OF_EVENT_TICKS - 1
    ElseIf lEndIndex < 0 Then
        lEndIndex = 0
    End If
    
    dDeltaMul = BarsToSamples(1 / (4 * EVENTS_QUANTIZATION))
    
    For lTickIndex = lStartIndex To lEndIndex
        
        For lParamIndex = 0 To NUM_OF_MIDI_EVENTS - 1
            With g_tSong.tEvents(lParamIndex)
                If .bInit Then
                    
                    If lTickIndex > 0 Then
                        If .dEvents(lTickIndex) <> .dEvents(lTickIndex - 1) Then
                            bAddEvent = True
                        Else
                            bAddEvent = False
                        End If
                    Else
                        bAddEvent = True
                    End If
                    
                    If bAddEvent Then
                        
                        Select Case lParamIndex
                        Case 0
                            ' // Pitch
                            lMidiEvent = .dEvents(lTickIndex) * &H3FFF
                            lMidiEvent = (((lMidiEvent * 2) And &H7F00) Or (lMidiEvent And &H7F)) * &H100 Or &HE0
                        Case 1
                            ' // Modulation wheel
                            lMidiEvent = ((.dEvents(lTickIndex) * &H7F) And &H7F) * &H10000 Or &H1B0&
                        Case 2
                            ' // Sustain pedal
                            lMidiEvent = ((.dEvents(lTickIndex) * &H7F) And &H7F) * &H10000 Or &H40B0&
                        Case 3
                            ' // Channel pressure
                            lMidiEvent = ((.dEvents(lTickIndex) * &H7F) And &H7F) * &H100 Or &HD0
                        Case 4
                            ' // Volume
                            lMidiEvent = ((.dEvents(lTickIndex) * &H7F) And &H7F) * &H10000 Or &H7B0&
                        End Select
                        
                        With g_tEvents(lEventIndex)
                            .byteSize = LenB(g_tEvents(lEventIndex))
                            .Type = kVstMidiType
                            .deltaFrames = lTickIndex * dDeltaMul
                            PutMem4 .Data(8), lMidiEvent
                        End With
                        
                        lEventIndex = lEventIndex + 1
                        
                        If lEventIndex > EventsBufferSize Then
                            GoTo exit_proc
                        End If
                        
                    End If
                    
                End If
            End With
        Next
    
    Next
    
    ' // Collect keys
    dEndPos = dStartPos + dLength
    dCurPos = dStartPos
    
    Do While dCurPos < dEndPos
        
        lBarIndex = Int(dCurPos)
        
        If lBarIndex >= NUM_OF_BARS Then
            Exit Do
        End If
        
        If g_tSong.bMidiTrack(lBarIndex) Then
            
            dPatOffset = (dCurPos - lBarIndex) * 16
            dPatEnd = (dEndPos - lBarIndex) * 16
            
            For lKeyIndex = 0 To g_tSong.tPattern.lNumOfKeys - 1
                
                With g_tSong.tPattern.tKeys(lKeyIndex)
                    
                    If .dPos < dPatOffset And .dPos + .dLength >= dPatOffset Then
                    
                        ' // Note off
                        If .dPos + .dLength < dPatEnd Then
                        
                            dEventPos = .dPos + .dLength
                            lMidiEvent = &H80 Or (.lValue * &H100) Or (((.dVelocity * &H7F) And &H7F) * &H10000)
                            bAddEvent = True
                            
                        Else
                            bAddEvent = False
                        End If
                        
                    ElseIf .dPos >= dPatOffset And .dPos < dPatEnd Then
                    
                        ' // Note start
                        dEventPos = .dPos
                        lMidiEvent = &H90 Or (.lValue * &H100) Or (((.dVelocity * &H7F) And &H7F) * &H10000)
                        bAddEvent = True
                        
                    Else
                        bAddEvent = False
                    End If
                    
                    If bAddEvent Then
                        
                        With g_tEvents(lEventIndex)
                            .byteSize = LenB(g_tEvents(lEventIndex))
                            .Type = kVstMidiType
                            .deltaFrames = BarsToSamples(lBarIndex + (dEventPos / 16) - dStartPos)
                            PutMem4 .Data(8), lMidiEvent
                        End With
                        
                        lEventIndex = lEventIndex + 1
                        
                        If lEventIndex > EventsBufferSize Then
                            GoTo exit_proc
                        End If
                        
                    End If
                    
                End With
                
            Next
        End If
        
        dCurPos = lBarIndex + 1
        
    Loop
    
exit_proc:
    
    CollectEventsInTimeslice = lEventIndex
    
End Function

Public Sub Log( _
           ByRef sMessage As String)
    frmLog.PutLog Now & ": " & sMessage
End Sub

' // Reset all midi notes
Private Sub ResetNotes()

    If Not m_bSupportMidi Then
        Exit Sub
    End If
    
    With g_tEvents(0)
        .byteSize = LenB(g_tEvents(0))
        .Type = kVstMidiType
        .deltaFrames = 0
        PutMem4 .Data(8), &H7BB0&
    End With
    
    ProcessEventsRemote m_cVstPlugin, 0, m_dPlaybackPos, m_dTempo, 1
    
End Sub

' // Update parameters according to position
Private Sub UpdateAutomationParameters( _
            ByVal dPos As Double, _
            Optional ByVal bForce As Boolean)
    Static s_fLastValues()  As Single
    Static s_bInit          As Boolean
    
    Dim lIndex      As Long
    Dim lPosEvent   As Long
    Dim hr          As Long
    
    If m_cVstPlugin Is Nothing Then
        Exit Sub
    End If
    
    If m_lNumOfParams <= 0 Then
        Exit Sub
    ElseIf Not s_bInit And m_lNumOfParams > 0 Then
        ReDim s_fLastValues(m_lNumOfParams - 1)
        s_bInit = True
    ElseIf UBound(s_fLastValues) + 1 < m_lNumOfParams Then
        ReDim s_fLastValues(m_lNumOfParams - 1)
    End If
    
    lPosEvent = Int(dPos * 4 * EVENTS_QUANTIZATION)
    
    If lPosEvent >= NUM_OF_EVENT_TICKS Then
        lPosEvent = NUM_OF_EVENT_TICKS - 1
    ElseIf lPosEvent < 0 Then
        lPosEvent = 0
    End If
    
    For lIndex = 0 To m_lNumOfParams - 1
        With g_tSong.tEvents(lIndex + NUM_OF_MIDI_EVENTS)
            If .bInit And .bCanBeAutomated Then
                If bForce Then

                    hr = m_cVstPlugin.ParamValue_put(lIndex, .dEvents(lPosEvent))
                    If hr < 0 Then
                        Log "ParamValue_put failed 0x" & Hex$(hr)
                    End If
                    
                    s_fLastValues(lIndex) = .dEvents(lPosEvent)
                    
                    frmVSTSite.ParameterChanged lIndex, .dEvents(lPosEvent)
                    
                Else
                    If s_fLastValues(lIndex) <> .dEvents(lPosEvent) Then
                    
                        hr = m_cVstPlugin.ParamValue_put(lIndex, .dEvents(lPosEvent))
                        If hr < 0 Then
                            Log "ParamValue_put failed 0x" & Hex$(hr)
                        End If
                        
                        s_fLastValues(lIndex) = .dEvents(lPosEvent)
                        
                        frmVSTSite.ParameterChanged lIndex, .dEvents(lPosEvent)
                        
                    End If
                End If
            End If
        End With
    Next

End Sub

Private Function FillSamples( _
                 ByRef fSamplesL() As Single, _
                 ByRef fSamplesR() As Single, _
                 ByVal lSamples As Long) As Boolean
    Dim dLenToNextEvt   As Double
    Dim dCurPos         As Double
    Dim lSampleStart    As Long
    Dim lSamplesCount   As Long
    Dim lSampleIndex    As Long
    Dim lIndex          As Long
    Dim tEvent          As tBarEvent

    tEvent.dPos = m_dPlaybackPos
    tEvent.eType = BET_INITIALIZE
    dCurPos = m_dPlaybackPos
    
    Do
        
        GetNextBarEvent tEvent
        
        dLenToNextEvt = tEvent.dPos - dCurPos
        
        Select Case tEvent.eType
        Case BET_START_AUDIO, BET_END_SONG, BET_START_SONG
            
            ' // Fill with silence
            lSamplesCount = BarsToSamples(dLenToNextEvt)
            
            If lSamplesCount > 0 Then
                
                If lSampleIndex + lSamplesCount > lSamples Then
                    lSamplesCount = lSamples - lSampleIndex
                End If
                
                memset fSamplesL(lSampleIndex), lSamplesCount * 4, 0
                memset fSamplesR(lSampleIndex), lSamplesCount * 4, 0
                
                lSampleIndex = lSampleIndex + lSamplesCount
            
            End If
            
        Case BET_STOP_AUDIO
        
            ' // Fill with audio samples
            lSamplesCount = BarsToSamples(dLenToNextEvt)
            lSampleStart = BarsToSamples(dCurPos - tEvent.lIndex)
            
            If lSampleIndex + lSamplesCount > lSamples Then
                lSamplesCount = lSamples - lSampleIndex
            End If
            
            If g_tSong.tAudioFile.bIsMono Then
                For lIndex = 0 To lSamplesCount - 1
                    fSamplesL(lSampleIndex + lIndex) = g_tSong.tAudioFile.fSamples(lSampleStart + lIndex)
                    fSamplesR(lSampleIndex + lIndex) = g_tSong.tAudioFile.fSamples(lSampleStart + lIndex)
                Next
            Else
                For lIndex = 0 To lSamplesCount - 1
                    fSamplesL(lSampleIndex + lIndex) = g_tSong.tAudioFile.fSamples(lSampleStart * 2 + lIndex * 2)
                    fSamplesR(lSampleIndex + lIndex) = g_tSong.tAudioFile.fSamples(lSampleStart * 2 + lIndex * 2 + 1)
                Next
            End If
            
            lSampleIndex = lSampleIndex + lSamplesCount

        End Select
        
        dCurPos = tEvent.dPos
        
    Loop While lSampleIndex < lSamples
    
    m_dPlaybackPos = m_dPlaybackPos + SamplesToBars(lSamples)

    If m_dPlaybackPos >= NUM_OF_BARS Then
        m_dPlaybackPos = m_dPlaybackPos - NUM_OF_BARS
    End If
    
    frmSongEdit.UpdatePlayback
    frmPatternEdit.UpdatePlayback
    
    FillSamples = True
    
End Function

Private Sub GetNextBarEvent( _
            ByRef tEvent As tBarEvent)
    Dim lAudio  As Long
    Dim lIndex  As Long
    Dim lIndex2 As Long
    
    Select Case tEvent.eType
    Case BET_INITIALIZE
            
        lAudio = GetAudioEventStartIndexInMap(tEvent.dPos)
        
        If lAudio >= 0 Then
            
            tEvent.eType = BET_PLAYING
            tEvent.lIndex = lAudio
            Exit Sub
            
        End If
    
    Case BET_END_SONG
    
        tEvent.eType = BET_START_SONG
        tEvent.dPos = 0
        Exit Sub
        
    End Select
    
    Select Case tEvent.eType
    Case BET_PLAYING, BET_START_AUDIO
        
        ' // Search for end
        lIndex = Int(tEvent.dPos)
        lIndex2 = Int(tEvent.lIndex + g_tSong.tAudioFile.dLenPerBars)
        
        If lIndex2 >= NUM_OF_BARS Then
            lIndex2 = NUM_OF_BARS - 1
        End If
        
        For lIndex = lIndex + 1 To lIndex2
            
            ' // Check break
            If g_tSong.bAudioTrack(lIndex) Then
                ' // Found
                tEvent.dPos = lIndex
                tEvent.eType = BET_STOP_AUDIO
                Exit Sub
                
            End If

        Next
        
        tEvent.dPos = tEvent.lIndex + g_tSong.tAudioFile.dLenPerBars
        
        If tEvent.dPos > NUM_OF_BARS Then
            tEvent.dPos = NUM_OF_BARS
        End If
        
        tEvent.eType = BET_STOP_AUDIO
        Exit Sub
        
    Case BET_STOP_AUDIO, BET_START_SONG, BET_INITIALIZE
    
        If tEvent.dPos = NUM_OF_BARS Then
            tEvent.eType = BET_END_SONG
            Exit Sub
        End If
        
        ' // Search for start
        lIndex = Int(tEvent.dPos)
        
        If tEvent.dPos <> lIndex Then
            lIndex = lIndex + 1
        End If

        For lIndex = lIndex To NUM_OF_BARS - 1
            If g_tSong.bAudioTrack(lIndex) Then
                ' // Found
                tEvent.dPos = lIndex
                tEvent.eType = BET_START_AUDIO
                tEvent.lIndex = lIndex
                Exit Sub
                
            End If
        Next
        
        tEvent.eType = BET_END_SONG
        tEvent.dPos = NUM_OF_BARS
    
    End Select
        
        
End Sub

Public Function GetAudioEventStartIndexInMap( _
                ByVal dPos As Double) As Long
    Dim lBar        As Long
    Dim dSampleLen  As Double
    Dim dDelta      As Double
    
    lBar = Int(dPos)

    Do While lBar >= 0 And lBar < NUM_OF_BARS
        
        If g_tSong.bAudioTrack(lBar) Then
                        
            dSampleLen = -SamplesToBars(1)
            dDelta = (dPos - lBar) - g_tSong.tAudioFile.dLenPerBars
            
            If dDelta > dSampleLen Then
                Exit Do
            End If
            
            GetAudioEventStartIndexInMap = lBar
            Exit Function
            
        End If
        
        lBar = lBar - 1

    Loop
    
    GetAudioEventStartIndexInMap = -1
    
End Function

Private Function GetAudioSampleIndexFromPos( _
                 ByVal dPos As Double) As Long
    Dim lBar    As Long
    Dim dOffset As Double
    
    lBar = Int(dPos)

    Do While lBar >= 0
        
        If g_tSong.bAudioTrack(lBar) Then
            
            dOffset = dPos - lBar
            
            If dOffset >= g_tSong.tAudioFile.dLenPerBars Then
                Exit Do
            End If
            
            GetAudioSampleIndexFromPos = BarsToSamples(dOffset)
            Exit Function
            
        End If
        
        lBar = lBar - 1

    Loop
    
    GetAudioSampleIndexFromPos = -1
    
End Function

Public Function RequestSamples( _
                ByRef fSamplesL() As Single, _
                ByRef fSamplesR() As Single, _
                ByVal lSamples As Long) As Boolean
    Dim lReqSamples     As Long
    Dim lIndex          As Long
    Dim dStartPos       As Double
    Dim dTimeSliceLen   As Double
    Dim dTempo          As Double
    Dim hr              As Long
    Dim lEventsCount    As Long
    
    dStartPos = m_dPlaybackPos
    dTempo = m_dTempo
    
    If Not FillSamples(g_fSamplesL, g_fSamplesR, lSamples) Then
        Exit Function
    End If
    
    If m_bVstConnected Then
        
        hr = GetRemoteHostState()
        
        If hr < 0 Then
            Log "GetRemoteHostState failed 0x" & Hex$(hr)
        ElseIf hr = 0 Then
        
            Log "Host has stopped"
            Set m_cVstPlugin = Nothing
            Set frmVSTSite.Plugin = Nothing
            m_bVstConnected = False
            
        ElseIf hr = 2 Then
            Log "Host is paused"
        Else
        
            ' // Send samlpes to VST
            
            Do While lIndex < lSamples
            
                If lSamples - lIndex > m_lBlockSize Then
                    lReqSamples = m_lBlockSize
                Else
                    lReqSamples = lSamples - lIndex
                End If
                
                dTimeSliceLen = SamplesToBars(lReqSamples)
                
                If m_bSupportMidi Then
                    
                    lEventsCount = CollectEventsInTimeslice(dStartPos, dTimeSliceLen)
                    
                    If Not ProcessEventsRemote(m_cVstPlugin, 0, dStartPos * 4, dTempo, lEventsCount) Then
                        Log "ProcessEventsRemote failed"
                    End If
                    
                End If
                
                If Not ProcessSamplesRemote(m_cVstPlugin, lIndex, dStartPos * 4, dTempo, lReqSamples) Then
                    Log "ProcessSamplesRemote failed"
                End If
                
                lIndex = lIndex + lReqSamples
                
                dStartPos = dStartPos + dTimeSliceLen
                
                UpdateAutomationParameters dStartPos, False
                
            Loop
            
        End If
        
    End If
    
    GetMemPtr ByVal ArrPtr(g_fSamplesL), ByVal ArrPtr(fSamplesL)
    GetMemPtr ByVal ArrPtr(g_fSamplesR), ByVal ArrPtr(fSamplesR)
    
    RequestSamples = True
    
End Function

Public Function BarsToSamples( _
                ByVal dPos As Double) As Long
    BarsToSamples = (240 / m_dTempo) * dPos * m_lSampleRate
End Function

Public Function SamplesToBars( _
                ByVal lSamples As Long) As Double
    SamplesToBars = m_dTempo * (lSamples / (m_lSampleRate * 240))
End Function

' // Poor quality (without filtering)
Private Sub Resample( _
            ByVal lSampleRate As Long)
    Dim dRatio          As Double
    Dim fRet()          As Single
    Dim lResultSamples  As Long
    Dim lSrcIndex       As Long
    Dim lDstIndex       As Long
    
    If lSampleRate = g_tSong.tAudioFile.lSampleRate Then
        Exit Sub
    End If
    
    dRatio = g_tSong.tAudioFile.lSampleRate / lSampleRate
    lResultSamples = Int(g_tSong.tAudioFile.lNumOfSamples / dRatio)
    
    ReDim fRet(lResultSamples * (2 - g_tSong.tAudioFile.bIsMono) - 1)
    
    If g_tSong.tAudioFile.bIsMono Then
        For lDstIndex = 0 To lResultSamples - 1
            lSrcIndex = Int(dRatio * lDstIndex)
            fRet(lDstIndex) = g_tSong.tAudioFile.fSamples(lSrcIndex)
        Next
    Else
        For lDstIndex = 0 To lResultSamples - 1
        
            lSrcIndex = Int(dRatio * lDstIndex)
            fRet(lDstIndex * 2) = g_tSong.tAudioFile.fSamples(lSrcIndex * 2)
            fRet(lDstIndex * 2 + 1) = g_tSong.tAudioFile.fSamples(lSrcIndex * 2 + 1)
            
        Next
    End If
    
    g_tSong.tAudioFile.fSamples = fRet
    g_tSong.tAudioFile.lNumOfSamples = lResultSamples
    g_tSong.tAudioFile.dLenPerBars = m_dTempo * lResultSamples / (lSampleRate * 240)
    g_tSong.tAudioFile.lSampleRate = lSampleRate
    
End Sub

Public Function LoadAudioFile( _
                ByRef sFileName As String) As Boolean
    Dim hMMFile     As Handle
    Dim tckRIFF     As MMCKINFO
    Dim tckWAVE     As MMCKINFO
    Dim tckFMT      As MMCKINFO
    Dim tckDATA     As MMCKINFO
    Dim tFMT        As WAVEFORMATEX
    Dim lIndex      As Long
    Dim lOutSamples As Long
    Dim pRawBytes   As PTR
    Dim bArr()      As Byte
    Dim iArr()      As Integer
    Dim tArrDesc    As SAFEARRAY1D
    Dim pSafeArray  As Long
    
    hMMFile = mmioOpen(sFileName, ByVal NULL_PTR, MMIO_READWRITE)
    If hMMFile = 0 Then
        MsgBox "Unable to open file", vbCritical
        Exit Function
    End If
        
    tckWAVE.fccType = mmioStringToFOURCC("WAVE", 0)

    If mmioDescend(hMMFile, tckWAVE, ByVal 0&, MMIO_FINDRIFF) Then
        MsgBox "Isn't valid file", vbCritical
        GoTo CleanUp
    End If
    
    tckFMT.ckid = mmioStringToFOURCC("fmt", 0)
    
    If mmioDescend(hMMFile, tckFMT, tckWAVE, MMIO_FINDCHUNK) Then
        MsgBox "Format chunk not found", vbCritical
        GoTo CleanUp
    End If
    
    If tckFMT.cksize < 0 Then
        MsgBox "Invalid format", vbCritical
        GoTo CleanUp
    End If
    
    ReDim bFMT(tckFMT.cksize - 1)
    
    If mmioRead(hMMFile, bFMT(0), tckFMT.cksize) = -1 Then
        MsgBox "Can't read format", vbCritical
        GoTo CleanUp
    End If
    
    mmioAscend hMMFile, tckFMT, 0
    
    tckDATA.ckid = mmioStringToFOURCC("data", 0)

    If mmioDescend(hMMFile, tckDATA, tckWAVE, MMIO_FINDCHUNK) Then
        MsgBox "Wave data isn't found", vbCritical
        GoTo CleanUp
    End If
    
    If tckDATA.cksize <= 0 Then
        MsgBox "Invalid data size", vbCritical
        GoTo CleanUp
    End If
    
    If tckFMT.cksize > Len(tFMT) Then
        tckFMT.cksize = Len(tFMT)
    End If
    
    memcpy tFMT, bFMT(0), tckFMT.cksize
    
    If (tFMT.wFormatTag <> WAVE_FORMAT_PCM Or tFMT.nChannels > 2 Or tFMT.nChannels <= 0 Or _
        tFMT.nBlockAlign <> tFMT.wBitsPerSample * tFMT.nChannels \ 8) Or _
        Not (tFMT.wBitsPerSample = 8 Or tFMT.wBitsPerSample = 16) Or tFMT.cbSize <> 0 Then
        MsgBox "Unsupported format", vbCritical
        GoTo CleanUp
    End If

    pRawBytes = GlobalAlloc(GMEM_FIXED, ((tckDATA.cksize + 3) \ 4) * 4)
    If pRawBytes = 0 Then
        MsgBox "GlobalAlloc failed", vbCritical
        GoTo CleanUp
    End If
    
    If mmioRead(hMMFile, ByVal pRawBytes, tckDATA.cksize) = -1 Then
        MsgBox "Unable to read wave data", vbCritical
        GoTo CleanUp
    End If

    lOutSamples = tckDATA.cksize \ tFMT.nBlockAlign
    
    If lOutSamples > 0 Then
                
        ReDim g_tSong.tAudioFile.fSamples(tFMT.nChannels * lOutSamples - 1)
        
        pSafeArray = VarPtr(tArrDesc)

        tArrDesc.cDims = 1
        tArrDesc.fFeatures = FADF_AUTO
        tArrDesc.pvData = pRawBytes

        Select Case tFMT.wBitsPerSample
        Case 8
            
            tArrDesc.cbElements = 1
            tArrDesc.rgsabound(0).cElements = lOutSamples * tFMT.nChannels

            PutMemPtr ByVal ArrPtr(bArr), ByVal pSafeArray
            
            For lIndex = 0 To tFMT.nChannels * lOutSamples - 1
                g_tSong.tAudioFile.fSamples(lIndex) = (CLng(bArr(lIndex)) - 128&) / 128
            Next
            
            PutMemPtr ByVal ArrPtr(bArr), ByVal NULL_PTR
            
        Case 16
        
            tArrDesc.cbElements = 2
            tArrDesc.rgsabound(0).cElements = lOutSamples * tFMT.nChannels
            
            PutMemPtr ByVal ArrPtr(iArr), ByVal pSafeArray
            
            For lIndex = 0 To tFMT.nChannels * lOutSamples - 1
                g_tSong.tAudioFile.fSamples(lIndex) = iArr(lIndex) / 32768
            Next
            
            PutMemPtr ByVal ArrPtr(iArr), ByVal NULL_PTR

        End Select
    Else
        Erase g_tSong.tAudioFile.fSamples
    End If
    
    g_tSong.tAudioFile.lNumOfSamples = lOutSamples
    g_tSong.tAudioFile.bIsMono = tFMT.nChannels = 1
    g_tSong.tAudioFile.dLenPerBars = m_dTempo * lOutSamples / (m_lSampleRate * 240)
    g_tSong.tAudioFile.lSampleRate = tFMT.nSamplesPerSec
    
    If g_tSong.tAudioFile.lSampleRate <> m_lSampleRate Then
        Resample m_lSampleRate
    End If
    
    LoadAudioFile = True
    
CleanUp:
    
    If pRawBytes Then
        GlobalFree pRawBytes
    End If
    
    mmioClose hMMFile, 0
    
End Function


