Attribute VB_Name = "modMain"
' //
' // modMain.bas - the first module of VbDebugVstDll library
' // Use Standard EXE type project to make native DLL
' // It uses external module definition file (Library.def) to specify export
' // By The trick, 2022
' //

Option Explicit

Public Const SAMPLES_BUFFER_SIZE    As Long = 100000    ' // Per samples
Public Const EVENTS_BUFFER_SIZE     As Long = 5000      ' // Per VstEvent items
Public Const AUTOMATION_BUFFER_SIZE As Long = 100       ' // Per tAutomationRecord items
Public Const DATA_BUFFER_SIZE       As Long = 4096
Public Const SHARED_MEMORY_SIZE     As Long = SAMPLES_BUFFER_SIZE * 4 + EVENTS_BUFFER_SIZE * 32 + DATA_BUFFER_SIZE + _
                                              AUTOMATION_BUFFER_SIZE * 8 + 1024
Public Const SHARED_MEMORY_NAME     As String = "VbDebugVst_SharedMem"

' // The data is shared between 2 processes
Public g_pSharedData    As PTR
Public g_tSharedData()  As SHARED_DATA
Public g_hInstance      As PTR

Private m_hMap          As Handle
Private m_tSharedDataSA As SAFEARRAY1D

Sub Main()

End Sub

Public Function DllMain( _
                ByVal hInstance As PTR, _
                ByVal fdwReason As Long, _
                ByVal lpReserved As PTR) As Long
    
    Select Case fdwReason
    Case DLL_PROCESS_ATTACH
        g_hInstance = hInstance
        DllMain = Initialize And 1
    Case DLL_PROCESS_DETACH
        Uninitialize
    End Select

End Function

Public Function GetSharedData() As PTR
    GetSharedData = g_pSharedData
End Function

Private Function Initialize() As Boolean
    Dim hMap    As Handle
    Dim pShared As PTR

    If g_pSharedData Then
        Initialize = True
        Exit Function
    End If
    
    hMap = CreateFileMapping(INVALID_HANDLE_VALUE, ByVal NULL_PTR, PAGE_READWRITE, 0, SHARED_MEMORY_SIZE, SHARED_MEMORY_NAME)
    If hMap = 0 Then
        GoTo CleanUp
    End If
    
    pShared = MapViewOfFile(hMap, FILE_MAP_WRITE, 0, 0, 0)
    If pShared = 0 Then
        GoTo CleanUp
    End If
    
    With m_tSharedDataSA
        .cDims = 1
        .fFeatures = FADF_AUTO
        .rgsabound(0).cElements = 1
        .pvData = pShared
    End With
    
    PutMemPtr ByVal ArrPtr(g_tSharedData), VarPtr(m_tSharedDataSA)
    
    m_hMap = hMap
    g_pSharedData = pShared
    
    Initialize = True
    
CleanUp:
    
    If Not Initialize Then
        
        If pShared Then
            UnmapViewOfFile pShared
        End If
        
        If hMap Then
            CloseHandle hMap
        End If
        
    End If
    
End Function

Private Sub Uninitialize()

    If g_pSharedData = NULL_PTR Then
        Exit Sub
    End If

    With g_tSharedData(0)
    
        If .lDebuggerProcessId = GetCurrentProcessId Then
            UninitializeDebugger
        ElseIf .lVstProcessId = GetCurrentProcessId Then
            UninitializeDebugee
        End If
        
    End With
    
    PutMemPtr ByVal ArrPtr(g_tSharedData), NULL_PTR
    
    UnmapViewOfFile g_pSharedData
    CloseHandle m_hMap
    
    g_pSharedData = NULL_PTR
    m_hMap = 0
    
End Sub

Public Function ProcessIdFromThreadId( _
                ByVal lThreadId As Long) As Long
    Dim hThread As Handle
    Dim tTBI    As THREAD_BASIC_INFORMATION
    
    hThread = OpenThread(THREAD_QUERY_INFORMATION, 0, lThreadId)
    If hThread = 0 Then
        Exit Function
    End If
    
    If NtQueryInformationThread(hThread, ThreadBasicInformation, tTBI, Len(tTBI), 0) >= 0 Then
        ProcessIdFromThreadId = tTBI.ClientId.UniqueProcess
    End If
    
    CloseHandle hThread
    
End Function

Public Function FAR_PROC( _
                ByVal pfn As PTR) As PTR
    FAR_PROC = pfn
End Function

Public Function HRESULTFromWin32( _
                ByVal lError As Long) As Long
    HRESULTFromWin32 = &H80070000 Or (lError And &HFFFF&)
End Function

Public Function IID_IVBVstEffect() As UUID
    PutMem8 IID_IVBVstEffect.Data1, &HA41826AA, &H431FAB40
    PutMem8 IID_IVBVstEffect.Data4(0), &H5A1775A9, &H1E9A86FB
End Function

