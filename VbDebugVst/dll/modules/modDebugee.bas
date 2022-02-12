Attribute VB_Name = "modDebugee"
' //
' // modDebugee.bas - debugger/debugee initialization, some exported functions
' // By The trick, 2022
' //

Option Explicit

Public g_hRemoteProcess     As Handle
Public g_hRemoteThread      As Handle

' // True when hook to avoid re-initilization
Private m_bIsInitializing   As Boolean
Private m_bIsInitialized    As Boolean
Private m_hWnd              As Handle   ' // Local handle of remote communication window

' // Initialize debugger side data
Public Function InitializeDebugger( _
                ByVal hMainHwnd As Handle, _
                ByVal hWndCallback As Handle) As Long
    Dim hr  As Long
    
    With g_tSharedData(0)
        
        .hWndApp = hMainHwnd
        .hWndCallback = hWndCallback
        .lDebuggerThreadId = GetCurrentThreadId
        .lDebuggerProcessId = GetCurrentProcessId
        .lEventsBufSize = EVENTS_BUFFER_SIZE
        .lEventsCount = 0
        .pEventsBuf = g_pSharedData + LenB(g_tSharedData(0))
        .lSamplesBufSize = SAMPLES_BUFFER_SIZE
        .lSamplesCount = 0
        .pSamplesBuf = .pEventsBuf + EVENTS_BUFFER_SIZE * 32
        .lAutomationBufSize = AUTOMATION_BUFFER_SIZE
        .lAutomationCount = 0
        .pAutomationBuf = .pSamplesBuf + SAMPLES_BUFFER_SIZE * 4
        .lDataBufferSize = DATA_BUFFER_SIZE
        .pDataBuffer = .pAutomationBuf + AUTOMATION_BUFFER_SIZE * 8
        memset .tCurTimeInfo, LenB(.tCurTimeInfo), 0

    End With

CleanUp:

    InitializeDebugger = hr
    
End Function

' // Unitialize both debugger and remote process
Public Sub UninitializeDebugger()
        
    ReleaseProxyData
    
    UninitializeDebugee

    With g_tSharedData(0)
        
        .lDebuggerThreadId = 0
        .lDebuggerProcessId = 0
        .lEventsBufSize = 0
        .lEventsCount = 0
        .pEventsBuf = NULL_PTR
        .lSamplesBufSize = 0
        .lSamplesCount = 0
        .pSamplesBuf = NULL_PTR
        .lAutomationBufSize = 0
        .lAutomationCount = 0
        .pAutomationBuf = NULL_PTR
        .lDataBufferSize = 0
        .pDataBuffer = NULL_PTR
        memset .tCurTimeInfo, LenB(.tCurTimeInfo), 0
        
    End With
    
End Sub

' // Initialize remote process.
' // It searches for the VB6 instance with AX project with specified ProgID class
Public Function InitializeDebugee( _
                ByRef sProgId As String) As Long
    Dim hr          As Long
    Dim lThreadId   As Long
    Dim hHook       As Handle
    Dim hEvent      As Handle
    Dim hProcess    As Handle
    Dim hThread     As Handle
    Dim lPID        As Long
    
    hr = FindPluginThread(sProgId, lThreadId)
    If hr < 0 Then
        GoTo CleanUp
    End If
    
    lPID = ProcessIdFromThreadId(lThreadId)
    If lPID = 0 Then
        hr = E_FAIL
        GoTo CleanUp
    End If

    hEvent = CreateEvent(ByVal NULL_PTR, 0, 0, vbNullString)
    If hEvent = 0 Then
        hr = HRESULTFromWin32(GetLastError)
        GoTo CleanUp
    End If
    
    hThread = OpenThread(SYNCHRONIZE Or THREAD_QUERY_INFORMATION, 0, lThreadId)
    If hThread = 0 Then
        hr = HRESULTFromWin32(GetLastError)
        GoTo CleanUp
    End If
    
    hProcess = OpenProcess(PROCESS_DUP_HANDLE Or PROCESS_CREATE_THREAD Or PROCESS_QUERY_INFORMATION Or _
                           PROCESS_VM_OPERATION Or PROCESS_VM_WRITE Or PROCESS_VM_READ, 0, lPID)
    If hProcess = 0 Then
        hr = HRESULTFromWin32(GetLastError)
        GoTo CleanUp
    End If
    
    If DuplicateHandle(-1, hEvent, hProcess, g_tSharedData(0).hEvent, 0, 0, DUPLICATE_SAME_ACCESS) = 0 Then
        hr = HRESULTFromWin32(GetLastError)
        GoTo CleanUp
    End If
    
    ' // Install hook to remote thread
    hHook = SetWindowsHookEx(WH_GETMESSAGE, AddressOf InitRemoteProc, g_hInstance, lThreadId)
    If hHook = 0 Then
        hr = HRESULTFromWin32(GetLastError)
        GoTo CleanUp
    End If
    
    ' // Trigger hook
    If PostThreadMessage(lThreadId, 0, 0, ByVal NULL_PTR) = 0 Then
        hr = HRESULTFromWin32(GetLastError)
        GoTo CleanUp
    End If
    
    If WaitForSingleObject(hEvent, 15000) <> WAIT_OBJECT_0 Then
        hr = E_FAIL
        GoTo CleanUp
    End If
    
    hr = g_tSharedData(0).hr
    
    If hr >= 0 Then

        m_hWnd = g_tSharedData(0).hWnd
        g_hRemoteProcess = hProcess
        g_hRemoteThread = hThread
        lstrcpyn g_tSharedData(0).pszProgId(0), ByVal StrPtr(sProgId), 39
        m_bIsInitialized = True
        
    End If
    
CleanUp:
    
    If hr < 0 Then
        
        If g_tSharedData(0).hEvent Then
            DuplicateHandle hProcess, g_tSharedData(0).hEvent, 0, 0, 0, 0, DUPLICATE_CLOSE_SOURCE
        End If
        
        If hProcess Then
            CloseHandle hProcess
        End If
        
        If hThread Then
            CloseHandle hThread
        End If
        
    End If
    
    If hEvent Then
        CloseHandle hEvent
    End If
    
    If hHook Then
        UnhookWindowsHookEx hHook
    End If

    InitializeDebugee = hr
    
End Function

' // Unitialize remote side data
Public Sub UninitializeDebugee()
    Dim pfnFreeLib  As PTR
    
    If m_bIsInitialized Then

        SendMessage m_hWnd, WM_EXIT_THREAD, 0, ByVal NULL_PTR

        pfnFreeLib = GetProcAddress(GetModuleHandle("kernel32"), "FreeLibrary")
        
        ' // Decrement library counter
        If pfnFreeLib Then
            CloseHandle CreateRemoteThread(g_hRemoteProcess, ByVal NULL_PTR, 0, pfnFreeLib, ByVal g_tSharedData(0).hDllRemote, 0, 0)
        End If
        
        CloseHandle g_hRemoteProcess
        CloseHandle g_hRemoteThread
        
        g_hRemoteProcess = 0
        g_hRemoteThread = 0
        m_hWnd = 0
        m_bIsInitialized = False
        
    End If
    
End Sub

' // Create new plugin instance
Public Function CreatePluginInstance( _
                ByRef ppObject As PTR) As Long
    
    ppObject = NULL_PTR
    
    If Not m_bIsInitialized Then
        CreatePluginInstance = E_UNEXPECTED
        Exit Function
    End If
    
    CreatePluginInstance = SendMessage(m_hWnd, WM_CREATE_PLUGIN, 0, ByVal NULL_PTR)
    
    If CreatePluginInstance >= 0 Then
        ppObject = GetProxyObject
    End If
    
End Function

Public Function DestroyPluginInstance() As Long

    If Not m_bIsInitialized Then
        DestroyPluginInstance = E_UNEXPECTED
        Exit Function
    End If
    
    DestroyPluginInstance = SendMessage(m_hWnd, WM_DESTROY_PLUGIN, 0, ByVal NULL_PTR)
    
End Function

' // Call EbMode remote
Public Function GetHostState() As Long

    If Not m_bIsInitialized Then
        GetHostState = E_UNEXPECTED
        Exit Function
    End If
    
    GetHostState = SendMessage(m_hWnd, WM_HOST_STATE, 0, ByVal NULL_PTR)
    
End Function

Public Function IsServerAlive() As Long
    Dim lThreadState    As Long
    
    If Not m_bIsInitialized Then
        IsServerAlive = E_UNEXPECTED
    Else
        If GetExitCodeThread(g_hRemoteThread, lThreadState) = 0 Then
            IsServerAlive = HRESULTFromWin32(GetLastError)
        Else
            If lThreadState <> STILL_ACTIVE Then
                IsServerAlive = S_FALSE
            Else
                IsServerAlive = S_OK
            End If
        End If
    End If
    
End Function

Public Function SendRequestGeneric( _
                ByVal lRequest As Long, _
                ByVal wParam As PTR, _
                ByVal lParam As PTR) As Long
    Dim lThreadState    As Long
    Dim tMSG            As MSG
    Dim lRet            As Long
    
    If Not m_bIsInitialized Then
        SendRequestGeneric = E_UNEXPECTED
    Else
        If GetExitCodeThread(g_hRemoteThread, lThreadState) = 0 Then
            SendRequestGeneric = HRESULTFromWin32(GetLastError)
        Else
            If lThreadState <> STILL_ACTIVE Then
                SendRequestGeneric = RPC_E_DISCONNECTED
            Else
                SendRequestGeneric = SendMessage(g_tSharedData(0).hWnd, lRequest, wParam, ByVal lParam)
            End If
        End If
    End If
    
End Function

' // This is called in remote process
Private Function InitRemoteProc( _
                 ByVal lCode As Long, _
                 ByVal wParam As PTR, _
                 ByRef lParam As MSG) As Long
    Dim hEvent  As Handle
    
    If lCode = HC_ACTION Then
        If Not m_bIsInitializing Then
            
            ' // Prevent further calls
            m_bIsInitializing = True
            
            If InitializeRemote() Then
                ' // Increment ref counter
                GetModuleHandleEx GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS, ByVal FAR_PROC(AddressOf InitializeRemote), (g_hInstance)
            End If
            
            hEvent = g_tSharedData(0).hEvent
            g_tSharedData(0).hEvent = 0 ' // Before zero then close
            
            ' // Unhold main process
            SetEvent hEvent
            CloseHandle hEvent

        End If
    End If
    
    InitRemoteProc = CallNextHookEx(0, lCode, wParam, VarPtr(lParam))
    
End Function

