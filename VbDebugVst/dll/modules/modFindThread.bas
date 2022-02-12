Attribute VB_Name = "modFindThread"
' //
' // modFindThread.bas - find the VB6.exe thread with active debugging session for specified class
' // By The trick, 2022
' //

Option Explicit

Private Const WM_REQUEST_COTHREADID As String = "WM_REQUEST_COTHREADID"

Private m_lRequestMessage   As Long ' // WM_REQUEST_COTHREADID
Private m_pfnPrevWndProc    As PTR

' // Search for thread ID by specified ProgID
Public Function FindPluginThread( _
                ByRef sProgId As String, _
                ByRef lThreadId As Long) As Long
    Dim pCoTIDList  As PTR
    Dim lComTID     As Long
    Dim lListCount  As Long
    Dim hr          As Long
    
    hr = FindComThreadIdInRegistry(sProgId, pCoTIDList, lListCount)
    If hr < 0 Then
        GoTo CleanUp
    End If
    
    Do While lListCount > 0
        
        lListCount = lListCount - 1
        
        GetMem4 ByVal pCoTIDList + lListCount * Len(lComTID), lComTID
        
        hr = FindThreadIdFromComThreadId(lComTID, lThreadId)
        If hr >= 0 Then
            Exit Do
        End If

    Loop
    
CleanUp:
    
    If pCoTIDList Then
        CoTaskMemFree pCoTIDList
    End If
    
    FindPluginThread = hr
    
End Function

' // Convert COM-ThreadID to normal ThreadID
Private Function FindThreadIdFromComThreadId( _
                 ByVal lComThreadId As Long, _
                 ByRef lThreadId As Long) As Long
    Dim hWnd        As Handle
    Dim lPID        As Long
    Dim lTID        As Long
    Dim lTestCoId   As Long
    Dim hProcess    As Handle
    Dim sProcName   As String
    Dim hr          As Long
    
    sProcName = Space$(MAX_PATH)
    hr = E_FAIL
    
    Do
        
        hWnd = FindWindowEx(0, hWnd, "wndclass_desked_gsk", vbNullString)
        
        If hWnd = 0 Then
            Exit Do
        End If
        
        lTID = GetWindowThreadProcessId(hWnd, lPID)
        
        hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, lPID)
        If hProcess Then
            If GetModuleFileNameEx(hProcess, NULL_PTR, sProcName, Len(sProcName)) Then
                If lstrcmpi(ByVal PathFindFileName(sProcName), ByVal StrPtr("vb6.exe")) = 0 Then
                    If GetComThreadIdFromWindowHandle(hWnd, lTestCoId) >= 0 Then
                        If lTestCoId = lComThreadId Then
                            
                            lThreadId = lTID
                            CloseHandle hProcess
                            hr = 0
                            Exit Do
                            
                        End If
                    End If
                End If
            End If
        End If
        
        CloseHandle hProcess
        
    Loop
    
    FindThreadIdFromComThreadId = hr
    
End Function

' // Get COM-ThreadId by specified window
Private Function GetComThreadIdFromWindowHandle( _
                 ByVal hWnd As Handle, _
                 ByRef lComThreadId As Long) As Long
    Dim hHook   As Handle
    Dim hr      As Long
    Dim lTID    As Long
    
    lTID = GetWindowThreadProcessId(hWnd, ByVal NULL_PTR)
    
    If m_lRequestMessage = 0 Then
    
        m_lRequestMessage = RegisterWindowMessage(WM_REQUEST_COTHREADID)
        
        If m_lRequestMessage = 0 Then
            hr = HRESULTFromWin32(GetLastError)
            GoTo CleanUp
        End If
    
    End If
    
    ' // Install hook in window's thread
    hHook = SetWindowsHookEx(WH_CALLWNDPROC, AddressOf HookProc, g_hInstance, lTID)
    If hHook = 0 Then
        hr = HRESULTFromWin32(GetLastError)
        GoTo CleanUp
    End If
    
    ' // Trigger hook
    If SendMessageTimeout(hWnd, m_lRequestMessage, 0, ByVal NULL_PTR, SMTO_BLOCK Or SMTO_ERRORONEXIT Or _
                          SMTO_NOTIMEOUTIFNOTHUNG, 5000, lComThreadId) Then
        hr = 0
    Else
        hr = HRESULTFromWin32(GetLastError)
    End If
    
CleanUp:
    
    If hHook Then
        UnhookWindowsHookEx hHook
    End If
    
    GetComThreadIdFromWindowHandle = hr
    
End Function

' // This is called inside VB6.exe remote process
Private Function HookProc( _
                 ByVal lCode As Long, _
                 ByVal wParam As PTR, _
                 ByRef lParam As CWPSTRUCT) As Long
    
    If lCode = HC_ACTION Then
        
        If m_lRequestMessage = 0 Then
            m_lRequestMessage = RegisterWindowMessage(WM_REQUEST_COTHREADID)
        End If
        
        If m_lRequestMessage Then
            If lParam.message = m_lRequestMessage Then
                ' // Subclass main VB6 window
                m_pfnPrevWndProc = SetWindowLongPtr(lParam.hWnd, GWLP_WNDPROC, AddressOf NewWndProc)
            End If
        End If
        
    End If
    
    HookProc = CallNextHookEx(0, lCode, wParam, VarPtr(lParam))
    
End Function

' // This is main vb6 replacement wnd proc
Private Function NewWndProc( _
                 ByVal hWnd As Handle, _
                 ByVal lMsg As Long, _
                 ByVal wParam As PTR, _
                 ByVal lParam As PTR) As PTR
         
    Select Case lMsg
    Case m_lRequestMessage
        
        ' // Get COM thread id. This will be returned to SendMessageTimeout in main process
        NewWndProc = CoGetCurrentProcess()
        SetWindowLongPtr hWnd, GWLP_WNDPROC, m_pfnPrevWndProc
        
    Case Else
        NewWndProc = CallWindowProc(m_pfnPrevWndProc, hWnd, lMsg, wParam, ByVal lParam)
    End Select
         
End Function

' // Search for COM-thread by specified ProgID
Private Function FindComThreadIdInRegistry( _
                 ByRef sProgId As String, _
                 ByRef pComThreadIdsList As PTR, _
                 ByRef lFound As Long) As Long
    Dim hVBKeySave5 As Handle
    Dim hProgIdKey  As Handle
    Dim lKeyIndex   As Long
    Dim hr          As Long
    Dim sName       As String
    Dim sProgIDPath As String
    Dim lStatus     As Long
    Dim lCoIndex    As Long
    
    pComThreadIdsList = NULL_PTR
    lFound = 0
    
    lStatus = RegOpenKeyEx(HKEY_CLASSES_ROOT, "VBKeySave5", 0, KEY_ENUMERATE_SUB_KEYS Or KEY_QUERY_VALUE Or KEY_READ, hVBKeySave5)
    If lStatus Then
        hr = HRESULTFromWin32(lStatus)
        GoTo CleanUp
    End If
    
    sName = Space$(255)
    sProgIDPath = Space$(64)
    
    swprintf_s sProgIDPath, Len(sProgIDPath), "\DeletePI\%s", ByVal StrPtr(sProgId)
    
    hr = REGDB_E_KEYMISSING
    
    Do
        
        If RegEnumKeyEx(hVBKeySave5, lKeyIndex, sName, Len(sName) + 1, NULL_PTR, vbNullString, ByVal NULL_PTR, ByVal NULL_PTR) Then
            Exit Do
        End If
        
        lstrcat sName, sProgIDPath
        
        If RegOpenKeyEx(hVBKeySave5, sName, 0, KEY_READ, hProgIdKey) = 0 Then
            
            ' // Found
            RegCloseKey hProgIdKey
            
            pComThreadIdsList = CoTaskMemRealloc(pComThreadIdsList, (lCoIndex + 1) * 4)
            If pComThreadIdsList = NULL_PTR Then
                hr = E_OUTOFMEMORY
                Exit Do
            End If
            
            If swscanf_s(sName, "%08lx.S%08lx", ByVal pComThreadIdsList + lCoIndex * 4, 0&) = 2 Then
                hr = 0
                lCoIndex = lCoIndex + 1
            End If
            
        End If
        
        lKeyIndex = lKeyIndex + 1
        
    Loop
    
CleanUp:
    
    If hVBKeySave5 Then
        RegCloseKey hVBKeySave5
    End If
    
    If hr >= 0 Then
        lFound = lCoIndex
    ElseIf pComThreadIdsList Then
        CoTaskMemFree pComThreadIdsList
        pComThreadIdsList = NULL_PTR
    End If
    
    FindComThreadIdInRegistry = hr
    
End Function

