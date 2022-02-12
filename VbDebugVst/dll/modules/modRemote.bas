Attribute VB_Name = "modRemote"
' //
' // modRemote.bas - remote side procedures
' // By The trick, 2022
' //

Option Explicit

Private Const RECEIVER_WND_CLASS    As String = "VbDebugVstDll_receiver_wnd_class"
Private Const CONTAINER_WND_CLASS   As String = "VbDebugVstDll_container_wnd_class"

Public Const WM_EXIT_THREAD              As Long = WM_USER + 0
Public Const WM_CREATE_PLUGIN            As Long = WM_USER + 1
Public Const WM_DESTROY_PLUGIN           As Long = WM_USER + 2
Public Const WM_HOST_STATE               As Long = WM_USER + 3
Public Const WM_STARTPROCESS             As Long = WM_USER + 4
Public Const WM_STOPPROCESS              As Long = WM_USER + 5
Public Const WM_SETBYPASS                As Long = WM_USER + 6
Public Const WM_PROCESSEVENTS            As Long = WM_USER + 7
Public Const WM_GETTAILSIZE              As Long = WM_USER + 8
Public Const WM_VENDORSPECIFIC           As Long = WM_USER + 9
Public Const WM_VENDORVERSION            As Long = WM_USER + 10
Public Const WM_GETPRODUCTSTRING         As Long = WM_USER + 11
Public Const WM_GETVENDORSTRING          As Long = WM_USER + 12
Public Const WM_GETEFFECTNAME            As Long = WM_USER + 13
Public Const WM_GETPROGRAMNAMEINDEXED    As Long = WM_USER + 14
Public Const WM_CANBEAUTOMATED           As Long = WM_USER + 15
Public Const WM_SETCHUNK                 As Long = WM_USER + 16
Public Const WM_GETCHUNK                 As Long = WM_USER + 17
Public Const WM_SETBLOCKSIZE             As Long = WM_USER + 18
Public Const WM_GETPROGRAMNAME           As Long = WM_USER + 19
Public Const WM_SETPROGRAMNAME           As Long = WM_USER + 20
Public Const WM_GETPROGRAM               As Long = WM_USER + 21
Public Const WM_SETPROGRAM               As Long = WM_USER + 22
Public Const WM_CANDO                    As Long = WM_USER + 23
Public Const WM_EDITIDLE                 As Long = WM_USER + 24
Public Const WM_EDITCLOSE                As Long = WM_USER + 25
Public Const WM_EDITOPEN                 As Long = WM_USER + 26
Public Const WM_SETSAMPLERATE            As Long = WM_USER + 27
Public Const WM_MAINSCHANGED             As Long = WM_USER + 28
Public Const WM_GETPARAMNAME             As Long = WM_USER + 29
Public Const WM_GETPARAMLABEL            As Long = WM_USER + 30
Public Const WM_GETPARAMDISPLAY          As Long = WM_USER + 31
Public Const WM_GETPARAMETER             As Long = WM_USER + 32
Public Const WM_SETPARAMETER             As Long = WM_USER + 33
Public Const WM_PROCESSREPLACING         As Long = WM_USER + 34
Public Const WM_PROCESS                  As Long = WM_USER + 35
Public Const WM_GETPARAMETERSCOUNT       As Long = WM_USER + 36
Public Const WM_PARAMETERPROPERTIES      As Long = WM_USER + 37
Public Const WM_UNIQUEID                 As Long = WM_USER + 38
Public Const WM_PLUGCATEGORY             As Long = WM_USER + 39
Public Const WM_PLUGVERSION              As Long = WM_USER + 40
Public Const WM_VSTVERSION               As Long = WM_USER + 41
Public Const WM_CANMONO                  As Long = WM_USER + 42
Public Const WM_HASEDITOR                As Long = WM_USER + 43
Public Const WM_PROGRAMSARECHUNK         As Long = WM_USER + 44
Public Const WM_SUPPORTSVSTEVENTS        As Long = WM_USER + 45
Public Const WM_NUMOFPROGRAMS            As Long = WM_USER + 46
Public Const WM_COPYPROGRAM              As Long = WM_USER + 47
Public Const WM_EDITORRECT               As Long = WM_USER + 48
Public Const WM_NUMOFINPUTS              As Long = WM_USER + 49
Public Const WM_NUMOFOUTPUTS             As Long = WM_USER + 50

Public Const WM_CALLBACK                 As Long = WM_USER

Private Type tOLEVariant
    iVT         As Integer
    iReserved0  As Integer
    iReserved1  As Integer
    iReserved2  As Integer
    pData0      As PTR
    pData1      As PTR
End Type

Private Type tDispCallFuncArgData
    pArgs(7)    As PTR
    iTypes(7)   As Integer
    lVarArgs(7) As tOLEVariant
    lVarRet     As tOLEVariant
End Type

Private m_hWnd          As Handle           ' // Communicate handle
Private m_hContainer    As Handle
Private m_cObject       As IVBVstEffect_dbg

' // Initialize remote things
Public Function InitializeRemote() As Boolean
    Dim tWndClass   As WNDCLASSEX
    Dim bClsRegRecv As Boolean
    Dim bClsRegCont As Boolean
    Dim hWndRecv    As Handle
    Dim hWndCont    As Handle
    Dim hr          As Long

    tWndClass.cbSize = LenB(tWndClass)
    
    If GetClassInfoEx(g_hInstance, RECEIVER_WND_CLASS, tWndClass) = 0 Then
        
        With tWndClass
            .hInstance = g_hInstance
            .lpszClassName = StrPtr(RECEIVER_WND_CLASS)
            .lpfnWndProc = FAR_PROC(AddressOf ReceiverWndProc)
        End With
        
        If RegisterClassEx(tWndClass) = 0 Then
            hr = HRESULTFromWin32(GetLastError)
            GoTo CleanUp
        End If
        
        bClsRegRecv = True
        
    ElseIf tWndClass.lpfnWndProc <> FAR_PROC(AddressOf ReceiverWndProc) Then
        ' // Probably error
        hr = E_UNEXPECTED
        GoTo CleanUp
    End If
    
    hWndRecv = CreateWindowEx(0, RECEIVER_WND_CLASS, vbNullString, 0, 0, 0, 0, 0, HWND_MESSAGE, 0, g_hInstance, ByVal NULL_PTR)
    If hWndRecv = 0 Then
        hr = HRESULTFromWin32(GetLastError)
        GoTo CleanUp
    End If
    
    If GetClassInfoEx(g_hInstance, CONTAINER_WND_CLASS, tWndClass) = 0 Then
        
        With tWndClass
            .hInstance = g_hInstance
            .lpszClassName = StrPtr(CONTAINER_WND_CLASS)
            .lpfnWndProc = FAR_PROC(AddressOf ContainerWndProc)
        End With
        
        If RegisterClassEx(tWndClass) = 0 Then
            hr = HRESULTFromWin32(GetLastError)
            GoTo CleanUp
        End If
        
        bClsRegCont = True
        
    ElseIf tWndClass.lpfnWndProc <> FAR_PROC(AddressOf ContainerWndProc) Then
        hr = E_UNEXPECTED
        GoTo CleanUp
    End If
    
    hWndCont = CreateWindowEx(WS_EX_TOOLWINDOW, CONTAINER_WND_CLASS, vbNullString, WS_CLIPCHILDREN Or WS_CLIPSIBLINGS Or _
                              WS_OVERLAPPED Or WS_SYSMENU Or WS_CAPTION, 0, 0, 0, 0, 0, 0, g_hInstance, ByVal NULL_PTR)
    If hWndCont = 0 Then
        hr = HRESULTFromWin32(GetLastError)
        GoTo CleanUp
    End If
    
    With g_tSharedData(0)
    
        .hWnd = hWndRecv
        .hWndContainer = hWndCont
        .hDllRemote = g_hInstance
        .lVstProcessId = GetCurrentProcessId
        .lVstThreadId = GetCurrentThreadId
        .pEventsBufRemote = g_pSharedData + LenB(g_tSharedData(0))
        .pSamplesBufRemote = .pEventsBufRemote + EVENTS_BUFFER_SIZE * 32
        .pAutomationBufRemote = .pSamplesBufRemote + SAMPLES_BUFFER_SIZE * 4
        .pDataBufferRemote = .pAutomationBufRemote + AUTOMATION_BUFFER_SIZE * 8
        .pfnHostCallbackRemote = FAR_PROC(AddressOf AudioMasterCallback)
        
    End With
    
    m_hWnd = hWndRecv
    m_hContainer = hWndCont
    
    InitializeRemote = True
    
CleanUp:
    
    If hr < 0 Then
        
        If hWndRecv Then
            DestroyWindow hWndRecv
        End If
        
        If hWndCont Then
            DestroyWindow hWndCont
        End If
        
        If bClsRegRecv Then
            UnregisterClass RECEIVER_WND_CLASS, g_hInstance
        End If
        
        If bClsRegCont Then
            UnregisterClass CONTAINER_WND_CLASS, g_hInstance
        End If
        
    End If
    
    g_tSharedData(0).hr = hr
    
End Function

' // Only get time info implemented
Private Function AudioMasterCallback CDecl( _
                 ByVal pAEffect As PTR, _
                 ByVal lOpcode As Long, _
                 ByVal lIndex As Long, _
                 ByVal lValue As PTR, _
                 ByVal lptr As PTR, _
                 ByVal fValue As Single) As PTR
         
    Select Case lOpcode
    Case audioMasterGetTime
        AudioMasterCallback = VarPtr(g_tSharedData(0).tCurTimeInfo)
    End Select
         
End Function

' // Container window proc
Private Function ContainerWndProc( _
                 ByVal hWnd As Handle, _
                 ByVal lMsg As Long, _
                 ByVal wParam As PTR, _
                 ByVal lParam As PTR) As PTR
    Dim tPS As PAINTSTRUCT
    
    Select Case lMsg
    Case WM_NCCREATE
        ContainerWndProc = 1
    Case WM_PAINT
        BeginPaint hWnd, tPS
        EndPaint hWnd, tPS
    Case WM_CLOSE
        If MessageBox(hWnd, "Are you sure?", "VbDebugVst", vbQuestion Or vbYesNo) = vbYes Then
            PostMessage g_tSharedData(0).hWndCallback, WM_CALLBACK, 0, ByVal NULL_PTR
        End If
    Case Else
        ContainerWndProc = DefWindowProc(hWnd, lMsg, wParam, ByVal lParam)
    End Select
    
End Function

Private Function ReceiverWndProc( _
                 ByVal hWnd As Handle, _
                 ByVal lMsg As Long, _
                 ByVal wParam As PTR, _
                 ByVal lParam As PTR) As PTR
    
    Select Case lMsg
    Case WM_EXIT_THREAD
        ReqExitThread
    Case WM_CREATE_PLUGIN
        ReceiverWndProc = ReqCreatePlugin()
    Case WM_DESTROY_PLUGIN
        ReceiverWndProc = ReqDestroyPlugin()
    Case WM_HOST_STATE
        ReceiverWndProc = ReqHostState()
    Case WM_STARTPROCESS
        ReceiverWndProc = ReqStartProcess()
    Case WM_STOPPROCESS
        ReceiverWndProc = ReqStopProcess()
    Case WM_SETBYPASS
        ReceiverWndProc = ReqSetBypass(wParam)
    Case WM_PROCESSEVENTS
        ReceiverWndProc = ReqProcessEvents()
    Case WM_GETTAILSIZE
        ReceiverWndProc = ReqGetTailSize()
    Case WM_VENDORSPECIFIC
        ReceiverWndProc = ReqVendorSpecific()
    Case WM_VENDORVERSION
        ReceiverWndProc = ReqVendorVersion()
    Case WM_GETPRODUCTSTRING
        ReceiverWndProc = ReqGetProductString()
    Case WM_GETVENDORSTRING
        ReceiverWndProc = ReqGetVendorString()
    Case WM_GETEFFECTNAME
        ReceiverWndProc = ReqGetEffectName()
    Case WM_GETPROGRAMNAMEINDEXED
        ReceiverWndProc = ReqGetProgramNameIndexed(wParam, lParam)
    Case WM_CANBEAUTOMATED
        ReceiverWndProc = ReqCanBeAutomated(wParam)
    Case WM_SETCHUNK
        ReceiverWndProc = ReqSetChunk(wParam, lParam)
    Case WM_GETCHUNK
        ReceiverWndProc = ReqGetChunk(wParam)
    Case WM_SETBLOCKSIZE
        ReceiverWndProc = ReqSetBlockSize(wParam)
    Case WM_GETPROGRAMNAME
        ReceiverWndProc = ReqGetProgramName()
    Case WM_SETPROGRAMNAME
        ReceiverWndProc = ReqSetProgramName()
    Case WM_GETPROGRAM
        ReceiverWndProc = ReqGetProgram()
    Case WM_SETPROGRAM
        ReceiverWndProc = ReqSetProgram(wParam)
    Case WM_CANDO
        ReceiverWndProc = ReqCanDo()
    Case WM_EDITIDLE
        ReceiverWndProc = ReqEditIdle()
    Case WM_EDITCLOSE
        ReceiverWndProc = ReqEditClose()
    Case WM_EDITOPEN
        ReceiverWndProc = ReqEditOpen(wParam)
    Case WM_SETSAMPLERATE
        ReceiverWndProc = ReqSetSampleRate()
    Case WM_MAINSCHANGED
        ReceiverWndProc = ReqMainsChanged(wParam)
    Case WM_GETPARAMNAME
        ReceiverWndProc = ReqGetParamName(wParam)
    Case WM_GETPARAMLABEL
        ReceiverWndProc = ReqGetParamLabel(wParam)
    Case WM_GETPARAMDISPLAY
        ReceiverWndProc = ReqGetParamDisplay(wParam)
    Case WM_GETPARAMETER
        ReceiverWndProc = ReqGetParameter(wParam)
    Case WM_SETPARAMETER
        ReceiverWndProc = ReqSetParameter(wParam)
    Case WM_PROCESSREPLACING
        ReceiverWndProc = ReqProcessReplacing()
    Case WM_PROCESS
        ReceiverWndProc = ReqProcess()
    Case WM_GETPARAMETERSCOUNT
        ReceiverWndProc = ReqGetParametersCount()
    Case WM_PARAMETERPROPERTIES
        ReceiverWndProc = ReqGetParameterProperties()
    Case WM_UNIQUEID
        ReceiverWndProc = ReqUniqueId()
    Case WM_PLUGCATEGORY
        ReceiverWndProc = ReqPlugCategory()
    Case WM_PLUGVERSION
        ReceiverWndProc = ReqVersion()
    Case WM_VSTVERSION
        ReceiverWndProc = ReqVstVersion()
    Case WM_CANMONO
        ReceiverWndProc = ReqCanMono()
    Case WM_HASEDITOR
        ReceiverWndProc = ReqHasEditor()
    Case WM_PROGRAMSARECHUNK
        ReceiverWndProc = ReqProgramsAreChunk()
    Case WM_SUPPORTSVSTEVENTS
        ReceiverWndProc = ReqSupportsVstEvents()
    Case WM_NUMOFPROGRAMS
        ReceiverWndProc = ReqNumOfPrograms()
    Case WM_COPYPROGRAM
        ReceiverWndProc = ReqCopyProgram(wParam)
    Case WM_EDITORRECT
        ReceiverWndProc = ReqEditorRect()
    Case WM_NUMOFINPUTS
        ReceiverWndProc = ReqNumOfInputs()
    Case Else
        ReceiverWndProc = DefWindowProc(hWnd, lMsg, wParam, ByVal lParam)
    End Select
    
End Function

Private Function ReqNumOfOutputs() As Long
    If m_cObject Is Nothing Then
        ReqNumOfOutputs = E_UNEXPECTED
    Else
        ReqNumOfOutputs = m_cObject.NumOfOutputs(g_tSharedData(0).lArg1)
    End If
End Function

Private Function ReqNumOfInputs() As Long
    If m_cObject Is Nothing Then
        ReqNumOfInputs = E_UNEXPECTED
    Else
        ReqNumOfInputs = m_cObject.NumOfInputs(g_tSharedData(0).lArg1)
    End If
End Function

Private Function ReqEditorRect() As Long
    Dim tRC As ERect
    
    If m_cObject Is Nothing Then
        ReqEditorRect = E_UNEXPECTED
    Else
        ReqEditorRect = m_cObject.EditorRect(tRC)
        memcpy ByVal g_tSharedData(0).pDataBufferRemote, tRC, Len(tRC)
    End If
    
End Function

Private Function ReqCopyProgram( _
                 ByVal lIndex As Long) As Long
    If m_cObject Is Nothing Then
        ReqCopyProgram = E_UNEXPECTED
    Else
        ReqCopyProgram = m_cObject.CopyProgram(lIndex)
    End If
End Function

Private Function ReqNumOfPrograms() As Long
    If m_cObject Is Nothing Then
        ReqNumOfPrograms = E_UNEXPECTED
    Else
        ReqNumOfPrograms = m_cObject.NumOfPrograms(g_tSharedData(0).lArg1)
    End If
End Function

Private Function ReqSupportsVstEvents() As Long
    Dim bRet    As Boolean
    
    If m_cObject Is Nothing Then
        ReqSupportsVstEvents = E_UNEXPECTED
    Else
        ReqSupportsVstEvents = m_cObject.SupportsVSTEvents(bRet)
        g_tSharedData(0).lArg1 = bRet
    End If
    
End Function

Private Function ReqProgramsAreChunk() As Long
    Dim bRet    As Boolean
    
    If m_cObject Is Nothing Then
        ReqProgramsAreChunk = E_UNEXPECTED
    Else
        ReqProgramsAreChunk = m_cObject.ProgramsAreChunks(bRet)
        g_tSharedData(0).lArg1 = bRet
    End If
    
End Function

Private Function ReqHasEditor() As Long
    Dim bRet    As Boolean
    
    If m_cObject Is Nothing Then
        ReqHasEditor = E_UNEXPECTED
    Else
        ReqHasEditor = m_cObject.HasEditor(bRet)
        g_tSharedData(0).lArg1 = bRet
    End If
    
End Function

Private Function ReqCanMono() As Long
    Dim bRet    As Boolean
    
    If m_cObject Is Nothing Then
        ReqCanMono = E_UNEXPECTED
    Else
        ReqCanMono = m_cObject.CanMono(bRet)
        g_tSharedData(0).lArg1 = bRet
    End If
    
End Function

Private Function ReqVstVersion() As Long
    If m_cObject Is Nothing Then
        ReqVstVersion = E_UNEXPECTED
    Else
        ReqVstVersion = m_cObject.VstVersion(g_tSharedData(0).lArg1)
    End If
End Function

Private Function ReqVersion() As Long
    If m_cObject Is Nothing Then
        ReqVersion = E_UNEXPECTED
    Else
        ReqVersion = m_cObject.Version(g_tSharedData(0).lArg1)
    End If
End Function

Private Function ReqPlugCategory() As Long
    If m_cObject Is Nothing Then
        ReqPlugCategory = E_UNEXPECTED
    Else
        ReqPlugCategory = m_cObject.PlugCategory(g_tSharedData(0).lArg1)
    End If
End Function

Private Function ReqUniqueId() As Long
    If m_cObject Is Nothing Then
        ReqUniqueId = E_UNEXPECTED
    Else
        ReqUniqueId = m_cObject.UniqueId(g_tSharedData(0).lArg1)
    End If
End Function

Private Function ReqGetParameterProperties() As Long
    Dim bRet    As Boolean
    
    If m_cObject Is Nothing Then
        ReqGetParameterProperties = E_UNEXPECTED
    Else
        ReqGetParameterProperties = m_cObject.ParameterProperties(g_tSharedData(0).pArg1, bRet)
        g_tSharedData(0).lArg1 = bRet
    End If
    
End Function

Private Function ReqGetParametersCount() As Long
    If m_cObject Is Nothing Then
        ReqGetParametersCount = E_UNEXPECTED
    Else
        ReqGetParametersCount = m_cObject.NumOfParam(g_tSharedData(0).lArg1)
    End If
End Function

Private Function ReqProcess() As Long
    If m_cObject Is Nothing Then
        ReqProcess = E_UNEXPECTED
    Else
        ReqProcess = m_cObject.Process(g_tSharedData(0).pArg1, g_tSharedData(0).pArg2, g_tSharedData(0).lArg1)
    End If
End Function

Private Function ReqProcessReplacing() As Long
    If m_cObject Is Nothing Then
        ReqProcessReplacing = E_UNEXPECTED
    Else
        ReqProcessReplacing = m_cObject.ProcessReplacing(g_tSharedData(0).pArg1, g_tSharedData(0).pArg2, g_tSharedData(0).lArg1)
    End If
End Function

Private Function ReqSetParameter( _
                 ByVal lIndex As Long) As Long
    If m_cObject Is Nothing Then
        ReqSetParameter = E_UNEXPECTED
    Else
        ReqSetParameter = m_cObject.ParamValue_put(lIndex, g_tSharedData(0).fArg1)
    End If
End Function

Private Function ReqGetParameter( _
                 ByVal lIndex As Long) As Long
    If m_cObject Is Nothing Then
        ReqGetParameter = E_UNEXPECTED
    Else
        ReqGetParameter = m_cObject.ParamValue_get(lIndex, g_tSharedData(0).fArg1)
    End If
End Function

Private Function ReqGetParamDisplay( _
                 ByVal lIndex As Long) As Long
    Dim sRet    As String
    
    If m_cObject Is Nothing Then
        ReqGetParamDisplay = E_UNEXPECTED
    Else
    
        ReqGetParamDisplay = m_cObject.ParamDisplay(lIndex, sRet)
        
        If Len(sRet) Then
            lstrcpyn ByVal g_tSharedData(0).pDataBufferRemote, ByVal StrPtr(sRet), g_tSharedData(0).lDataBufferSize \ 2
        Else
            PutMem2 ByVal g_tSharedData(0).pDataBufferRemote, 0
        End If
        
    End If
    
End Function

Private Function ReqGetParamLabel( _
                 ByVal lIndex As Long) As Long
    Dim sRet    As String
    
    If m_cObject Is Nothing Then
        ReqGetParamLabel = E_UNEXPECTED
    Else
    
        ReqGetParamLabel = m_cObject.ParamLabel(lIndex, sRet)
        
        If Len(sRet) Then
            lstrcpyn ByVal g_tSharedData(0).pDataBufferRemote, ByVal StrPtr(sRet), g_tSharedData(0).lDataBufferSize \ 2
        Else
            PutMem2 ByVal g_tSharedData(0).pDataBufferRemote, 0
        End If
        
    End If
    
End Function

Private Function ReqGetParamName( _
                 ByVal lIndex As Long) As Long
    Dim sRet    As String
    
    If m_cObject Is Nothing Then
        ReqGetParamName = E_UNEXPECTED
    Else
    
        ReqGetParamName = m_cObject.ParamName(lIndex, sRet)
        
        If Len(sRet) Then
            lstrcpyn ByVal g_tSharedData(0).pDataBufferRemote, ByVal StrPtr(sRet), g_tSharedData(0).lDataBufferSize \ 2
        Else
            PutMem2 ByVal g_tSharedData(0).pDataBufferRemote, 0
        End If
        
    End If
    
End Function

Private Function ReqMainsChanged( _
                 ByVal lValue As Long) As Long
    If m_cObject Is Nothing Then
        ReqMainsChanged = E_UNEXPECTED
    Else
        If lValue Then
            ReqMainsChanged = m_cObject.Resume()
        Else
            ReqMainsChanged = m_cObject.Suspend()
        End If
    End If
End Function

Private Function ReqSetSampleRate() As Long
    If m_cObject Is Nothing Then
        ReqSetSampleRate = E_UNEXPECTED
    Else
        ReqSetSampleRate = m_cObject.SampleRate_put(g_tSharedData(0).fArg1)
    End If
End Function

Private Function ReqEditOpen( _
                 ByVal hWnd As Handle) As Long
    Dim bRet    As Boolean
    Dim tRC     As ERect
    Dim tWndRc  As RECT
    Dim hr      As Long
    
    ' // We can't use remote hWnd because the code is blocked when caller wait response. So input queue is blocked.
    If m_cObject Is Nothing Then
        hr = E_UNEXPECTED
        GoTo exit_proc
    End If

    hr = m_cObject.EditorRect(tRC)
    If hr < 0 Then
        GoTo exit_proc
    End If
    
    SetRect tWndRc, 0, 0, tRC.wLeft + tRC.wRight, tRC.wTop + tRC.wBottom
    
    If AdjustWindowRectEx(tWndRc, GetWindowLongPtr(m_hContainer, GWL_STYLE), 0, _
                          GetWindowLongPtr(m_hContainer, GWL_EXSTYLE)) = 0 Then
        hr = HRESULTFromWin32(GetLastError)
        GoTo exit_proc
    End If
    
    OffsetRect tWndRc, (GetSystemMetrics(SM_CXSCREEN) - (tWndRc.Right - tWndRc.Left)) \ 2, _
                       (GetSystemMetrics(SM_CYSCREEN) - (tWndRc.Bottom - tWndRc.Top)) \ 2
    
    If SetWindowPos(m_hContainer, 0, tWndRc.Left, tWndRc.Top, tWndRc.Right - tWndRc.Left, _
                    tWndRc.Bottom - tWndRc.Top, SWP_SHOWWINDOW Or SWP_NOZORDER) = 0 Then
        hr = HRESULTFromWin32(GetLastError)
        GoTo exit_proc
    End If

    hr = m_cObject.EditorOpen(m_hContainer, bRet)
    g_tSharedData(0).lArg1 = bRet

exit_proc:
    
    ReqEditOpen = hr
    
End Function

Private Function ReqEditClose() As Long
    If m_cObject Is Nothing Then
        ReqEditClose = E_UNEXPECTED
    Else
        ReqEditClose = m_cObject.EditorClose
        ShowWindow m_hContainer, SW_HIDE
    End If
End Function

Private Function ReqEditIdle() As Long
    Dim tBuf()  As tAutomationRecord
    Dim lCount  As Long
    
    If m_cObject Is Nothing Then
        ReqEditIdle = E_UNEXPECTED
    Else

        ReqEditIdle = m_cObject.EditorIdle(tBuf, lCount)
        
        If lCount > g_tSharedData(0).lAutomationBufSize Then
            lCount = g_tSharedData(0).lAutomationBufSize
        End If
        
        If lCount > 0 Then
            memcpy ByVal g_tSharedData(0).pAutomationBufRemote, tBuf(0), lCount
        End If
        
        g_tSharedData(0).lAutomationCount = lCount
        
    End If
    
End Function

Private Function ReqCanDo() As Long
    Dim bRet    As Boolean
    
    If m_cObject Is Nothing Then
        ReqCanDo = E_UNEXPECTED
    Else
        ReqCanDo = m_cObject.CanDo(SysAllocString(ByVal g_tSharedData(0).pDataBufferRemote), bRet)
        g_tSharedData(0).lArg1 = bRet
    End If
    
End Function

Private Function ReqSetProgram( _
                 ByVal lIndex As Long) As Long
    If m_cObject Is Nothing Then
        ReqSetProgram = E_UNEXPECTED
    Else
        ReqSetProgram = m_cObject.Program_put(lIndex)
    End If
End Function

Private Function ReqGetProgram() As Long
    If m_cObject Is Nothing Then
        ReqGetProgram = E_UNEXPECTED
    Else
        ReqGetProgram = m_cObject.Program_get(g_tSharedData(0).lArg1)
    End If
End Function

Private Function ReqSetProgramName() As Long
    If m_cObject Is Nothing Then
        ReqSetProgramName = E_UNEXPECTED
    Else
        ReqSetProgramName = m_cObject.ProgramName_put(SysAllocString(ByVal g_tSharedData(0).pDataBufferRemote))
    End If
End Function

Private Function ReqGetProgramName() As Long
    Dim sRet    As String
    
    If m_cObject Is Nothing Then
        ReqGetProgramName = E_UNEXPECTED
    Else
        ReqGetProgramName = m_cObject.ProgramName_get(sRet)
        lstrcpyn ByVal g_tSharedData(0).pDataBufferRemote, ByVal StrPtr(sRet), g_tSharedData(0).lDataBufferSize \ 2
    End If
    
End Function

Private Function ReqSetBlockSize( _
                 ByVal lSize As Long) As Long
    If m_cObject Is Nothing Then
        ReqSetBlockSize = E_UNEXPECTED
    Else
        ReqSetBlockSize = m_cObject.BlockSize(lSize)
    End If
End Function

Private Function ReqGetChunk( _
                 ByVal lIsPreset As Long) As Long
    If m_cObject Is Nothing Then
        ReqGetChunk = E_UNEXPECTED
    Else
        ReqGetChunk = m_cObject.GetStateChunk(lIsPreset, g_tSharedData(0).pArg1, g_tSharedData(0).lArg1)
    End If
End Function

Private Function ReqSetChunk( _
                 ByVal lIsPreset As Long, _
                 ByVal lSize As Long) As Long
    Dim bRet    As Boolean

    If m_cObject Is Nothing Then
        ReqSetChunk = E_UNEXPECTED
    Else
        ReqSetChunk = m_cObject.SetStateChunk(lIsPreset, g_tSharedData(0).pArg1, lSize, bRet)
        g_tSharedData(0).lArg1 = bRet
    End If
    
End Function

Private Function ReqCanBeAutomated( _
                 ByVal lIndex As Long) As Long
    Dim bRet    As Boolean
    
    If m_cObject Is Nothing Then
        ReqCanBeAutomated = E_UNEXPECTED
    Else
        ReqCanBeAutomated = m_cObject.CanParameterBeAutomated(lIndex, bRet)
        g_tSharedData(0).lArg1 = bRet
    End If
    
End Function

Private Function ReqGetProgramNameIndexed( _
                 ByVal lCategory As Long, _
                 ByVal lIndex As Long) As Long
    Dim sRet    As String
    
    If m_cObject Is Nothing Then
        ReqGetProgramNameIndexed = E_UNEXPECTED
    Else
    
        ReqGetProgramNameIndexed = m_cObject.ProgramNameIndexed(lCategory, lIndex, sRet)
        
        If Len(sRet) Then
            lstrcpyn ByVal g_tSharedData(0).pDataBufferRemote, ByVal StrPtr(sRet), g_tSharedData(0).lDataBufferSize \ 2
        Else
            PutMem2 ByVal g_tSharedData(0).pDataBufferRemote, 0
        End If
        
    End If
    
End Function

Private Function ReqGetEffectName() As Long
    Dim sRet    As String
    
    If m_cObject Is Nothing Then
        ReqGetEffectName = E_UNEXPECTED
    Else
    
        ReqGetEffectName = m_cObject.EffectName(sRet)
        
        If Len(sRet) Then
            lstrcpyn ByVal g_tSharedData(0).pDataBufferRemote, ByVal StrPtr(sRet), g_tSharedData(0).lDataBufferSize \ 2
        Else
            PutMem2 ByVal g_tSharedData(0).pDataBufferRemote, 0
        End If
        
    End If
    
End Function

Private Function ReqGetVendorString() As Long
    Dim sRet    As String
    
    If m_cObject Is Nothing Then
        ReqGetVendorString = E_UNEXPECTED
    Else
    
        ReqGetVendorString = m_cObject.VendorString(sRet)
        
        If Len(sRet) Then
            lstrcpyn ByVal g_tSharedData(0).pDataBufferRemote, ByVal StrPtr(sRet), g_tSharedData(0).lDataBufferSize \ 2
        Else
            PutMem2 ByVal g_tSharedData(0).pDataBufferRemote, 0
        End If
        
    End If
    
End Function

Private Function ReqGetProductString() As Long
    Dim sRet    As String
    
    If m_cObject Is Nothing Then
        ReqGetProductString = E_UNEXPECTED
    Else
    
        ReqGetProductString = m_cObject.ProductString(sRet)
        
        If Len(sRet) Then
            lstrcpyn ByVal g_tSharedData(0).pDataBufferRemote, ByVal StrPtr(sRet), g_tSharedData(0).lDataBufferSize \ 2
        Else
            PutMem2 ByVal g_tSharedData(0).pDataBufferRemote, 0
        End If
        
    End If
    
End Function

Private Function ReqVendorVersion() As Long
    Dim lRet    As Long
    
    If m_cObject Is Nothing Then
        ReqVendorVersion = E_UNEXPECTED
    Else
        ReqVendorVersion = m_cObject.VendorVersion(lRet)
        g_tSharedData(0).lArg1 = lRet
    End If
    
End Function

Private Function ReqVendorSpecific() As Long
    If m_cObject Is Nothing Then
        ReqVendorSpecific = E_UNEXPECTED
    Else
        With g_tSharedData(0)
            ReqVendorSpecific = m_cObject.VendorSpecific(.lArg1, .lArg2, .pArg1, .fArg1)
        End With
    End If
End Function

Private Function ReqGetTailSize() As Long
    Dim lRet    As Long
    
    If m_cObject Is Nothing Then
        ReqGetTailSize = E_UNEXPECTED
    Else
        ReqGetTailSize = m_cObject.TailSize(lRet)
        g_tSharedData(0).lArg1 = lRet
    End If
    
End Function

Private Function ReqProcessEvents() As Long
    Dim bRet    As Boolean
    Dim tEvents As VstEvents

    If m_cObject Is Nothing Then
        ReqProcessEvents = E_UNEXPECTED
    Else
        
        tEvents.numEvents = g_tSharedData(0).lArg1
        tEvents.pEvents = g_tSharedData(0).pArg1
        
        ReqProcessEvents = m_cObject.ProcessEvents(tEvents, bRet)
        g_tSharedData(0).lArg1 = bRet
        
    End If
    
End Function

Private Function ReqSetBypass( _
                 ByVal lValue As Long) As Long
    Dim bRet    As Boolean
    
    If m_cObject Is Nothing Then
        ReqSetBypass = E_UNEXPECTED
    Else
        ReqSetBypass = m_cObject.SetBypass(lValue, bRet)
        g_tSharedData(0).lArg1 = bRet
    End If
    
End Function

Private Function ReqStopProcess() As Long
    Dim bRet    As Boolean
    
    If m_cObject Is Nothing Then
        ReqStopProcess = E_UNEXPECTED
    Else
        ReqStopProcess = m_cObject.StopProcess(bRet)
        g_tSharedData(0).lArg1 = bRet
    End If
    
End Function

Private Function ReqStartProcess() As Long
    Dim bRet    As Boolean
    
    If m_cObject Is Nothing Then
        ReqStartProcess = E_UNEXPECTED
    Else
        ReqStartProcess = m_cObject.StartProcess(bRet)
        g_tSharedData(0).lArg1 = bRet
    End If
    
End Function

Private Function ReqHostState() As Long
    Dim pfn     As PTR
    Dim tRet    As tOLEVariant
    Dim hr      As Long
    
    pfn = GetProcAddress(GetModuleHandle("vba6"), "EbMode")
    If pfn = 0 Then
        ReqHostState = HRESULTFromWin32(GetLastError)
        Exit Function
    End If
    
    hr = DispCallFunc(ByVal NULL_PTR, pfn, CC_STDCALL, vbLong, 0, ByVal NULL_PTR, ByVal NULL_PTR, tRet)
    If hr < 0 Then
        ReqHostState = hr
        Exit Function
    End If
    
    ReqHostState = tRet.pData0
    
End Function

Private Function ReqDestroyPlugin() As Long
    Set m_cObject = Nothing
    ShowWindow m_hContainer, SW_HIDE
End Function

Private Function ReqCreatePlugin() As Long
    Dim tClsId  As UUID
    Dim pObj    As PTR
    Dim hr      As Long
    
    hr = CLSIDFromProgID(g_tSharedData(0).pszProgId(0), tClsId)
    If hr < 0 Then
        ReqCreatePlugin = hr
        Exit Function
    End If
    
    hr = CoCreateInstance(tClsId, ByVal NULL_PTR, CLSCTX_INPROC_SERVER, IID_IVBVstEffect, pObj)
    If hr < 0 Then
        ReqCreatePlugin = hr
        Exit Function
    End If
    
    vbaObjSet m_cObject, ByVal pObj
    
    m_cObject.AudioMasterCallback_put AddressOf AudioMasterCallback
    
End Function

Private Sub ReqExitThread()

    DestroyWindow m_hWnd
    UnregisterClass RECEIVER_WND_CLASS, g_hInstance
    g_tSharedData(0).hWnd = 0
    Set m_cObject = Nothing
    m_hWnd = 0
    DestroyWindow m_hContainer
    UnregisterClass CONTAINER_WND_CLASS, g_hInstance
    m_hContainer = 0
    
End Sub


