Attribute VB_Name = "modVSTRemote"
' //
' // modVSTRemote.bas - module for managing remote things
' // by The trick, 2022
' //

Option Explicit

Private Const CALLBACK_WND_CLASS    As String = "VbDebugVst_callback_wnd_class"
Private Const WM_CALLBACK           As Long = WM_USER

Private Declare Function InitializeDebugger Lib "VbDebugVstDll" ( _
                         ByVal hWndMain As Handle, _
                         ByVal hWndCallback As Handle) As Long
Private Declare Function UninitializeDebugger Lib "VbDebugVstDll" () As Long
Private Declare Function InitializeDebugee Lib "VbDebugVstDll" ( _
                         ByRef pProgId As PTR) As Long
Private Declare Function CreatePluginInstance Lib "VbDebugVstDll" ( _
                         ByRef cObject As IVBVstEffect_dbg) As Long
Private Declare Function GetSharedData Lib "VbDebugVstDll" () As PTR
Private Declare Function DestroyPluginInstance Lib "VbDebugVstDll" () As Long
Private Declare Function IsServerAlive Lib "VbDebugVstDll" () As Long
Private Declare Function GetHostState Lib "VbDebugVstDll" () As Long

' // The data are shared between processes, so no need copy
Public g_fSamplesL()    As Single
Public g_fSamplesR()    As Single
Public g_tEvents()      As VstEvent
Public g_tAutomation()  As tAutomationRecord

Private m_tSharedData()     As SHARED_DATA
Private m_tSharedData_SA    As SAFEARRAY1D
Private m_tSamplesL_SA      As SAFEARRAY1D
Private m_tSamplesR_SA      As SAFEARRAY1D
Private m_tEvents_SA        As SAFEARRAY1D
Private m_tAutomation_SA    As SAFEARRAY1D
Private m_bIsInitialized    As Boolean
Private m_bSupportsEvents   As Boolean
Private m_pSharedMem        As PTR
Private m_hCallback         As Handle

Public Property Get EventsBufferSize() As Long
    EventsBufferSize = m_tSharedData(0).lEventsBufSize
End Property

Public Property Get SamplesBufferSize() As Long
    SamplesBufferSize = m_tSharedData(0).lSamplesBufSize
End Property

Public Property Get AutomationBufferSize() As Long
    AutomationBufferSize = m_tSharedData(0).lAutomationBufSize
End Property

Public Property Get VSTContainerHandle() As Handle
    VSTContainerHandle = m_tSharedData(0).hWndContainer
End Property

' // Initialize remote things
Public Function InitializeRemote( _
                ByVal hWndMain As Handle) As Boolean
    Dim pShared     As PTR
    Dim bIsInIde    As Boolean
    Dim hLib        As PTR
    Dim tWndCls     As WNDCLASSEX
    Dim hr          As Long
    Dim hWnd        As Handle
    Dim bClsReg     As Boolean
    
    If m_bIsInitialized Then
        InitializeRemote = True
        Exit Function
    End If
    
    Debug.Assert MakeTrue(bIsInIde)
    
    If bIsInIde Then
        hLib = LoadLibrary(App.Path & "\release\VbDebugVstDll.dll")
    Else
        hLib = LoadLibrary("VbDebugVstDll.dll")
    End If
    
    If hLib = NULL_PTR Then
        Exit Function
    End If
    
    tWndCls.cbSize = LenB(tWndCls)
    
    If GetClassInfoEx(App.hInstance, CALLBACK_WND_CLASS, tWndCls) = 0 Then
        
        With tWndCls
            .lpfnWndProc = FAR_PROC(AddressOf CallbackWndProc)
            .hInstance = App.hInstance
            .lpszClassName = StrPtr(CALLBACK_WND_CLASS)
        End With
        
        If RegisterClassEx(tWndCls) = 0 Then
            Log "RegisterClassEx failed 0x" & Hex$(GetLastError)
            GoTo CleanUp
        End If
        
        bClsReg = True
        
    ElseIf tWndCls.lpfnWndProc <> FAR_PROC(AddressOf CallbackWndProc) Then
        Log "Callback already registered with wrong address"
        GoTo CleanUp
    End If
    
    hWnd = CreateWindowEx(0, CALLBACK_WND_CLASS, vbNullString, 0, 0, 0, 0, 0, HWND_MESSAGE, 0, App.hInstance, ByVal NULL_PTR)
    If hWnd = 0 Then
        Log "CreateWindowEx failed 0x" & Hex$(GetLastError)
        GoTo CleanUp
    End If
    
    hr = InitializeDebugger(hWndMain, hWnd)
    If hr < 0 Then
        Log "InitializeDebugger failed 0x" & Hex$(hr)
        GoTo CleanUp
    End If
    
    pShared = GetSharedData ' // Increment lib counter
    
    If pShared = NULL_PTR Then
        Log "GetSharedData failed"
        GoTo CleanUp
    End If
    
    With m_tSharedData_SA
        .cDims = 1
        .fFeatures = FADF_AUTO
        .pvData = pShared
        .rgsabound(0).cElements = 1
    End With
    
    PutMemPtr ByVal ArrPtr(m_tSharedData), VarPtr(m_tSharedData_SA)
    
    With m_tSamplesL_SA
        .cbElements = 4
        .cDims = 1
        .fFeatures = FADF_AUTO
        .pvData = m_tSharedData(0).pSamplesBuf
        .rgsabound(0).cElements = m_tSharedData(0).lSamplesBufSize \ 2
    End With
    
    With m_tSamplesR_SA
        .cbElements = 4
        .cDims = 1
        .fFeatures = FADF_AUTO
        .pvData = m_tSamplesL_SA.pvData + m_tSamplesL_SA.rgsabound(0).cElements * 4
        .rgsabound(0).cElements = m_tSamplesL_SA.rgsabound(0).cElements
    End With
    
    With m_tEvents_SA
        .cbElements = 32
        .cDims = 1
        .fFeatures = FADF_AUTO
        .pvData = m_tSharedData(0).pEventsBuf
        .rgsabound(0).cElements = m_tSharedData(0).lEventsBufSize
    End With
    
    With m_tAutomation_SA
        .cbElements = 8
        .cDims = 1
        .fFeatures = FADF_AUTO
        .pvData = m_tSharedData(0).pAutomationBuf
        .rgsabound(0).cElements = m_tSharedData(0).lAutomationBufSize
    End With
    
    PutMemPtr ByVal ArrPtr(g_fSamplesL), VarPtr(m_tSamplesL_SA)
    PutMemPtr ByVal ArrPtr(g_fSamplesR), VarPtr(m_tSamplesR_SA)
    PutMemPtr ByVal ArrPtr(g_tEvents), VarPtr(m_tEvents_SA)
    PutMemPtr ByVal ArrPtr(g_tAutomation), VarPtr(m_tEvents_SA)
    
    m_bIsInitialized = True
    m_pSharedMem = pShared
    m_hCallback = hWnd
    
    InitializeRemote = True
    
CleanUp:
    
    If Not m_bIsInitialized Then
    
        If hWnd Then
            DestroyWindow hWnd
        End If
        
        If bClsReg Then
            UnregisterClass CALLBACK_WND_CLASS, App.hInstance
        End If
        
    End If
    
    If hLib Then
        If Not bIsInIde Then
            FreeLibrary hLib
        End If
    End If
    
End Function

Public Sub UninitializeRemote()
    
    If Not m_bIsInitialized Then
        Exit Sub
    End If
    
    PutMemPtr ByVal ArrPtr(g_tEvents), NULL_PTR
    PutMemPtr ByVal ArrPtr(g_tAutomation), NULL_PTR
    PutMemPtr ByVal ArrPtr(g_fSamplesL), NULL_PTR
    PutMemPtr ByVal ArrPtr(g_fSamplesR), NULL_PTR
    PutMemPtr ByVal ArrPtr(m_tSharedData), NULL_PTR
    
    DestroyWindow m_hCallback
    UnregisterClass CALLBACK_WND_CLASS, App.hInstance
    
    UninitializeDebugger
    
    m_bIsInitialized = False
    
End Sub

' // Create plugin object in remote process
Public Function CreatePluginRemote( _
                ByRef sProgId As String) As IVBVstEffect_dbg
    Dim hr      As Long
    Dim cObj    As IVBVstEffect_dbg
    
    hr = IsServerAlive()
    
    Log "IsServerAlive returns 0x" & Hex$(hr)
    
    If hr <> S_OK Then
        
        hr = InitializeDebugee(StrPtr(sProgId))
        If hr < 0 Then
            Log "InitializeDebugee failed 0x" & Hex$(hr)
            GoTo CleanUp
        End If
    
    End If

    hr = GetHostState
    
    If hr < 0 Then
        Log "GetHostState failed 0x" & Hex$(hr)
    ElseIf hr = 0 Then
        Log "Host has stopped"
        GoTo CleanUp
    End If
    
    hr = CreatePluginInstance(cObj)
    If hr < 0 Then
        Log "CreatePluginInstance failed 0x" & Hex$(hr)
        GoTo CleanUp
    End If
    
    hr = cObj.SupportsVSTEvents(m_bSupportsEvents)
    If hr < 0 Then
        Log "SupportsVSTEvents failed 0x" & Hex$(hr)
        GoTo CleanUp
    End If
    
    Set CreatePluginRemote = cObj
    
CleanUp:

End Function

Public Function GetRemoteHostState() As Long
    GetRemoteHostState = GetHostState
End Function

Public Function DestroyPluginRemote() As Long
    DestroyPluginRemote = DestroyPluginInstance
End Function

Public Function ProcessEventsRemote( _
                ByVal cPlugin As IVBVstEffect_dbg, _
                ByVal lStartIndex As Long, _
                ByVal dPosPPQ As Double, _
                ByVal dTempo As Double, _
                ByVal lEvents As Long) As Boolean
    Dim hr      As Long
    Dim tEvents As VstEvents
    Dim bRet    As Boolean
    
    If cPlugin Is Nothing Then
        Exit Function
    ElseIf lEvents <= 0 Then
        ProcessEventsRemote = True
        Exit Function
    End If
    
    With m_tSharedData(0).tCurTimeInfo
    
        .barStartPos = Int(dPosPPQ * 4)
        .ppqPos = dPosPPQ
        .Tempo = dTempo
        .SampleRate = SampleRate
        .Flags = kVstBarsValid Or kVstPpqPosValid Or kVstTempoValid
    
    End With
    
    tEvents.numEvents = lEvents
    tEvents.pEvents = VarPtr(g_tEvents(lStartIndex))
    
    hr = cPlugin.ProcessEvents(tEvents, bRet)
    If hr < 0 Then
        Log "ProcessEvents failed 0x" & Hex$(hr)
    Else
        ProcessEventsRemote = bRet
    End If
    
End Function

Public Function ProcessSamplesRemote( _
                ByVal cPlugin As IVBVstEffect_dbg, _
                ByVal lStartIndex As Long, _
                ByVal dPosPPQ As Double, _
                ByVal dTempo As Double, _
                ByVal lSamples As Long) As Boolean
    Static s_pInSamples(1)     As PTR
    Static s_pOutSamples(1)    As PTR

    Dim hr  As Long

    If cPlugin Is Nothing Then
        Exit Function
    End If
    
    With m_tSharedData(0).tCurTimeInfo
    
        .barStartPos = Int(dPosPPQ * 4)
        .ppqPos = dPosPPQ
        .Tempo = dTempo
        .SampleRate = SampleRate
        .Flags = kVstBarsValid Or kVstPpqPosValid Or kVstTempoValid
    
    End With
    
    s_pInSamples(0) = VarPtr(g_fSamplesL(lStartIndex))
    s_pInSamples(1) = VarPtr(g_fSamplesR(lStartIndex))
    s_pOutSamples(0) = VarPtr(g_fSamplesL(lStartIndex))
    s_pOutSamples(1) = VarPtr(g_fSamplesR(lStartIndex))
    
    hr = cPlugin.ProcessReplacing(VarPtr(s_pInSamples(0)), VarPtr(s_pOutSamples(0)), lSamples)
    If hr < 0 Then
        Log "ProcessReplacing failed 0x" & Hex$(hr)
    Else
        ProcessSamplesRemote = True
    End If
    
End Function

Private Function CallbackWndProc( _
                 ByVal hWnd As Handle, _
                 ByVal lMsg As Long, _
                 ByVal wParam As PTR, _
                 ByVal lParam As PTR) As PTR
                     
    Select Case lMsg
    Case WM_CALLBACK
        Select Case wParam
        Case 0: frmVSTSite.RequestClose
        End Select
    Case Else
        CallbackWndProc = DefWindowProc(hWnd, lMsg, wParam, ByVal lParam)
    End Select
                     
End Function

Private Function MakeTrue( _
                 ByRef bValue As Boolean) As Boolean
    MakeTrue = True
    bValue = True
End Function

