Attribute VB_Name = "modVBVst2X"
' //
' // modVBVst2X.bas - VB6 Vst framework.
' // by The trick, 2022
' //
' // It doesn't support all the features of VST 2.4 yet
' // To use this module you should:
' // - specify ActiveX DLL project type with the single apartment threading;
' // - install VBCDeclFix Add-in (https://github.com/thetrik/VBCDeclFix);
' // - add vbvst2x.tlb to the Project->References;
' // - add a public creatable class. Add global constant VST_PLUGIN_CLASS_NAME with its name.
' //   You should impement IVBVstEffect interface in this class;
' // - export VSTPluginMain function with names 'VSTPluginMain' and 'main' using the following link switches:
' //   -EXPORT:VSTPluginMain -EXPORT:main=VSTPluginMain http://bbs.vbstreets.ru/viewtopic.php?f=9&t=43618 (VBCompiler/LinkSwitches);
' // - compile project to native code with all the optimizations enabled.
' //

Option Explicit

Private Const WM_CREATEPLUGIN             As Long = WM_USER + 1
Private Const WM_EXITTHREAD               As Long = WM_USER + 2
Private Const WM_RELEASEREF               As Long = WM_USER + 3
Private Const WM_STARTPROCESS             As Long = WM_USER + 4
Private Const WM_STOPPROCESS              As Long = WM_USER + 5
Private Const WM_SETBYPASS                As Long = WM_USER + 6
Private Const WM_PROCESSEVENTS            As Long = WM_USER + 7
Private Const WM_GETTAILSIZE              As Long = WM_USER + 8
Private Const WM_VENDORSPECIFIC           As Long = WM_USER + 9
Private Const WM_VENDORVERSION            As Long = WM_USER + 10
Private Const WM_GETPRODUCTSTRING         As Long = WM_USER + 11
Private Const WM_GETVENDORSTRING          As Long = WM_USER + 12
Private Const WM_GETEFFECTNAME            As Long = WM_USER + 13
Private Const WM_GETPROGRAMNAMEINDEXED    As Long = WM_USER + 14
Private Const WM_CANBEAUTOMATED           As Long = WM_USER + 15
Private Const WM_SETCHUNK                 As Long = WM_USER + 16
Private Const WM_GETCHUNK                 As Long = WM_USER + 17
Private Const WM_SETBLOCKSIZE             As Long = WM_USER + 18
Private Const WM_GETPROGRAMNAME           As Long = WM_USER + 19
Private Const WM_SETPROGRAMNAME           As Long = WM_USER + 20
Private Const WM_GETPROGRAM               As Long = WM_USER + 21
Private Const WM_SETPROGRAM               As Long = WM_USER + 22
Private Const WM_CANDO                    As Long = WM_USER + 23
Private Const WM_EDITIDLE                 As Long = WM_USER + 24
Private Const WM_EDITCLOSE                As Long = WM_USER + 25
Private Const WM_EDITOPEN                 As Long = WM_USER + 26
Private Const WM_SETSAMPLERATE            As Long = WM_USER + 27
Private Const WM_MAINSCHANGED             As Long = WM_USER + 28
Private Const WM_GETPARAMNAME             As Long = WM_USER + 29
Private Const WM_GETPARAMLABEL            As Long = WM_USER + 30
Private Const WM_GETPARAMDISPLAY          As Long = WM_USER + 31
Private Const WM_GETPARAMETER             As Long = WM_USER + 32
Private Const WM_SETPARAMETER             As Long = WM_USER + 33
Private Const WM_PROCESSREPLACING         As Long = WM_USER + 34
Private Const WM_PROCESS                  As Long = WM_USER + 35

Private Const THREAD_WND_CLASS      As String = "VBVSTWndClass"             ' // This class manages STA requests
Private Const CONTAINER_WND_CLASS   As String = "VBContainerVSTWndClass"    ' // This class is container to VST editor window

' // Process (ProcessReplacing) data
Private Type tProcessSamplesData
    pIn         As PTR
    pOut        As PTR
    lSamples    As Long
End Type

Private Type tVendorSpecificData
    lArg1       As Long
    lArg2       As Long
    lpPtr       As PTR
    fArg3       As Single
End Type

Private Type tGetProgramNameIndexedData
    lCategory   As Long
    lIndex      As Long
    pName       As PTR
End Type

Private Type tGetSetChunkData
    bIsPreset   As Boolean
    pData       As PTR
    lSize       As Long
End Type

Private Type tParamNameLabelData
    lIndex      As Long
    pName       As PTR
End Type

Private Type tParamValueData
    lIndex      As Long
    fValue      As Single
End Type

Private Type tEditOpenData
    hWndHost        As Handle
    hWndContainer   As Handle
End Type

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
    
Private m_lInstances    As Long                 ' // Number of instances (volatile)
Private m_lGlobalLock   As Long                 ' // Global data lock (volatile)
Private m_hInstance     As PTR                  ' // Dll base address (hInstance)
Private m_tVSTClsID     As UUID                 ' // CLSID of plugin class (code finds it by name VST_PLUGIN_CLASS_NAME)
Private m_hSTAThread    As Handle               ' // Handle of STA thread
Private m_lSTAThreadID  As Long                 ' // STA thread ID
Private m_hWndThread    As Handle               ' // STA thread window handle whci manages requests
Private m_pfnCanUnload  As PTR                  ' // DllCanUnloadNow address
Private m_bInitialized  As Boolean              ' // True if module is initialized

' // Get current time info from host
Public Function GetTimeInfo( _
                ByVal pHostCB As PTR, _
                ByVal pAEffect As PTR, _
                ByVal lFilter As Long) As VstTimeInfo
    Dim pRet    As PTR
    
    pRet = CallHostCallback(pHostCB, pAEffect, audioMasterGetTime, 0, lFilter, 0, 0)
                    
    If pRet Then
        memcpy GetTimeInfo, ByVal pRet, Len(GetTimeInfo)
    End If
    
End Function

' // Get channel data
Public Sub GetChannelData( _
           ByVal pData As PTR, _
           ByVal lChannel As Long, _
           ByVal lCount As Long, _
           ByRef fOut() As Single, _
           ByRef tSADesc As SAFEARRAY1D)
    
    With tSADesc
        .cbElements = 4
        .cDims = 1
        .fFeatures = FADF_AUTO
        GetMemPtr ByVal pData + lChannel * Len(.pvData), .pvData
        .rgsabound(0).cElements = lCount
    End With
    
    vbaAryMove ByVal ArrPtr(fOut), VarPtr(tSADesc)
    
End Sub

' // Create an audio effect
Private Function VSTPluginMain CDecl( _
                 ByVal pfnAudioMasterCallback As PTR) As PTR
    Dim tEffect     As AEEffect
    Dim tSADispDesc As SAFEARRAY1D
    Dim pEffect     As PTR
    
    If GetHostVersion(pfnAudioMasterCallback) = 0 Then
        Exit Function
    End If
    
    pEffect = CoTaskMemRealloc(NULL_PTR, LenB(tEffect) + LenB(tSADispDesc) + LenB(tEffect.object(0)))
    If pEffect = NULL_PTR Then
        Exit Function
    End If
    
    ' // Setup SafeArray
    With tSADispDesc
        .cbElements = LenB(tEffect.object(0))
        .cDims = 1
        .fFeatures = FADF_AUTO
        .pvData = pEffect + LenB(tEffect) + LenB(tSADispDesc)
        .rgsabound(0).cElements = 1
    End With
    
    ' // Init SafeArray
    memcpy ByVal pEffect + LenB(tEffect), tSADispDesc, Len(tSADispDesc)
    memset ByVal pEffect + LenB(tEffect) + LenB(tSADispDesc), Len(tEffect.object(0)), 0
    PutMemPtr ByVal ArrPtr(tEffect.object), pEffect + LenB(tEffect)
    
    With tEffect
        
        If CreateVSTEffect(pfnAudioMasterCallback, pEffect, .object(0)) < 0 Then
            CoTaskMemFree pEffect
            Exit Function
        End If

        .magic = kEffectMagic
        
        .Dispatcher = FAR_PROC(AddressOf Dispatcher)
        .SetParameter = FAR_PROC(AddressOf SetParameter)
        .GetParameter = FAR_PROC(AddressOf GetParameter)
        .ProcessReplacing = FAR_PROC(AddressOf ProcessReplacing)
        .Process = FAR_PROC(AddressOf Process)
        
        .numParams = .object(0).lNumOfParams
        .numInputs = .object(0).lNumOfInputs
        .numOutputs = .object(0).lNumOfOutputs
        .numPrograms = .object(0).lNumOfPrograms
        .UniqueId = .object(0).lUniqueId
        .Version = .object(0).lVersion

        .flags = effFlagsCanReplacing
        
        If .object(0).bHasEditor Then
            .flags = .flags Or effFlagsHasEditor
        End If
        
        If .object(0).bProgAreChunks Then
            .flags = .flags Or effFlagsProgramChunks
        End If
        
        .Version = 1
        .ioRatio = 1

        vbaCopyBytesZero Len(tEffect), ByVal pEffect, tEffect

        VSTPluginMain = pEffect
        
    End With
    
End Function

' // Opcodes dispatcher
Private Function Dispatcher CDecl( _
                 ByRef tEffect As AEEffect, _
                 ByVal lOpcode As AEffectOpcodes, _
                 ByVal lIndex As Long, _
                 ByVal lValue As Long, _
                 ByVal lpPtr As PTR, _
                 ByVal fOpt As Single) As Long
    
    Select Case lOpcode
    Case effClose
        Dispatcher = OnEffectClose(tEffect)
    Case effSetSampleRate
        Dispatcher = OnSetSampleRate(tEffect, fOpt)
    Case effMainsChanged
        Dispatcher = OnEnable(tEffect, lValue)
    Case effGetParamName
        Dispatcher = OnGetParamName(tEffect, lIndex, lpPtr)
    Case effGetParamLabel
        Dispatcher = OnGetParamLabel(tEffect, lIndex, lpPtr)
    Case effGetParamDisplay
        Dispatcher = OnGetParamDisplay(tEffect, lIndex, lpPtr)
    Case effEditGetRect
        Dispatcher = OnEffectEditGetRect(tEffect, lpPtr)
    Case effEditOpen
        Dispatcher = OnEffectEditOpen(tEffect, lpPtr)
    Case effEditClose
        Dispatcher = OnEffectEditClose(tEffect)
    Case effEditIdle
        Dispatcher = OnEffectEditIdle(tEffect)
    Case effCanDo
        Dispatcher = OnEffectCanDo(tEffect, lpPtr)
    Case effSetProgram
        Dispatcher = OnSetProgram(tEffect, lValue)
    Case effGetProgram
        Dispatcher = OnGetProgram(tEffect)
    Case effSetProgramName
        Dispatcher = OnSetProgramName(tEffect, lpPtr)
    Case effGetProgramName
        Dispatcher = OnGetProgramName(tEffect, lpPtr)
    Case effSetBlockSize
        Dispatcher = OnSetBlockSize(tEffect, lValue)
    Case effGetChunk
        Dispatcher = OnGetChunk(tEffect, lpPtr, lIndex)
    Case effSetChunk
        Dispatcher = OnSetChunk(tEffect, lpPtr, lValue, lIndex)
    Case effGetVstVersion
        Dispatcher = OnGetVstVersion(tEffect)
    Case effGetParameterProperties
        Dispatcher = OnGetParameterProperties(tEffect, lIndex, lpPtr)
    Case effCanBeAutomated
        Dispatcher = OnCanBeAutomated(tEffect, lIndex)
    Case effGetProgramNameIndexed
        Dispatcher = OnGetProgramNameIndexed(tEffect, lValue, lIndex, lpPtr)
    Case effGetEffectName
        Dispatcher = OnGetEffectName(tEffect, lpPtr)
    Case effGetVendorString
        Dispatcher = OnGetVendorString(tEffect, lpPtr)
    Case effGetProductString
        Dispatcher = OnGetProductString(tEffect, lpPtr)
    Case effGetVendorVersion
        Dispatcher = OnGetVendorVersion(tEffect)
    Case effVendorSpecific
        Dispatcher = OnVendorSpecific(tEffect, lIndex, lValue, lpPtr, fOpt)
    Case effGetTailSize
        Dispatcher = OnGetTailSize(tEffect)
    Case effProcessEvents
        Dispatcher = OnProcessEvents(tEffect, lpPtr)
    Case effGetPlugCategory
        Dispatcher = OnGetPlugCategory(tEffect)
    Case effSetBypass
        Dispatcher = OnSetBypass(tEffect, lValue)
    Case effStartProcess
        Dispatcher = OnStartProcess(tEffect)
    Case effStopProcess
        Dispatcher = OnStopProcess(tEffect)
    End Select
    
End Function

' // Handlers
Private Function OnStopProcess( _
                 ByRef tEffect As AEEffect) As Long
    OnStopProcess = SendMessage(m_hWndThread, WM_STOPPROCESS, 0, ByVal tEffect.object(0).pPlugObjPtr)
End Function

Private Function OnStartProcess( _
                 ByRef tEffect As AEEffect) As Long
    OnStartProcess = SendMessage(m_hWndThread, WM_STARTPROCESS, 0, ByVal tEffect.object(0).pPlugObjPtr)
End Function

Private Function OnSetBypass( _
                 ByRef tEffect As AEEffect, _
                 ByVal lValue As Long) As Long
    OnSetBypass = SendMessage(m_hWndThread, WM_SETBYPASS, lValue, ByVal tEffect.object(0).pPlugObjPtr)
End Function

Private Function OnGetPlugCategory( _
                 ByRef tEffect As AEEffect) As Long
    OnGetPlugCategory = tEffect.object(0).lPlugCategory
End Function

Private Function OnProcessEvents( _
                 ByRef tEffect As AEEffect, _
                 ByVal pEvents As PTR) As Long
    If tEffect.object(0).bSupportsEvents Then
        OnProcessEvents = SendMessage(m_hWndThread, WM_PROCESSEVENTS, pEvents, ByVal tEffect.object(0).pPlugObjPtr)
    End If
End Function

Private Function OnGetTailSize( _
                 ByRef tEffect As AEEffect) As Long
    OnGetTailSize = SendMessage(m_hWndThread, WM_GETTAILSIZE, 0, ByVal tEffect.object(0).pPlugObjPtr)
End Function

Private Function OnVendorSpecific( _
                 ByRef tEffect As AEEffect, _
                 ByVal lArg1 As Long, _
                 ByVal lArg2 As Long, _
                 ByVal lpPtr As PTR, _
                 ByVal fArg3 As Single) As Long
    Dim tData   As tVendorSpecificData
    
    tData.lArg1 = lArg1
    tData.lArg2 = lArg2
    tData.lpPtr = lpPtr
    tData.fArg3 = fArg3
    
    OnVendorSpecific = SendMessage(m_hWndThread, WM_VENDORSPECIFIC, VarPtr(tData), ByVal tEffect.object(0).pPlugObjPtr)
    
End Function

Private Function OnGetVendorVersion( _
                 ByRef tEffect As AEEffect) As Long
    OnGetVendorVersion = SendMessage(m_hWndThread, WM_VENDORVERSION, 0, ByVal tEffect.object(0).pPlugObjPtr)
End Function

Private Function OnGetProductString( _
                 ByRef tEffect As AEEffect, _
                 ByVal pName As PTR) As Long
    OnGetProductString = SendMessage(m_hWndThread, WM_GETPRODUCTSTRING, pName, ByVal tEffect.object(0).pPlugObjPtr)
End Function

Private Function OnGetVendorString( _
                 ByRef tEffect As AEEffect, _
                 ByVal pName As PTR) As Long
    OnGetVendorString = SendMessage(m_hWndThread, WM_GETVENDORSTRING, pName, ByVal tEffect.object(0).pPlugObjPtr)
End Function

Private Function OnGetEffectName( _
                 ByRef tEffect As AEEffect, _
                 ByVal pName As PTR) As Long
    OnGetEffectName = SendMessage(m_hWndThread, WM_GETEFFECTNAME, pName, ByVal tEffect.object(0).pPlugObjPtr)
End Function

Private Function OnGetProgramNameIndexed( _
                 ByRef tEffect As AEEffect, _
                 ByVal lCategory As Long, _
                 ByVal lIndex As Long, _
                 ByVal pName As PTR) As Long
    Dim tData   As tGetProgramNameIndexedData
    
    tData.lCategory = lCategory
    tData.lIndex = lIndex
    tData.pName = pName
    
    OnGetProgramNameIndexed = SendMessage(m_hWndThread, WM_GETPROGRAMNAMEINDEXED, VarPtr(tData), _
                                          ByVal tEffect.object(0).pPlugObjPtr)
    
End Function

Private Function OnCanBeAutomated( _
                 ByRef tEffect As AEEffect, _
                 ByVal lIndex As Long) As Long
    OnCanBeAutomated = SendMessage(m_hWndThread, WM_CANBEAUTOMATED, lIndex, ByVal tEffect.object(0).pPlugObjPtr)
End Function

Private Function OnGetParameterProperties( _
                 ByRef tEffect As AEEffect, _
                 ByVal lIndex As Long, _
                 ByVal pProperties As PTR) As Long
    Dim pAllProp    As PTR
    
    pAllProp = tEffect.object(0).pParamProp
    
    If pAllProp = NULL_PTR Then
        Exit Function
    End If
    
    memcpy ByVal pProperties, ByVal pAllProp + lIndex * SIZEOF_VstParameterProperties, SIZEOF_VstParameterProperties
    
    OnGetParameterProperties = 1
    
End Function

Private Function OnGetVstVersion( _
                 ByRef tEffect As AEEffect) As Long
    OnGetVstVersion = tEffect.object(0).lVstVersion
End Function

Private Function OnSetChunk( _
                 ByRef tEffect As AEEffect, _
                 ByVal pData As PTR, _
                 ByVal lSize As Long, _
                 ByVal bIsPreset As Boolean) As Long
    Dim tData   As tGetSetChunkData
    
    tData.bIsPreset = bIsPreset
    tData.pData = pData
    tData.lSize = lSize
    
    OnSetChunk = SendMessage(m_hWndThread, WM_SETCHUNK, VarPtr(tData), ByVal tEffect.object(0).pPlugObjPtr)
    
End Function

Private Function OnGetChunk( _
                 ByRef tEffect As AEEffect, _
                 ByVal ppData As PTR, _
                 ByVal bIsPreset As Boolean) As Long
    Dim tData   As tGetSetChunkData
    
    tData.bIsPreset = bIsPreset
    tData.pData = ppData
    
    OnGetChunk = SendMessage(m_hWndThread, WM_GETCHUNK, VarPtr(tData), ByVal tEffect.object(0).pPlugObjPtr)

End Function

Private Function OnSetBlockSize( _
                 ByRef tEffect As AEEffect, _
                 ByVal lValue As Long) As Long
    SendMessage m_hWndThread, WM_SETBLOCKSIZE, lValue, ByVal tEffect.object(0).pPlugObjPtr
End Function

Private Function OnGetProgramName( _
                 ByRef tEffect As AEEffect, _
                 ByVal pName As PTR) As Long
    SendMessage m_hWndThread, WM_GETPROGRAMNAME, pName, ByVal tEffect.object(0).pPlugObjPtr
End Function

Private Function OnSetProgramName( _
                 ByRef tEffect As AEEffect, _
                 ByVal pName As PTR) As Long
    SendMessage m_hWndThread, WM_SETPROGRAMNAME, pName, ByVal tEffect.object(0).pPlugObjPtr
End Function

Private Function OnGetProgram( _
                 ByRef tEffect As AEEffect) As Long
    OnGetProgram = SendMessage(m_hWndThread, WM_GETPROGRAM, 0, ByVal tEffect.object(0).pPlugObjPtr)
End Function

Private Function OnSetProgram( _
                 ByRef tEffect As AEEffect, _
                 ByVal lIndex As Long) As Long
    SendMessage m_hWndThread, WM_SETPROGRAM, lIndex, ByVal tEffect.object(0).pPlugObjPtr
End Function

Private Function OnEffectCanDo( _
                 ByRef tEffect As AEEffect, _
                 ByVal pszRequest As PTR) As Long
    OnEffectCanDo = SendMessage(m_hWndThread, WM_CANDO, pszRequest, ByVal tEffect.object(0).pPlugObjPtr)
End Function

Private Function OnEffectEditIdle( _
                 ByRef tEffect As AEEffect) As Long
    Dim tRecords()  As tAutomationRecord
    Dim lCountRec   As Long
    Dim lIndex      As Long
    
    lCountRec = SendMessage(m_hWndThread, WM_EDITIDLE, ArrPtr(tRecords), ByVal tEffect.object(0).pPlugObjPtr)

    ' // Write automation events
    For lIndex = 0 To lCountRec - 1
        CallHostCallback tEffect.object(0).pHostCallback, tEffect.object(0).pAEffect, audioMasterAutomate, _
                         tRecords(lIndex).lParamIndex, 0, 0, tRecords(lIndex).fParamValue
    Next

End Function

Private Function OnEffectEditClose( _
                 ByRef tEffect As AEEffect) As Long
    SendMessage m_hWndThread, WM_EDITCLOSE, tEffect.object(0).hWndContainer, ByVal tEffect.object(0).pPlugObjPtr
End Function

Private Function OnEffectEditOpen( _
                 ByRef tEffect As AEEffect, _
                 ByVal hWnd As Handle) As Long
    Dim tData   As tEditOpenData
    
    tData.hWndContainer = tEffect.object(0).hWndContainer
    tData.hWndHost = hWnd
    
    OnEffectEditOpen = SendMessage(m_hWndThread, WM_EDITOPEN, VarPtr(tData), ByVal tEffect.object(0).pPlugObjPtr)
    
End Function

Private Function OnEffectEditGetRect( _
                 ByRef tEffect As AEEffect, _
                 ByVal ppRect As PTR) As Long

    PutMemPtr ByVal ppRect, VarPtr(tEffect.object(0).tRect)
    
    If tEffect.object(0).tRect.wRight > 0 Then
        OnEffectEditGetRect = 1
    End If
    
End Function

Private Function OnEffectClose( _
                 ByRef tEffect As AEEffect) As Long
    DestroyPlugin tEffect
    OnEffectClose = 1
End Function

Private Function OnSetSampleRate( _
                 ByRef tEffect As AEEffect, _
                 ByVal fValue As Single) As Long
    Dim lValue  As Long
    
    GetMem4 fValue, lValue
    
    SendMessage m_hWndThread, WM_SETSAMPLERATE, lValue, ByVal tEffect.object(0).pPlugObjPtr
    
End Function

Private Function OnEnable( _
                 ByRef tEffect As AEEffect, _
                 ByVal lValue As Long) As Long
    SendMessage m_hWndThread, WM_MAINSCHANGED, lValue, ByVal tEffect.object(0).pPlugObjPtr
End Function

Private Function OnGetParamName( _
                 ByRef tEffect As AEEffect, _
                 ByVal lIndex As Long, _
                 ByVal pBuffer As PTR) As Long
    Dim tData   As tParamNameLabelData
    
    tData.lIndex = lIndex
    tData.pName = pBuffer
    
    SendMessage m_hWndThread, WM_GETPARAMNAME, VarPtr(tData), ByVal tEffect.object(0).pPlugObjPtr
    
End Function

Private Function OnGetParamLabel( _
                 ByRef tEffect As AEEffect, _
                 ByVal lIndex As Long, _
                 ByVal pBuffer As PTR) As Long
    Dim tData   As tParamNameLabelData
    
    tData.lIndex = lIndex
    tData.pName = pBuffer
    
    SendMessage m_hWndThread, WM_GETPARAMLABEL, VarPtr(tData), ByVal tEffect.object(0).pPlugObjPtr
    
End Function

Private Function OnGetParamDisplay( _
                 ByRef tEffect As AEEffect, _
                 ByVal lIndex As Long, _
                 ByVal pBuffer As PTR) As Long
    Dim tData   As tParamNameLabelData
    
    tData.lIndex = lIndex
    tData.pName = pBuffer
    
    SendMessage m_hWndThread, WM_GETPARAMDISPLAY, VarPtr(tData), ByVal tEffect.object(0).pPlugObjPtr
                 
End Function

Private Function GetParameter CDecl( _
                 ByRef tEffect As AEEffect, _
                 ByVal lIndex As Long) As Single
    GetMem4 SendMessage(m_hWndThread, WM_GETPARAMETER, lIndex, ByVal tEffect.object(0).pPlugObjPtr), GetParameter
End Function

Private Function SetParameter CDecl( _
                 ByRef tEffect As AEEffect, _
                 ByVal lIndex As Long, _
                 ByVal fValue As Single) As Long
    Dim tData   As tParamValueData
    
    tData.lIndex = lIndex
    tData.fValue = fValue
    
    SendMessage m_hWndThread, WM_SETPARAMETER, VarPtr(tData), ByVal tEffect.object(0).pPlugObjPtr
    
End Function

Private Function ProcessReplacing CDecl( _
                 ByRef tEffect As AEEffect, _
                 ByVal lpInputs As PTR, _
                 ByVal lpOutputs As PTR, _
                 ByVal lSampleFrames As Long) As Long
    Dim tData   As tProcessSamplesData
    
    tData.pIn = lpInputs
    tData.pOut = lpOutputs
    tData.lSamples = lSampleFrames
    
    SendMessage m_hWndThread, WM_PROCESSREPLACING, VarPtr(tData), ByVal tEffect.object(0).pPlugObjPtr

End Function

Private Function Process CDecl( _
                 ByRef tEffect As AEEffect, _
                 ByVal lpInputs As PTR, _
                 ByVal lpOutputs As PTR, _
                 ByVal lSampleFrames As Long) As Long
    Dim tData   As tProcessSamplesData
    
    tData.pIn = lpInputs
    tData.pOut = lpOutputs
    tData.lSamples = lSampleFrames
    
    SendMessage m_hWndThread, WM_PROCESS, VarPtr(tData), ByVal tEffect.object(0).pPlugObjPtr
    
End Function

Private Function GetHostVersion( _
                 ByVal pHostCallback As PTR) As Long
    
    If pHostCallback = NULL_PTR Then
        GetHostVersion = 0
    Else
        GetHostVersion = CallHostCallback(pHostCallback, NULL_PTR, audioMasterVersion, 0, 0, NULL_PTR, 0)
    End If
    
End Function

Private Function CallHostCallback( _
                 ByVal pfn As PTR, _
                 ByVal pAEffect As PTR, _
                 ByVal lOpcode As Long, _
                 ByVal lIndex As Long, _
                 ByVal lValue As PTR, _
                 ByVal lptr As PTR, _
                 ByVal fValue As Single) As PTR
    Dim tArgs   As tDispCallFuncArgData
    Dim hr      As Long

    With tArgs
    
        .iTypes(0) = vbString:  .iTypes(1) = vbLong
        .iTypes(2) = vbLong:    .iTypes(3) = vbString
        .iTypes(4) = vbString:  .iTypes(5) = vbSingle
        
        .lVarArgs(0).iVT = vbString:    .lVarArgs(0).pData0 = pAEffect
        .lVarArgs(1).iVT = vbLong:      .lVarArgs(1).pData0 = lOpcode
        .lVarArgs(2).iVT = vbLong:      .lVarArgs(2).pData0 = lIndex
        .lVarArgs(3).iVT = vbString:    .lVarArgs(3).pData0 = lValue
        .lVarArgs(4).iVT = vbString:    .lVarArgs(4).pData0 = lptr
        .lVarArgs(5).iVT = vbSingle:    GetMem4 fValue, .lVarArgs(5).pData0
        
        .pArgs(0) = VarPtr(.lVarArgs(0)):   .pArgs(1) = VarPtr(.lVarArgs(1))
        .pArgs(2) = VarPtr(.lVarArgs(2)):   .pArgs(3) = VarPtr(.lVarArgs(3))
        .pArgs(4) = VarPtr(.lVarArgs(4)):   .pArgs(5) = VarPtr(.lVarArgs(5))
        
    End With
    
    hr = DispCallFunc(ByVal NULL_PTR, pfn, CC_CDECL, vbString, 6, tArgs.iTypes(0), tArgs.pArgs(0), tArgs.lVarRet)
    
    If hr >= 0 Then
        CallHostCallback = tArgs.lVarRet.pData0
    End If

End Function

' // Lock global data access
Private Sub AcquireGlobalLock()
    Do While InterlockedCompareExchange(m_lGlobalLock, 1, 0)
        Sleep 10
    Loop
End Sub

' // Release global lock
Private Sub ReleaseGlobalLock()
    InterlockedExchange m_lGlobalLock, 0
End Sub

' // Initialize global data
' // Don't forget ensure serialized access to this function (use AcquireGlobalLock)
Private Function InitializeGlobalData() As Long
    Dim hInstance   As PTR
    Dim tClsID      As UUID
    Dim hThread     As Handle
    Dim lThreadID   As Long
    Dim hEvent      As Handle
    Dim tWndClass   As WNDCLASSEX
    
    If m_bInitialized Then
        Exit Function
    End If

    ' // Get hInstance
    If GetModuleHandleEx(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS Or GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT, _
                         ByVal FAR_PROC(AddressOf FAR_PROC), hInstance) = 0 Then
        InitializeGlobalData = HRESULTFromWin32(GetLastError)
        Exit Function
    End If

    m_hInstance = hInstance
    
    m_pfnCanUnload = GetProcAddress(hInstance, "DllCanUnloadNow")
    
    ' // Modify header to bypass overwriting global variables
    If Not ModifyVBHeader(hInstance) Then
        InitializeGlobalData = E_FAIL
        GoTo CleanUp
    End If
    
    ' // Search for CLSID
    If Not GetVSTClassID(hInstance, tClsID) Then
        InitializeGlobalData = E_UNEXPECTED
        GoTo CleanUp
    End If
    
    ' // Create initialization event
    hEvent = CreateEvent(ByVal NULL_PTR, 0, 0, vbNullString)
    If hEvent = NULL_PTR Then
        InitializeGlobalData = HRESULTFromWin32(GetLastError)
        GoTo CleanUp
    End If

    With tWndClass
        
        .cbSize = Len(tWndClass)
        .hInstance = hInstance
        .lpszClassName = StrPtr(THREAD_WND_CLASS)
        .lpfnWndProc = FAR_PROC(AddressOf ThreadWndProc)
        
    End With
    
    ' // Register STA window class
    If RegisterClassEx(tWndClass) = 0 Then
        InitializeGlobalData = HRESULTFromWin32(GetLastError)
        GoTo CleanUp
    End If
    
    ' // Register container class
    With tWndClass

        .lpszClassName = StrPtr(CONTAINER_WND_CLASS)
        .lpfnWndProc = FAR_PROC(AddressOf ContainerWndProc)
        
    End With
    
    If RegisterClassEx(tWndClass) = 0 Then
        InitializeGlobalData = HRESULTFromWin32(GetLastError)
        GoTo CleanUp
    End If
    
    ' // Create STA thread
    hThread = CreateThread(ByVal NULL_PTR, 0, AddressOf STAThread, ByVal hEvent, 0, lThreadID)
    If hThread = 0 Then
        InitializeGlobalData = HRESULTFromWin32(GetLastError)
        GoTo CleanUp
    End If
    
    ' // Wait thread initialization
    If WaitForSingleObject(hEvent, -1) <> WAIT_OBJECT_0 Then
        InitializeGlobalData = HRESULTFromWin32(GetLastError)
        GoTo CleanUp
    ElseIf m_hWndThread = 0 Then
        InitializeGlobalData = E_FAIL
        GoTo CleanUp
    End If
    
    m_hSTAThread = hThread
    m_lSTAThreadID = lThreadID
    m_tVSTClsID = tClsID
    m_bInitialized = True

CleanUp:
    
    If hEvent Then
        CloseHandle hEvent
    End If
    
    If Not m_bInitialized Then
        
        ' // Clean up on error
        UnregisterClass CONTAINER_WND_CLASS, hInstance
        UnregisterClass THREAD_WND_CLASS, hInstance

        If hThread Then
            ExitSTAThread
        End If
        
    End If
    
End Function

' // Uninitialize global data
' // Don't forget ensure serialized access to this function
Private Sub UninitializeGlobalData()
    
    If Not m_bInitialized Then
        Exit Sub
    End If
    
    ExitSTAThread
    UnregisterClass CONTAINER_WND_CLASS, m_hInstance
    UnregisterClass THREAD_WND_CLASS, m_hInstance
    m_hInstance = NULL_PTR
    m_bInitialized = False
    
    CoFreeUnusedLibraries
    
End Sub

' // Exit STA thread
Private Sub ExitSTAThread()
    
    ' // Send request
    SendMessage m_hWndThread, WM_EXITTHREAD, 0, ByVal NULL_PTR
    WaitForSingleObject m_hSTAThread, -1
    CloseHandle m_hSTAThread
    
    m_hSTAThread = 0
    m_lSTAThreadID = 0
    
End Sub

' // Modify VB-headers to prevent zeroing global variables !in this module!
Private Function ModifyVBHeader( _
                 ByVal hInstance As PTR) As Boolean
    Dim pfn             As PTR
    Dim pVbHeader       As PTR
    Dim pVbProjInfo     As PTR
    Dim pVbObjTable     As PTR
    Dim pVbObjDesc      As PTR
    Dim lModulesCount   As Long
    Dim lModuleIndex    As Long
    Dim pPubResDesc     As PTR
    Dim pPubVars        As PTR
    Dim lDataSize       As Long
    Dim pTestVar        As PTR
    Dim lOldProtect     As Long
    
    pfn = GetProcAddress(hInstance, "DllGetClassObject")
    
    If pfn = 0 Then
        Exit Function
    End If
    
    GetMemPtr ByVal pfn + &H2, pVbHeader
    GetMemPtr ByVal pVbHeader + &H30, pVbProjInfo
    GetMemPtr ByVal pVbProjInfo + &H4, pVbObjTable
    GetMemPtr ByVal pVbObjTable + &H30, pVbObjDesc
    GetMem4 ByVal pVbObjTable + &H2A, lModulesCount
    
    pTestVar = VarPtr(m_lInstances)
    
    ' // There is no static vars in this module so we don't touch them
    For lModuleIndex = 0 To lModulesCount - 1
        
        GetMemPtr ByVal pVbObjDesc + &H8, pPubResDesc
        GetMemPtr ByVal pVbObjDesc + &H10, pPubVars
        GetMem2 ByVal pPubResDesc + 2, lDataSize
        
        If pTestVar >= pPubVars And pTestVar < pPubVars + lDataSize Then
            
            ' // This module
            If VirtualProtect(pPubResDesc + 2, 2, PAGE_EXECUTE_READWRITE, lOldProtect) = 0 Then
                Exit Function
            End If
            
            PutMem2 ByVal pPubResDesc + 2, 0
            
            VirtualProtect pPubResDesc + 2, 2, lOldProtect, lOldProtect
            
            ModifyVBHeader = True
            Exit Function
            
        End If
        
        pVbObjDesc = pVbObjDesc + &H30
        
    Next

End Function

' // Get CLSID of VST plugin
Private Function GetVSTClassID( _
                 ByVal hInstance As PTR, _
                 ByRef tClsID As UUID) As Boolean
    Dim pfn         As PTR
    Dim pVbHdr      As PTR
    Dim lSignature  As Long
    Dim pCOMData    As PTR
    Dim lOffset     As Long
    Dim lOffsetName As Long
    Dim sNameANSI   As String
    
    pfn = GetProcAddress(hInstance, "DllGetClassObject")
    
    If pfn = 0 Then
        Exit Function
    End If
    
    GetMemPtr ByVal pfn + 2, pVbHdr
    GetMem4 ByVal pVbHdr, lSignature
    
    If lSignature <> &H21354256 Then
        Exit Function
    End If
    
    vbaStrToAnsi sNameANSI, VST_PLUGIN_CLASS_NAME

    GetMemPtr ByVal pVbHdr + &H54, pCOMData
    GetMem4 ByVal pCOMData, lOffset
    
    Do While lOffset
        
        GetMem4 ByVal pCOMData + lOffset + 4, lOffsetName
        
        If lstrcmpA(ByVal pCOMData + lOffsetName, ByVal StrPtr(sNameANSI)) = 0 Then
        
            memcpy tClsID, ByVal pCOMData + lOffset + &H14, Len(tClsID)
            GetVSTClassID = True
            Exit Do
            
        End If
        
        GetMem4 ByVal pCOMData + lOffset, lOffset
        
    Loop

End Function

' // Create class object in STA thread and init dispatcher
Private Function CreateVSTEffect( _
                 ByVal pAudioMasterCallback As PTR, _
                 ByVal pAEffect As PTR, _
                 ByRef tOut As CVBVstDispatcher) As Long
    Dim hr          As Long
    Dim bComInit    As Boolean
    
    ' // Initialize global data if need
    AcquireGlobalLock
    
    hr = InitializeGlobalData()
    InterlockedIncrement m_lInstances
    
    ReleaseGlobalLock
    
    If hr < 0 Then
        GoTo CleanUp
    End If
    
    hr = CoInitialize(ByVal NULL_PTR)
    
    If hr < 0 Then
        GoTo CleanUp
    End If
    
    bComInit = True

    hr = CreateWrappedPlugin(pAudioMasterCallback, pAEffect, tOut)

    If hr < 0 Then
        GoTo CleanUp
    End If

CleanUp:

    If hr < 0 Then
        
        If bComInit Then
            CoUninitialize
        End If
        
        AcquireGlobalLock
        
        If InterlockedCompareExchange(m_lInstances, 0, 1) = 1 Then
            UninitializeGlobalData
        Else
            InterlockedDecrement m_lInstances
        End If
        
        ReleaseGlobalLock
        
    End If
    
    CreateVSTEffect = hr
    
End Function

' // Create MTA object which dispatch all the calls to STA object
Private Function CreateWrappedPlugin( _
                 ByVal pAudioMasterCallback As PTR, _
                 ByVal pAEffect As PTR, _
                 ByRef tOut As CVBVstDispatcher) As Long
    Dim hr          As Long
    Dim bLockInit   As Boolean
    Dim lTlsIndex   As Long
    
    lTlsIndex = TLS_OUT_OF_INDEXES

    lTlsIndex = TlsAlloc()
    If lTlsIndex = TLS_OUT_OF_INDEXES Then
        hr = HRESULTFromWin32(GetLastError)
        GoTo CleanUp
    End If

    With tOut

        InitializeCriticalSection .tLock
        
        bLockInit = True

        .lTlsIndex = lTlsIndex
        .pHostCallback = pAudioMasterCallback
        .pAEffect = pAEffect
        
        ' // Send request to STA thread
        hr = SendMessage(m_hWndThread, WM_CREATEPLUGIN, 0, tOut)
        
        If hr < 0 Then
            GoTo CleanUp
        End If
        
    End With

CleanUp:

    If hr < 0 Then

        If bLockInit Then
            DeleteCriticalSection tOut.tLock
        End If

        If lTlsIndex <> TLS_OUT_OF_INDEXES Then
            TlsFree lTlsIndex
        End If
        
    End If
    
    CreateWrappedPlugin = hr
    
End Function

' // Destroy plugin
Private Sub DestroyPlugin( _
            ByRef tEffect As AEEffect)
    Dim lIndex  As Long
    Dim pObj    As PTR
    
    With tEffect.object(0)
    
        EnterCriticalSection .tLock
    
        ' // Request to release reference in STA thread
        SendMessage m_hWndThread, WM_RELEASEREF, .hWndContainer, ByVal .pPlugObjPtr
    
        TlsFree .lTlsIndex
        
        LeaveCriticalSection .tLock
        DeleteCriticalSection .tLock
        
    End With
    
    AcquireGlobalLock
    
    If InterlockedDecrement(m_lInstances) = 0 Then
        UninitializeGlobalData
    End If
    
    ReleaseGlobalLock
    
    CoTaskMemFree VarPtr(tEffect)
    CoUninitialize
    
End Sub

'
' // These functions are called in STA thread  //

'  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
'  v  v  v  v  v  v  v  v  v  v  v  v  v  v  v

' // STA thread
' // This thread is used by all the objects
Private Function STAThread( _
                 ByVal hEvent As Handle) As Long
    Dim tMsg    As MSG
    Dim lRet    As Long
    Dim hr      As Long
    Dim hWnd    As Handle
    Dim tVar    As tOLEVariant
    
    ' // Create dispatcher window
    hWnd = CreateWindowEx(0, THREAD_WND_CLASS, vbNullString, 0, 0, 0, 0, 0, HWND_MESSAGE, 0, m_hInstance, ByVal NULL_PTR)
    If hWnd = 0 Then
        
        STAThread = HRESULTFromWin32(GetLastError)
        SetEvent hEvent
        Exit Function
        
    End If
    
    ' // Other threads can use this window now
    m_hWndThread = hWnd
    
    ' // Release thread
    SetEvent hEvent
    
    hr = CoInitialize(ByVal NULL_PTR)
    If hr < 0 Then
    
        DestroyWindow hWnd
        STAThread = hr
        Exit Function
        
    End If
    
    ' // Message pump
    Do
        
        lRet = GetMessage(tMsg, 0, 0, 0)
        
        If lRet = -1 Then
            ' // Error
            Exit Do
        ElseIf lRet = 0 Then
            Exit Do
        Else
            TranslateMessage tMsg
            DispatchMessage tMsg
        End If
        
    Loop
    
    m_hWndThread = 0
    
    ' // Release VB-project context
    If m_pfnCanUnload Then
        DispCallFunc ByVal NULL_PTR, m_pfnCanUnload, CC_STDCALL, vbEmpty, 0, ByVal NULL_PTR, ByVal NULL_PTR, tVar
    End If
    
    ' // Release all the COM libs used by our plugin (if any)
    CoFreeUnusedLibraries
    CoUninitialize
    
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
    Case Else
        ' // Do nothing to avoid deadlocks
    End Select
    
End Function

' // Thread window proc
Private Function ThreadWndProc( _
                 ByVal hWnd As Handle, _
                 ByVal lMsg As Long, _
                 ByVal wParam As PTR, _
                 ByVal lParam As PTR) As PTR
                 
    Select Case lMsg
    Case WM_STARTPROCESS
        ThreadWndProc = ReqStartProcess(lParam)
    Case WM_STOPPROCESS
        ThreadWndProc = ReqStopProcess(lParam)
    Case WM_SETBYPASS
        ThreadWndProc = ReqSetBypass(lParam, wParam)
    Case WM_PROCESSEVENTS
        ThreadWndProc = ReqProcessEvents(lParam, wParam)
    Case WM_GETTAILSIZE
        ThreadWndProc = ReqGetTailSize(lParam)
    Case WM_VENDORSPECIFIC
        ThreadWndProc = ReqVendorSpecific(lParam, wParam)
    Case WM_VENDORVERSION
        ThreadWndProc = ReqVendorVersion(lParam)
    Case WM_GETPRODUCTSTRING
        ThreadWndProc = ReqGetProductString(lParam, wParam)
    Case WM_GETVENDORSTRING
        ThreadWndProc = ReqGetVendorString(lParam, wParam)
    Case WM_GETEFFECTNAME
        ThreadWndProc = ReqGetEffectName(lParam, wParam)
    Case WM_GETPROGRAMNAMEINDEXED
        ThreadWndProc = ReqGetProgramNameIndexed(lParam, wParam)
    Case WM_CANBEAUTOMATED
        ThreadWndProc = ReqCanBeAutomated(lParam, wParam)
    Case WM_SETCHUNK
        ThreadWndProc = ReqSetChunk(lParam, wParam)
    Case WM_GETCHUNK
        ThreadWndProc = ReqGetChunk(lParam, wParam)
    Case WM_SETBLOCKSIZE
        ThreadWndProc = ReqSetBlockSize(lParam, wParam)
    Case WM_GETPROGRAMNAME
        ThreadWndProc = ReqGetProgramName(lParam, wParam)
    Case WM_SETPROGRAMNAME
        ThreadWndProc = ReqSetProgramName(lParam, wParam)
    Case WM_GETPROGRAM
        ThreadWndProc = ReqGetProgram(lParam)
    Case WM_SETPROGRAM
        ThreadWndProc = ReqSetProgram(lParam, wParam)
    Case WM_CANDO
        ThreadWndProc = ReqCanDo(lParam, wParam)
    Case WM_EDITIDLE
        ThreadWndProc = ReqEditIdle(lParam, wParam)
    Case WM_EDITCLOSE
        ThreadWndProc = ReqEditClose(lParam, wParam)
    Case WM_EDITOPEN
        ThreadWndProc = ReqEditOpen(lParam, wParam)
    Case WM_SETSAMPLERATE
        ThreadWndProc = ReqSetSampleRate(lParam, wParam)
    Case WM_MAINSCHANGED
        ThreadWndProc = ReqMainsChanged(lParam, wParam)
    Case WM_GETPARAMNAME
        ThreadWndProc = ReqGetParamName(lParam, wParam)
    Case WM_GETPARAMLABEL
        ThreadWndProc = ReqGetParamLabel(lParam, wParam)
    Case WM_GETPARAMDISPLAY
        ThreadWndProc = ReqGetParamDisplay(lParam, wParam)
    Case WM_GETPARAMETER
        ThreadWndProc = ReqGetParameter(lParam, wParam)
    Case WM_SETPARAMETER
        ThreadWndProc = ReqSetParameter(lParam, wParam)
    Case WM_PROCESSREPLACING
        ThreadWndProc = ReqProcessOrProcessReplacing(lParam, wParam, True)
    Case WM_PROCESS
        ThreadWndProc = ReqProcessOrProcessReplacing(lParam, wParam, False)
    Case WM_EXITTHREAD
        ThreadWndProc = ReqExitThread(hWnd)
    Case WM_CREATEPLUGIN
        ThreadWndProc = ReqCreatePlugin(lParam)
    Case WM_RELEASEREF
        ThreadWndProc = ReqReleaseRef(wParam, lParam)
    Case Else
        ThreadWndProc = DefWindowProc(hWnd, lMsg, wParam, ByVal lParam)
    End Select
                
End Function

' //////////////////////////////////////////////////////////////
' //                   Requests handlers                      //
' //////////////////////////////////////////////////////////////

Private Function ReqStartProcess( _
                 ByVal pObject As PTR) As Long
    Dim cPlugObj    As IVBVstEffect
    
    vbaObjSetAddref cPlugObj, ByVal pObject
    
    ReqStartProcess = cPlugObj.StartProcess And 1
    
End Function

Private Function ReqStopProcess( _
                 ByVal pObject As PTR) As Long
    Dim cPlugObj    As IVBVstEffect
    
    vbaObjSetAddref cPlugObj, ByVal pObject
    
    ReqStopProcess = cPlugObj.StopProcess And 1
    
End Function

Private Function ReqSetBypass( _
                 ByVal pObject As PTR, _
                 ByVal lValue As Long) As Long
    Dim cPlugObj    As IVBVstEffect
    
    vbaObjSetAddref cPlugObj, ByVal pObject
    
    ReqSetBypass = cPlugObj.SetBypass(lValue) And 1
    
End Function

Private Function ReqProcessEvents( _
                 ByVal pObject As PTR, _
                 ByVal pEvents As PTR) As Long
    Dim cPlugObj    As IVBVstEffect
    Dim tEvents     As VstEvents
    
    vbaObjSetAddref cPlugObj, ByVal pObject
    
    memcpy tEvents, ByVal pEvents, Len(tEvents)
    
    ReqProcessEvents = cPlugObj.ProcessEvents(tEvents) And 1
    
End Function

Private Function ReqGetTailSize( _
                 ByVal pObject As PTR) As Long
    Dim cPlugObj    As IVBVstEffect
    
    vbaObjSetAddref cPlugObj, ByVal pObject
    
    ReqGetTailSize = cPlugObj.TailSize
    
End Function

Private Function ReqVendorSpecific( _
                 ByVal pObject As PTR, _
                 ByVal pData As PTR) As Long
    Dim cPlugObj    As IVBVstEffect
    Dim tData       As tVendorSpecificData
    
    vbaObjSetAddref cPlugObj, ByVal pObject
    memcpy tData, ByVal pData, Len(tData)
    
    cPlugObj.VendorSpecific tData.lArg1, tData.lArg2, tData.lpPtr, tData.fArg3
    
    ReqVendorSpecific = 1
    
End Function

Private Function ReqVendorVersion( _
                 ByVal pObject As PTR) As Long
    Dim cPlugObj    As IVBVstEffect
    
    vbaObjSetAddref cPlugObj, ByVal pObject
    
    ReqVendorVersion = cPlugObj.VendorVersion
    
End Function

Private Function ReqGetProductString( _
                 ByVal pObject As PTR, _
                 ByVal pName As PTR) As Long
    Dim cPlugObj    As IVBVstEffect
    Dim sAnsi       As String
    
    vbaObjSetAddref cPlugObj, ByVal pObject
    
    sAnsi = Unicode2Ansi(cPlugObj.ProductString)

    If LenB(sAnsi) Then
        lstrcpynA ByVal pName, ByVal StrPtr(sAnsi), kVstMaxProductStrLen
    Else
        PutMem1 ByVal pName, 0
    End If
    
    ReqGetProductString = 1
    
End Function

Private Function ReqGetVendorString( _
                 ByVal pObject As PTR, _
                 ByVal pName As PTR) As Long
    Dim cPlugObj    As IVBVstEffect
    Dim sAnsi       As String
    
    vbaObjSetAddref cPlugObj, ByVal pObject

    sAnsi = Unicode2Ansi(cPlugObj.VendorString)
    
    If LenB(sAnsi) Then
        lstrcpynA ByVal pName, ByVal StrPtr(sAnsi), kVstMaxVendorStrLen
    Else
        PutMem1 ByVal pName, 0
    End If
    
    ReqGetVendorString = 1
    
End Function

Private Function ReqGetEffectName( _
                 ByVal pObject As PTR, _
                 ByVal pName As PTR) As Long
    Dim cPlugObj    As IVBVstEffect
    Dim sAnsi       As String
    
    vbaObjSetAddref cPlugObj, ByVal pObject

    sAnsi = Unicode2Ansi(cPlugObj.EffectName)
    
    If LenB(sAnsi) Then
        lstrcpynA ByVal pName, ByVal StrPtr(sAnsi), kVstMaxEffectNameLen
    Else
        PutMem1 ByVal pName, 0
    End If
    
    ReqGetEffectName = 1
    
End Function

Private Function ReqGetProgramNameIndexed( _
                 ByVal pObject As PTR, _
                 ByVal pData As PTR) As Long
    Dim cPlugObj    As IVBVstEffect
    Dim tData       As tGetProgramNameIndexedData
    Dim sAnsi       As String
    
    vbaObjSetAddref cPlugObj, ByVal pObject
    memcpy tData, ByVal pData, Len(tData)
    
    sAnsi = Unicode2Ansi(cPlugObj.ProgramNameIndexed(tData.lCategory, tData.lIndex))
    
    If LenB(sAnsi) Then
        lstrcpynA ByVal tData.pName, ByVal StrPtr(sAnsi), kVstMaxProgNameLen
    Else
        PutMem1 ByVal tData.pName, 0
    End If
    
    ReqGetProgramNameIndexed = 1
    
End Function

Private Function ReqCanBeAutomated( _
                 ByVal pObject As PTR, _
                 ByVal lIndex As Long) As Long
    Dim cPlugObj    As IVBVstEffect

    vbaObjSetAddref cPlugObj, ByVal pObject

    ReqCanBeAutomated = cPlugObj.CanParameterBeAutomated(lIndex) And 1
    
End Function

Private Function ReqSetChunk( _
                 ByVal pObject As PTR, _
                 ByVal pData As PTR) As Long
    Dim cPlugObj    As IVBVstEffect
    Dim tData       As tGetSetChunkData
    
    vbaObjSetAddref cPlugObj, ByVal pObject
    memcpy tData, ByVal pData, Len(tData)
    
    ReqSetChunk = cPlugObj.SetStateChunk(tData.bIsPreset, tData.pData, tData.lSize) And 1
    
End Function

Private Function ReqGetChunk( _
                 ByVal pObject As PTR, _
                 ByVal pData As PTR) As Long
    Dim cPlugObj    As IVBVstEffect
    Dim tData       As tGetSetChunkData
    Dim pResult     As PTR
    
    vbaObjSetAddref cPlugObj, ByVal pObject
    memcpy tData, ByVal pData, Len(tData)
    
    ReqGetChunk = cPlugObj.GetStateChunk(tData.bIsPreset, pResult)
    
    PutMemPtr ByVal tData.pData, pResult
    
End Function

Private Function ReqSetBlockSize( _
                 ByVal pObject As PTR, _
                 ByVal lValue As Long) As Long
    Dim cPlugObj    As IVBVstEffect

    vbaObjSetAddref cPlugObj, ByVal pObject

    cPlugObj.BlockSize = lValue

End Function

Private Function ReqGetProgramName( _
                 ByVal pObject As PTR, _
                 ByVal pName As PTR) As Long
    Dim cPlugObj    As IVBVstEffect
    Dim sAnsi       As String
    
    vbaObjSetAddref cPlugObj, ByVal pObject
    
    sAnsi = Unicode2Ansi(cPlugObj.ProgramName)
    
    If LenB(sAnsi) Then
        lstrcpynA ByVal pName, ByVal StrPtr(sAnsi), kVstMaxProgNameLen
    Else
        PutMem1 ByVal pName, 0
    End If
 
End Function

Private Function ReqSetProgramName( _
                 ByVal pObject As PTR, _
                 ByVal pName As PTR) As Long
    Dim cPlugObj    As IVBVstEffect

    vbaObjSetAddref cPlugObj, ByVal pObject

    cPlugObj.ProgramName = Ansi2Unicode(pName)

End Function

Private Function ReqGetProgram( _
                 ByVal pObject As PTR) As Long
    Dim cPlugObj    As IVBVstEffect

    vbaObjSetAddref cPlugObj, ByVal pObject

    ReqGetProgram = cPlugObj.Program

End Function

Private Function ReqSetProgram( _
                 ByVal pObject As PTR, _
                 ByVal lIndex As Long) As Long
    Dim cPlugObj    As IVBVstEffect

    vbaObjSetAddref cPlugObj, ByVal pObject

    cPlugObj.Program = lIndex

End Function

Private Function ReqCanDo( _
                 ByVal pObject As PTR, _
                 ByVal pRequest As PTR) As Long
    Dim cPlugObj    As IVBVstEffect

    vbaObjSetAddref cPlugObj, ByVal pObject

    ReqCanDo = cPlugObj.CanDo(Ansi2Unicode(pRequest)) And 1

End Function

Private Function ReqEditIdle( _
                 ByVal pObject As PTR, _
                 ByVal pData As PTR) As Long
    Dim cPlugObj    As IVBVstEffect
    Dim tData()     As tAutomationRecord
    
    vbaObjSetAddref cPlugObj, ByVal pObject

    ReqEditIdle = cPlugObj.EditorIdle(tData)
    
    vbaAryMove ByVal pData, ByVal ArrPtr(tData)

End Function

Private Function ReqEditClose( _
                 ByVal pObject As PTR, _
                 ByVal hWndContainer As Handle) As Long
    Dim cPlugObj    As IVBVstEffect

    vbaObjSetAddref cPlugObj, ByVal pObject
    
    SetParent hWndContainer, HWND_MESSAGE

    cPlugObj.EditorClose
    
End Function

Private Function ReqEditOpen( _
                 ByVal pObject As PTR, _
                 ByVal pData As PTR) As Long
    Dim cPlugObj    As IVBVstEffect
    Dim tData       As tEditOpenData
    Dim tRC         As ERect
    
    vbaObjSetAddref cPlugObj, ByVal pObject
    memcpy tData, ByVal pData, Len(tData)
    
    ' // Put container to host container
    If SetParent(tData.hWndContainer, tData.hWndHost) = 0 Then
        Exit Function
    End If
    
    tRC = cPlugObj.EditorRect
    
    MoveWindow tData.hWndContainer, tRC.wLeft, tRC.wTop, tRC.wRight - tRC.wLeft, tRC.wBottom - tRC.wTop, 0
    ShowWindow tData.hWndContainer, SW_SHOW
    
    ReqEditOpen = cPlugObj.EditorOpen(tData.hWndContainer) And 1
    
End Function

Private Function ReqSetSampleRate( _
                 ByVal pObject As PTR, _
                 ByVal lSampleRate As Long) As Long
    Dim cPlugObj    As IVBVstEffect
    Dim fValue      As Single
    
    vbaObjSetAddref cPlugObj, ByVal pObject
    
    GetMem4 lSampleRate, fValue
    
    cPlugObj.SampleRate = fValue
    
End Function

Private Function ReqMainsChanged( _
                 ByVal pObject As PTR, _
                 ByVal lValue As Long) As Long
    Dim cPlugObj    As IVBVstEffect

    vbaObjSetAddref cPlugObj, ByVal pObject
    
    If lValue Then
        cPlugObj.Resume
    Else
        cPlugObj.Suspend
    End If
    
End Function

Private Function ReqGetParamName( _
                 ByVal pObject As PTR, _
                 ByVal pData As PTR) As Long
    Dim cPlugObj    As IVBVstEffect
    Dim tData       As tParamNameLabelData
    Dim sAnsi       As String
    
    vbaObjSetAddref cPlugObj, ByVal pObject
    memcpy tData, ByVal pData, Len(tData)
    
    sAnsi = Unicode2Ansi(cPlugObj.ParamName(tData.lIndex))
    
    If LenB(sAnsi) Then
        lstrcpynA ByVal tData.pName, ByVal StrPtr(sAnsi), kVstMaxParamStrLen
    Else
        PutMem1 ByVal tData.pName, 0
    End If
    
End Function

Private Function ReqGetParamLabel( _
                 ByVal pObject As PTR, _
                 ByVal pData As PTR) As Long
    Dim cPlugObj    As IVBVstEffect
    Dim tData       As tParamNameLabelData
    Dim sAnsi       As String
    
    vbaObjSetAddref cPlugObj, ByVal pObject
    memcpy tData, ByVal pData, Len(tData)

    sAnsi = Unicode2Ansi(cPlugObj.ParamLabel(tData.lIndex))
    
    If LenB(sAnsi) Then
        lstrcpynA ByVal tData.pName, ByVal StrPtr(sAnsi), kVstMaxParamStrLen
    Else
        PutMem1 ByVal tData.pName, 0
    End If
    
End Function

Private Function ReqGetParamDisplay( _
                 ByVal pObject As PTR, _
                 ByVal pData As PTR) As Long
    Dim cPlugObj    As IVBVstEffect
    Dim tData       As tParamNameLabelData
    Dim sAnsi       As String
    
    vbaObjSetAddref cPlugObj, ByVal pObject
    memcpy tData, ByVal pData, Len(tData)
    
    sAnsi = Unicode2Ansi(cPlugObj.ParamDisplay(tData.lIndex))
    
    If LenB(sAnsi) Then
        lstrcpynA ByVal tData.pName, ByVal StrPtr(sAnsi), kVstMaxParamStrLen
    Else
        PutMem1 ByVal tData.pName, 0
    End If
    
End Function

Private Function ReqGetParameter( _
                 ByVal pObject As PTR, _
                 ByVal lIndex As Long) As Long
    Dim cPlugObj    As IVBVstEffect

    vbaObjSetAddref cPlugObj, ByVal pObject

    GetMem4 cPlugObj.ParamValue(lIndex), ReqGetParameter
    
End Function

Private Function ReqSetParameter( _
                 ByVal pObject As PTR, _
                 ByVal pData As PTR) As Long
    Dim cPlugObj    As IVBVstEffect
    Dim tData       As tParamValueData
    
    vbaObjSetAddref cPlugObj, ByVal pObject
    memcpy tData, ByVal pData, Len(tData)
    
    cPlugObj.ParamValue(tData.lIndex) = tData.fValue
    
End Function

Private Function ReqProcessOrProcessReplacing( _
                 ByVal pObject As PTR, _
                 ByVal pData As PTR, _
                 ByVal bReplacing As Boolean) As Long
    Dim cPlugObj    As IVBVstEffect
    Dim tProcData   As tProcessSamplesData
    
    On Error GoTo error_handler
    
    vbaObjSetAddref cPlugObj, ByVal pObject
    
    memcpy tProcData, ByVal pData, Len(tProcData)
    
    If bReplacing Then
        cPlugObj.ProcessReplacing tProcData.pIn, tProcData.pOut, tProcData.lSamples
    Else
        cPlugObj.Process tProcData.pIn, tProcData.pOut, tProcData.lSamples
    End If
    
error_handler:
    
    memset tProcData, Len(tProcData), 0
    
End Function

' // Create plugin object
Private Function ReqCreatePlugin( _
                 ByVal pDispatcher As PTR) As Long
    Dim hr              As Long
    Dim cPluginObj      As IVBVstEffect
    Dim tDispatcher     As CVBVstDispatcher
    Dim hWndContainer   As Handle
    Dim pfnDGCO         As PTR
    Dim tArgs           As tDispCallFuncArgData
    Dim pClassFactory   As PTR
    
    memcpy tDispatcher, ByVal pDispatcher, LenB(tDispatcher)

    ' // Create class instance
    pfnDGCO = GetProcAddress(m_hInstance, "DllGetClassObject")
    If pfnDGCO = NULL_PTR Then
        hr = HRESULTFromWin32(GetLastError)
        GoTo CleanUp
    End If
    
    With tArgs
        
        ' // Call DllGetClassObject
        .iTypes(0) = vbString:  .iTypes(1) = vbString:  .iTypes(2) = vbString
        .pArgs(0) = VarPtr(.lVarArgs(0)):  .pArgs(1) = VarPtr(.lVarArgs(1)): .pArgs(2) = VarPtr(.lVarArgs(2)):
        
        .lVarArgs(0).iVT = vbString: .lVarArgs(0).pData0 = VarPtr(m_tVSTClsID)
        .lVarArgs(1).iVT = vbString: .lVarArgs(1).pData0 = VarPtr(IID_IClassFactory)
        .lVarArgs(2).iVT = vbString: .lVarArgs(2).pData0 = VarPtr(pClassFactory)
        
        hr = DispCallFunc(ByVal NULL_PTR, pfnDGCO, CC_STDCALL, vbLong, 3, .iTypes(0), .pArgs(0), .lVarRet)
        If hr < 0 Then
            GoTo CleanUp
        ElseIf .lVarRet.pData0 < 0 Then
            hr = .lVarRet.pData0
            GoTo CleanUp
        End If
        
        ' // Call IClassFactory::CreateInstance
        .lVarArgs(0).pData0 = NULL_PTR
        .lVarArgs(1).pData0 = VarPtr(IID_IVBVstEffect)
        .lVarArgs(2).pData0 = VarPtr(cPluginObj)
        
        hr = DispCallFunc(ByVal pClassFactory, 3 * Len(pDispatcher), CC_STDCALL, vbString, 3, .iTypes(0), .pArgs(0), .lVarRet)
        If hr < 0 Then
            GoTo CleanUp
        ElseIf .lVarRet.pData0 < 0 Then
            hr = .lVarRet.pData0
            GoTo CleanUp
        End If
        
    End With

    ' // Create container window
    hWndContainer = CreateWindowEx(WS_EX_NOPARENTNOTIFY, CONTAINER_WND_CLASS, vbNullString, _
                                   WS_CLIPCHILDREN Or WS_CLIPSIBLINGS Or WS_CHILD, 0, 0, 0, 0, _
                                   HWND_MESSAGE, 0, m_hInstance, ByVal NULL_PTR)
    
    If hWndContainer = 0 Then
        hr = HRESULTFromWin32(GetLastError)
        GoTo CleanUp
    End If
    
CleanUp:

    If pClassFactory Then
        vbaObjSetAddref pClassFactory, ByVal NULL_PTR
    End If
    
    If hr < 0 Then
        If hWndContainer Then
            DestroyWindow hWndContainer
        End If
    Else
    
        With tDispatcher
            
            cPluginObj.AudioMasterCallback = .pHostCallback
            cPluginObj.AEffectPtr = .pAEffect
            
            vbaObjSetAddref .pPlugObjPtr, ByVal ObjPtr(cPluginObj)
            
            .lVstVersion = cPluginObj.VstVersion
            .lNumOfInputs = cPluginObj.NumOfInputs
            .lNumOfOutputs = cPluginObj.NumOfOutputs
            .lNumOfParams = cPluginObj.NumOfParam

            If Not cPluginObj.ParameterProperties(.pParamProp) Then
                .pParamProp = NULL_PTR
            End If
            
            .lUniqueId = cPluginObj.UniqueId
            .lVersion = cPluginObj.Version
            .bHasEditor = cPluginObj.HasEditor
            .bProgAreChunks = cPluginObj.ProgramsAreChunks
            .bSupportsEvents = cPluginObj.SupportsVSTEvents
            .lNumOfPrograms = cPluginObj.NumOfPrograms
            .bCanMono = cPluginObj.CanMono
            .lPlugCategory = cPluginObj.PlugCategory
            .hWndContainer = hWndContainer
            .tRect = cPluginObj.EditorRect
            
        End With
    
    End If
    
    memcpy ByVal pDispatcher, tDispatcher, LenB(tDispatcher)
    
    ReqCreatePlugin = hr
    
End Function

Private Function ReqReleaseRef( _
                 ByVal hWndContainer As Handle, _
                 ByVal pObject As PTR) As Long
                 
    vbaObjSetAddref pObject, ByVal NULL_PTR
    DestroyWindow hWndContainer
    
End Function

Private Function ReqExitThread( _
                 ByVal hWnd As Handle) As Long
    DestroyWindow hWnd
    PostQuitMessage 0
End Function

'  ^  ^  ^  ^  ^  ^  ^  ^  ^  ^  ^  ^  ^  ^  ^
'  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |

' // These functions are called in STA thread  //

Private Function FAR_PROC( _
                 ByVal pfn As PTR) As PTR
    FAR_PROC = pfn
End Function

Private Function HRESULTFromWin32( _
                 ByVal lError As Long) As Long
    HRESULTFromWin32 = &H80070000 Or (lError And &HFFFF&)
End Function

Private Function IID_IVBVstEffect() As UUID
    PutMem8 IID_IVBVstEffect.Data1, &HA41826AA, &H431FAB40
    PutMem8 IID_IVBVstEffect.Data4(0), &H5A1775A9, &H1E9A86FB
End Function

Private Function IID_IClassFactory() As UUID
    IID_IClassFactory.Data1 = 1
    PutMem8 IID_IClassFactory.Data4(0), &HC0, &H46000000
End Function

Private Function Unicode2Ansi( _
                 ByVal psz As String) As String
    vbaStrToAnsi Unicode2Ansi, psz
End Function

Private Function Ansi2Unicode( _
                 ByVal psz As PTR) As String
    Dim lSize   As Long
    
    lSize = MultiByteToWideChar(CP_ACP, 0, ByVal psz, -1, ByVal NULL_PTR, 0)
    If lSize = 0 Then
        Exit Function
    End If
    
    Ansi2Unicode = Space$(lSize - 1)
    
    MultiByteToWideChar CP_ACP, 0, ByVal psz, -1, ByVal StrPtr(Ansi2Unicode), lSize
    
End Function
