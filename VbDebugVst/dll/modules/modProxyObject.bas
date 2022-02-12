Attribute VB_Name = "modProxyObject"
' //
' // modProxyObject.bas - proxy object (IVBVstEvent) implementation
' // By The trick, 2022
' //

Option Explicit

Private Type IVBVstEvent_VTable
    pfnQueryInterface           As PTR
    pfnAddRef                   As PTR
    pfnRelease                  As PTR
    pfnVstVersion               As PTR
    pfnNumOfParam               As PTR
    pfnNumOfInputs              As PTR
    pfnNumOfOutputs             As PTR
    pfnParameterProperties      As PTR
    pfnVersion                  As PTR
    pfnUniqueId                 As PTR
    pfnParamName                As PTR
    pfnParamLabel               As PTR
    pfnParamDisplay             As PTR
    pfnParamValue_put           As PTR
    pfnParamValue_get           As PTR
    pfnSuspend                  As PTR
    pfnResume                   As PTR
    pfnSampleRate_put           As PTR
    pfnSampleRate_get           As PTR
    pfnProcess                  As PTR
    pfnProcessReplacing         As PTR
    pfnEditorRect               As PTR
    pfnEditorOpen               As PTR
    pfnEditorClose              As PTR
    pfnEditorIdle               As PTR
    pfnHasEditor                As PTR
    pfnProgramsAreChunks        As PTR
    pfnSupportsVSTEvents        As PTR
    pfnCanDo                    As PTR
    pfnAudioMasterCallback_put  As PTR
    pfnAudioMasterCallback_get  As PTR
    pfnAEffectPtr_put           As PTR
    pfnAEffectPtr_get           As PTR
    pfnCanParameterBeAutomated  As PTR
    pfnTailSize                 As PTR
    pfnNumOfPrograms            As PTR
    pfnProgramName_get          As PTR
    pfnProgramName_put          As PTR
    pfnProgram_get              As PTR
    pfnProgram_put              As PTR
    pfnCanMono                  As PTR
    pfnBlockSize                As PTR
    pfnProgramNameIndexed       As PTR
    pfnCopyProgram              As PTR
    pfnEffectName               As PTR
    pfnVendorString             As PTR
    pfnProductString            As PTR
    pfnGetStateChunk            As PTR
    pfnSetStateChunk            As PTR
    pfnVendorVersion            As PTR
    pfnVendorSpecific           As PTR
    pfnProcessEvents            As PTR
    pfnPlugCategory             As PTR
    pfnSetBypass                As PTR
    pfnStartProcess             As PTR
    pfnStopProcess              As PTR
    pfnThreadId                 As PTR
End Type

Private m_pParameterProperties  As PTR
Private m_pStateChunk           As PTR
Private m_lNumOfInputs          As Long
Private m_lNumOfOutputs         As Long
Private m_tVTable               As IVBVstEvent_VTable
Private m_pProxyObject          As PTR

' // Release plugin data in this process
Public Sub ReleaseProxyData()
    
    If m_pParameterProperties Then
        CoTaskMemFree m_pParameterProperties
    End If
    
    If m_pStateChunk Then
        CoTaskMemFree m_pStateChunk
    End If
    
    m_pParameterProperties = NULL_PTR
    m_pStateChunk = NULL_PTR
    
End Sub

' // There is only single instance. Ref counter isn't used because this is global object
Public Function GetProxyObject() As PTR

    If m_tVTable.pfnQueryInterface = NULL_PTR Then
        
        With m_tVTable
            
            .pfnQueryInterface = FAR_PROC(AddressOf CVstProxy_QueryInterface)
            .pfnAddRef = FAR_PROC(AddressOf CVstProxy_AddRef)
            .pfnRelease = FAR_PROC(AddressOf CVstProxy_Release)
            .pfnVstVersion = FAR_PROC(AddressOf CVstProxy_VstVersion)
            .pfnNumOfParam = FAR_PROC(AddressOf CVstProxy_NumOfParam)
            .pfnNumOfInputs = FAR_PROC(AddressOf CVstProxy_NumOfInputs)
            .pfnNumOfOutputs = FAR_PROC(AddressOf CVstProxy_NumOfOutputs)
            .pfnParameterProperties = FAR_PROC(AddressOf CVstProxy_ParameterProperties)
            .pfnVersion = FAR_PROC(AddressOf CVstProxy_Version)
            .pfnUniqueId = FAR_PROC(AddressOf CVstProxy_UniqueId)
            .pfnParamName = FAR_PROC(AddressOf CVstProxy_ParamName)
            .pfnParamLabel = FAR_PROC(AddressOf CVstProxy_ParamLabel)
            .pfnParamDisplay = FAR_PROC(AddressOf CVstProxy_ParamDisplay)
            .pfnParamValue_put = FAR_PROC(AddressOf CVstProxy_ParamValue_put)
            .pfnParamValue_get = FAR_PROC(AddressOf CVstProxy_ParamValue_get)
            .pfnSuspend = FAR_PROC(AddressOf CVstProxy_Suspend)
            .pfnResume = FAR_PROC(AddressOf CVstProxy_Resume)
            .pfnSampleRate_put = FAR_PROC(AddressOf CVstProxy_SampleRate_put)
            .pfnSampleRate_get = FAR_PROC(AddressOf CVstProxy_SampleRate_get)
            .pfnProcess = FAR_PROC(AddressOf CVstProxy_Process)
            .pfnProcessReplacing = FAR_PROC(AddressOf CVstProxy_ProcessReplacing)
            .pfnEditorRect = FAR_PROC(AddressOf CVstProxy_EditorRect)
            .pfnEditorOpen = FAR_PROC(AddressOf CVstProxy_EditorOpen)
            .pfnEditorClose = FAR_PROC(AddressOf CVstProxy_EditorClose)
            .pfnEditorIdle = FAR_PROC(AddressOf CVstProxy_EditorIdle)
            .pfnHasEditor = FAR_PROC(AddressOf CVstProxy_HasEditor)
            .pfnProgramsAreChunks = FAR_PROC(AddressOf CVstProxy_ProgramsAreChunks)
            .pfnSupportsVSTEvents = FAR_PROC(AddressOf CVstProxy_SupportsVSTEvents)
            .pfnCanDo = FAR_PROC(AddressOf CVstProxy_CanDo)
            .pfnAudioMasterCallback_put = FAR_PROC(AddressOf CVstProxy_AudioMasterCallback_put)
            .pfnAudioMasterCallback_get = FAR_PROC(AddressOf CVstProxy_AudioMasterCallback_get)
            .pfnAEffectPtr_put = FAR_PROC(AddressOf CVstProxy_AEffectPtr_put)
            .pfnAEffectPtr_get = FAR_PROC(AddressOf CVstProxy_AEffectPtr_get)
            .pfnCanParameterBeAutomated = FAR_PROC(AddressOf CVstProxy_CanParameterBeAutomated)
            .pfnTailSize = FAR_PROC(AddressOf CVstProxy_TailSize)
            .pfnNumOfPrograms = FAR_PROC(AddressOf CVstProxy_NumOfPrograms)
            .pfnProgramName_get = FAR_PROC(AddressOf CVstProxy_ProgramName_get)
            .pfnProgramName_put = FAR_PROC(AddressOf CVstProxy_ProgramName_put)
            .pfnProgram_get = FAR_PROC(AddressOf CVstProxy_Program_get)
            .pfnProgram_put = FAR_PROC(AddressOf CVstProxy_Program_put)
            .pfnCanMono = FAR_PROC(AddressOf CVstProxy_CanMono)
            .pfnBlockSize = FAR_PROC(AddressOf CVstProxy_BlockSize)
            .pfnProgramNameIndexed = FAR_PROC(AddressOf CVstProxy_ProgramNameIndexed)
            .pfnCopyProgram = FAR_PROC(AddressOf CVstProxy_CopyProgram)
            .pfnEffectName = FAR_PROC(AddressOf CVstProxy_EffectName)
            .pfnVendorString = FAR_PROC(AddressOf CVstProxy_VendorString)
            .pfnProductString = FAR_PROC(AddressOf CVstProxy_ProductString)
            .pfnGetStateChunk = FAR_PROC(AddressOf CVstProxy_GetStateChunk)
            .pfnSetStateChunk = FAR_PROC(AddressOf CVstProxy_SetStateChunk)
            .pfnVendorVersion = FAR_PROC(AddressOf CVstProxy_VendorVersion)
            .pfnVendorSpecific = FAR_PROC(AddressOf CVstProxy_VendorSpecific)
            .pfnProcessEvents = FAR_PROC(AddressOf CVstProxy_ProcessEvents)
            .pfnPlugCategory = FAR_PROC(AddressOf CVstProxy_PlugCategory)
            .pfnSetBypass = FAR_PROC(AddressOf CVstProxy_SetBypass)
            .pfnStartProcess = FAR_PROC(AddressOf CVstProxy_StartProcess)
            .pfnStopProcess = FAR_PROC(AddressOf CVstProxy_StopProcess)
            .pfnThreadId = FAR_PROC(AddressOf CVstProxy_ThreadId)

        End With
        
        m_pProxyObject = VarPtr(m_tVTable)
        
    End If
    
    m_lNumOfInputs = 0
    m_lNumOfOutputs = 0
    
    GetProxyObject = VarPtr(m_pProxyObject)
    
End Function

Private Function CVstProxy_QueryInterface( _
                 ByVal pObj As PTR, _
                 ByRef tIID As UUID, _
                 ByRef ppObj As PTR) As Long
    ' // Lazy for IUnknown
    If IsEqualGUID(tIID, IID_IVBVstEffect) = 0 Then
        ppObj = pObj
        CVstProxy_AddRef pObj
    Else
        ppObj = NULL_PTR
        CVstProxy_QueryInterface = E_NOINTERFACE
    End If
End Function

Private Function CVstProxy_AddRef( _
                 ByVal pObj As PTR) As Long
    CVstProxy_AddRef = 1
End Function

Private Function CVstProxy_Release( _
                 ByVal pObj As PTR) As Long
    CVstProxy_Release = 1
End Function

'    long VstVersion([in, out] long* pValue);
Private Function CVstProxy_VstVersion( _
                 ByVal pObj As PTR, _
                 ByRef lRet As Long) As Long
    CVstProxy_VstVersion = SendRequestGeneric(WM_VSTVERSION, 0, ByVal NULL_PTR)
    lRet = g_tSharedData(0).lArg1
End Function

'    long NumOfParam([in, out] long* pValue);
Private Function CVstProxy_NumOfParam( _
                 ByVal pObj As PTR, _
                 ByRef lRet As Long) As Long
    CVstProxy_NumOfParam = SendRequestGeneric(WM_GETPARAMETERSCOUNT, 0, ByVal NULL_PTR)
    lRet = g_tSharedData(0).lArg1
End Function

'    long NumOfInputs([in, out] long* pValue);
Private Function CVstProxy_NumOfInputs( _
                 ByVal pObj As PTR, _
                 ByRef lRet As Long) As Long
                 
    CVstProxy_NumOfInputs = SendRequestGeneric(WM_NUMOFINPUTS, 0, ByVal NULL_PTR)
    lRet = g_tSharedData(0).lArg1
    m_lNumOfInputs = g_tSharedData(0).lArg1
    
End Function

'    long NumOfOutputs([in, out] long* pValue);
Private Function CVstProxy_NumOfOutputs( _
                 ByVal pObj As PTR, _
                 ByRef lRet As Long) As Long
                 
    CVstProxy_NumOfOutputs = SendRequestGeneric(WM_NUMOFOUTPUTS, 0, ByVal NULL_PTR)
    lRet = g_tSharedData(0).lArg1
    m_lNumOfOutputs = g_tSharedData(0).lArg1
    
End Function

'    long ParameterProperties([in, out] PTR* pProperties, [in, out] BOOLEAN* pValue);
Private Function CVstProxy_ParameterProperties( _
                 ByVal pObj As PTR, _
                 ByRef pProperties As PTR, _
                 ByRef bRet As Boolean) As Long
    Dim hr      As Long
    Dim lParams As Long
    
    pProperties = NULL_PTR
    
    hr = SendRequestGeneric(WM_PARAMETERPROPERTIES, 0, ByVal NULL_PTR)
    If hr >= 0 Then
        
        bRet = g_tSharedData(0).lArg1
        
        If bRet Then
            hr = CVstProxy_NumOfParam(pObj, lParams)
            If hr >= 0 Then
                If lParams > 0 Then
                    m_pParameterProperties = CoTaskMemRealloc(m_pParameterProperties, lParams * &H98)
                    If m_pParameterProperties Then
                        
                        If ReadProcessMemory(g_hRemoteProcess, g_tSharedData(0).pArg1, ByVal m_pParameterProperties, _
                                             lParams * &H98, 0) = 0 Then
                            hr = HRESULTFromWin32(GetLastError)
                        Else
                            pProperties = m_pParameterProperties
                        End If

                    Else
                        hr = E_OUTOFMEMORY
                    End If
                End If
            End If
        End If
        
    End If
    
    CVstProxy_ParameterProperties = hr
    
End Function

'    long Version([in, out] long* pValue);
Private Function CVstProxy_Version( _
                 ByVal pObj As PTR, _
                 ByRef lRet As Long) As Long
    CVstProxy_Version = SendRequestGeneric(WM_PLUGVERSION, 0, ByVal NULL_PTR)
    lRet = g_tSharedData(0).lArg1
End Function

'    long UniqueId([in, out] long* pValue);
Private Function CVstProxy_UniqueId( _
                 ByVal pObj As PTR, _
                 ByRef lRet As Long) As Long
    CVstProxy_UniqueId = SendRequestGeneric(WM_UNIQUEID, 0, ByVal NULL_PTR)
    lRet = g_tSharedData(0).lArg1
End Function

'    long ParamName([in] long lIndex, [in, out] BSTR* pOut);
Private Function CVstProxy_ParamName( _
                 ByVal pObj As PTR, _
                 ByVal lIndex As Long, _
                 ByRef sRet As String) As Long
    CVstProxy_ParamName = SendRequestGeneric(WM_GETPARAMNAME, lIndex, ByVal NULL_PTR)
    sRet = SysAllocString(ByVal g_tSharedData(0).pDataBuffer)
End Function

'    long ParamLabel([in] long lIndex, [in, out] BSTR* pOut);
Private Function CVstProxy_ParamLabel( _
                 ByVal pObj As PTR, _
                 ByVal lIndex As Long, _
                 ByRef sRet As String) As Long
    CVstProxy_ParamLabel = SendRequestGeneric(WM_GETPARAMLABEL, lIndex, ByVal NULL_PTR)
    sRet = SysAllocString(ByVal g_tSharedData(0).pDataBuffer)
End Function

'    long ParamDisplay([in] long lIndex, [in, out] BSTR* pOut);
Private Function CVstProxy_ParamDisplay( _
                 ByVal pObj As PTR, _
                 ByVal lIndex As Long, _
                 ByRef sRet As String) As Long
    CVstProxy_ParamDisplay = SendRequestGeneric(WM_GETPARAMDISPLAY, lIndex, ByVal NULL_PTR)
    sRet = SysAllocString(ByVal g_tSharedData(0).pDataBuffer)
End Function

'    long ParamValue_put([in] long lIndex, [in] float fValue);
Private Function CVstProxy_ParamValue_put( _
                 ByVal pObj As PTR, _
                 ByVal lIndex As Long, _
                 ByVal fValue As Single) As Long
    g_tSharedData(0).fArg1 = fValue
    CVstProxy_ParamValue_put = SendRequestGeneric(WM_SETPARAMETER, lIndex, ByVal NULL_PTR)
End Function

'    long ParamValue_get([in] long lIndex, [in, out] float* pOut);
Private Function CVstProxy_ParamValue_get( _
                 ByVal pObj As PTR, _
                 ByVal lIndex As Long, _
                 ByRef fValue As Single) As Long
    CVstProxy_ParamValue_get = SendRequestGeneric(WM_GETPARAMETER, lIndex, ByVal NULL_PTR)
    fValue = g_tSharedData(0).fArg1
End Function

'    long Suspend();
Private Function CVstProxy_Suspend( _
                 ByVal pObj As PTR) As Long
    CVstProxy_Suspend = SendRequestGeneric(WM_MAINSCHANGED, 0, ByVal NULL_PTR)
End Function

'    long Resume();
Private Function CVstProxy_Resume( _
                 ByVal pObj As PTR) As Long
    CVstProxy_Resume = SendRequestGeneric(WM_MAINSCHANGED, 1, ByVal NULL_PTR)
End Function

'    long SampleRate_put([in] float lValue);
Private Function CVstProxy_SampleRate_put( _
                 ByVal pObj As PTR, _
                 ByVal fValue As Single) As Long
    g_tSharedData(0).fArg1 = fValue
    CVstProxy_SampleRate_put = SendRequestGeneric(WM_SETSAMPLERATE, 0, ByVal NULL_PTR)
End Function

'    long SampleRate_get([in, out] float* pOut);
Private Function CVstProxy_SampleRate_get( _
                 ByVal pObj As PTR, _
                 ByRef fValue As Single) As Long
    CVstProxy_SampleRate_get = E_FAIL
End Function

'    long Process([in] PTR pInput, [in] PTR pOutput, [in] long sampleFrames);
Private Function CVstProxy_Process( _
                 ByVal pObj As PTR, _
                 ByVal pInput As PTR, _
                 ByVal pOutput As PTR, _
                 ByVal lSampleFrames As Long) As Long
    Dim hr      As Long
    Dim lIndex  As Long
    Dim pChData As PTR
    
    With g_tSharedData(0)
        
        If m_lNumOfInputs <= 0 Then
            hr = CVstProxy_NumOfInputs(pObj, m_lNumOfInputs)
            If hr < 0 Or m_lNumOfInputs <= 0 Then
                GoTo exit_proc
            End If
        ElseIf m_lNumOfOutputs <= 0 Then
            hr = CVstProxy_NumOfOutputs(pObj, m_lNumOfOutputs)
            If hr < 0 Or m_lNumOfOutputs <= 0 Then
                GoTo exit_proc
            End If
        End If
        
        For lIndex = 0 To m_lNumOfInputs - 1
            
            ' // Validate
            GetMemPtr ByVal pInput + lIndex * Len(pChData), pChData
            
            If pChData < .pSamplesBuf Or (pChData + lSampleFrames * 4) > .pSamplesBuf + SAMPLES_BUFFER_SIZE * 4 Then
                hr = E_UNEXPECTED
                GoTo exit_proc
            End If
            
            pChData = pChData - .pSamplesBuf + .pSamplesBufRemote
            
            PutMem4 ByVal .pDataBuffer + lIndex * Len(pChData), pChData
            
        Next
        
        For lIndex = 0 To m_lNumOfOutputs - 1
            
            ' // Validate
            GetMemPtr ByVal pOutput + lIndex * Len(pChData), pChData
            
            If pChData < .pSamplesBuf Or (pChData + lSampleFrames * 4) > .pSamplesBuf + SAMPLES_BUFFER_SIZE * 4 Then
                hr = E_UNEXPECTED
                GoTo exit_proc
            End If
            
            pChData = pChData - .pSamplesBuf + .pSamplesBufRemote
            
            PutMem4 ByVal .pDataBuffer + (lIndex + m_lNumOfInputs) * Len(pChData), pChData
            
        Next
        
        .pArg1 = .pDataBufferRemote
        .pArg2 = .pDataBufferRemote + m_lNumOfInputs * Len(pInput)
        .lArg1 = lSampleFrames
        
        hr = SendRequestGeneric(WM_PROCESS, 0, ByVal NULL_PTR)
        
    End With
    
exit_proc:
    
    CVstProxy_Process = hr
    
End Function

'    long ProcessReplacing([in] PTR pInput, [in] PTR pOutput, [in] long sampleFrames);
Private Function CVstProxy_ProcessReplacing( _
                 ByVal pObj As PTR, _
                 ByVal pInput As PTR, _
                 ByVal pOutput As PTR, _
                 ByVal lSampleFrames As Long) As Long
    Dim hr      As Long
    Dim lIndex  As Long
    Dim pChData As PTR
    
    With g_tSharedData(0)
        
        If m_lNumOfInputs <= 0 Then
            hr = CVstProxy_NumOfInputs(pObj, m_lNumOfInputs)
            If hr < 0 Or m_lNumOfInputs <= 0 Then
                GoTo exit_proc
            End If
        ElseIf m_lNumOfOutputs <= 0 Then
            hr = CVstProxy_NumOfOutputs(pObj, m_lNumOfOutputs)
            If hr < 0 Or m_lNumOfOutputs <= 0 Then
                GoTo exit_proc
            End If
        End If
        
        For lIndex = 0 To m_lNumOfInputs - 1
            
            ' // Validate
            GetMemPtr ByVal pInput + lIndex * Len(pChData), pChData
            
            If pChData < .pSamplesBuf Or (pChData + lSampleFrames * 4) > .pSamplesBuf + SAMPLES_BUFFER_SIZE * 4 Then
                hr = E_UNEXPECTED
                GoTo exit_proc
            End If
            
            pChData = pChData - .pSamplesBuf + .pSamplesBufRemote
            
            PutMem4 ByVal .pDataBuffer + lIndex * Len(pChData), pChData
            
        Next
        
        For lIndex = 0 To m_lNumOfOutputs - 1
            
            ' // Validate
            GetMemPtr ByVal pOutput + lIndex * Len(pChData), pChData
            
            If pChData < .pSamplesBuf Or (pChData + lSampleFrames * 4) > .pSamplesBuf + SAMPLES_BUFFER_SIZE * 4 Then
                hr = E_UNEXPECTED
                GoTo exit_proc
            End If
            
            pChData = pChData - .pSamplesBuf + .pSamplesBufRemote
            
            PutMem4 ByVal .pDataBuffer + (lIndex + m_lNumOfInputs) * Len(pChData), pChData
            
        Next
        
        .pArg1 = .pDataBufferRemote
        .pArg2 = .pDataBufferRemote + m_lNumOfInputs * Len(pChData)
        .lArg1 = lSampleFrames
        
        hr = SendRequestGeneric(WM_PROCESSREPLACING, 0, ByVal NULL_PTR)
        
    End With
    
exit_proc:
    
    CVstProxy_ProcessReplacing = hr
    
End Function

'    long EditorRect([in, out] VBVST2X.ERect* tRet);
Private Function CVstProxy_EditorRect( _
                 ByVal pObj As PTR, _
                 ByRef tRect As ERect) As Long
                 
    CVstProxy_EditorRect = SendRequestGeneric(WM_EDITORRECT, 0, ByVal NULL_PTR)
    
    If CVstProxy_EditorRect >= 0 Then
        memcpy tRect, ByVal g_tSharedData(0).pDataBuffer, Len(tRect)
    End If
    
End Function

'    long EditorOpen([in] HANDLE hWnd, [in, out] BOOLEAN* bResult);
Private Function CVstProxy_EditorOpen( _
                 ByVal pObj As PTR, _
                 ByVal hWnd As Handle, _
                 ByRef pRet As Boolean) As Long
    CVstProxy_EditorOpen = SendRequestGeneric(WM_EDITOPEN, hWnd, ByVal NULL_PTR)
    pRet = g_tSharedData(0).lArg1
End Function

'    long EditorClose();
Private Function CVstProxy_EditorClose( _
                 ByVal pObj As PTR) As Long
    CVstProxy_EditorClose = SendRequestGeneric(WM_EDITCLOSE, 0, ByVal NULL_PTR)
End Function

'    long EditorIdle([in, out] SAFEARRAY(tAutomationRecord) *tRecords, [in, out] long* lNumOfRecords);
Private Function CVstProxy_EditorIdle( _
                 ByVal pObj As PTR, _
                 ByRef tRecords() As tAutomationRecord, _
                 ByRef lNumOfRecords As Long) As Long
                 
    CVstProxy_EditorIdle = SendRequestGeneric(WM_EDITIDLE, 0, ByVal NULL_PTR)
    
    If CVstProxy_EditorIdle >= 0 Then
    
        lNumOfRecords = g_tSharedData(0).lAutomationCount
        
        If lNumOfRecords > 0 Then
            ReDim tRecords(lNumOfRecords - 1)
            memcpy tRecords(0), ByVal g_tSharedData(0).pAutomationBuf, lNumOfRecords * 8
        Else
            Erase tRecords
        End If
        
    End If
    
End Function

'    long HasEditor([in, out] BOOLEAN* pValue);
Private Function CVstProxy_HasEditor( _
                 ByVal pObj As PTR, _
                 ByRef bRet As Boolean) As Long
    CVstProxy_HasEditor = SendRequestGeneric(WM_HASEDITOR, 0, ByVal NULL_PTR)
    bRet = g_tSharedData(0).lArg1
End Function

'    long ProgramsAreChunks([in, out] BOOLEAN *pValue);
Private Function CVstProxy_ProgramsAreChunks( _
                 ByVal pObj As PTR, _
                 ByRef bRet As Boolean) As Long
    CVstProxy_ProgramsAreChunks = SendRequestGeneric(WM_PROGRAMSARECHUNK, 0, ByVal NULL_PTR)
    bRet = g_tSharedData(0).lArg1
End Function

'    long SupportsVSTEvents([in, out] BOOLEAN *pValue);
Private Function CVstProxy_SupportsVSTEvents( _
                 ByVal pObj As PTR, _
                 ByRef bRet As Boolean) As Long
    CVstProxy_SupportsVSTEvents = SendRequestGeneric(WM_SUPPORTSVSTEVENTS, 0, ByVal NULL_PTR)
    bRet = g_tSharedData(0).lArg1
End Function

'    long CanDo([in] BSTR *ppszRequest, [in, out] BOOLEAN* pValue);
Private Function CVstProxy_CanDo( _
                 ByVal pObj As PTR, _
                 ByVal pszRequest As PTR, _
                 ByRef bRet As Boolean) As Long
    lstrcpyn ByVal g_tSharedData(0).pDataBuffer, ByVal pszRequest, DATA_BUFFER_SIZE
    CVstProxy_CanDo = SendRequestGeneric(WM_CANDO, 0, ByVal NULL_PTR)
    bRet = g_tSharedData(0).lArg1
End Function

'    long AudioMasterCallback_put([in] PTR pfn);
Private Function CVstProxy_AudioMasterCallback_put( _
                 ByVal pObj As PTR, _
                 ByVal pfn As PTR) As Long
    CVstProxy_AudioMasterCallback_put = E_UNEXPECTED
End Function

'    long AudioMasterCallback_get([in, out] PTR *pfn);
Private Function CVstProxy_AudioMasterCallback_get( _
                 ByVal pObj As PTR, _
                 ByRef pfn As PTR) As Long
    CVstProxy_AudioMasterCallback_get = E_UNEXPECTED
End Function

'    long AEffectPtr_put([in] PTR pValue);
Private Function CVstProxy_AEffectPtr_put( _
                 ByVal pObj As PTR, _
                 ByVal pfn As PTR) As Long
    CVstProxy_AEffectPtr_put = E_UNEXPECTED
End Function

'    long AEffectPtr_get([in, out] PTR* pValue);
Private Function CVstProxy_AEffectPtr_get( _
                 ByVal pObj As PTR, _
                 ByRef pfn As PTR) As Long
    CVstProxy_AEffectPtr_get = E_UNEXPECTED
End Function

'    long CanParameterBeAutomated([in] long lIndex, [in, out] BOOLEAN* pValue);
Private Function CVstProxy_CanParameterBeAutomated( _
                 ByVal pObj As PTR, _
                 ByVal lIndex As Long, _
                 ByRef bRet As Boolean) As Long
    CVstProxy_CanParameterBeAutomated = SendRequestGeneric(WM_CANBEAUTOMATED, lIndex, ByVal NULL_PTR)
    bRet = g_tSharedData(0).lArg1
End Function

'    long TailSize([in, out] long* pValue);
Private Function CVstProxy_TailSize( _
                 ByVal pObj As PTR, _
                 ByRef lRet As Long) As Long
    CVstProxy_TailSize = SendRequestGeneric(WM_GETTAILSIZE, 0, ByVal NULL_PTR)
    lRet = g_tSharedData(0).lArg1
End Function

'    long NumOfPrograms([in, out] long* pValue);
Private Function CVstProxy_NumOfPrograms( _
                 ByVal pObj As PTR, _
                 ByRef lRet As Long) As Long
    CVstProxy_NumOfPrograms = SendRequestGeneric(WM_NUMOFPROGRAMS, 0, ByVal NULL_PTR)
    lRet = g_tSharedData(0).lArg1
End Function

'    long ProgramName_get([in, out] BSTR* pValue);
Private Function CVstProxy_ProgramName_get( _
                 ByVal pObj As PTR, _
                 ByRef sRet As String) As Long
    CVstProxy_ProgramName_get = SendRequestGeneric(WM_GETPROGRAMNAME, 0, ByVal NULL_PTR)
    sRet = SysAllocString(ByVal g_tSharedData(0).pDataBuffer)
End Function

'    long ProgramName_put([in] BSTR pValue);
Private Function CVstProxy_ProgramName_put( _
                 ByVal pObj As PTR, _
                 ByVal sRet As String) As Long
    lstrcpyn ByVal g_tSharedData(0).pDataBuffer, ByVal StrPtr(sRet), DATA_BUFFER_SIZE
    CVstProxy_ProgramName_put = SendRequestGeneric(WM_SETPROGRAMNAME, 0, ByVal NULL_PTR)
End Function

'    long Program_get([in, out] long* pValue);
Private Function CVstProxy_Program_get( _
                 ByVal pObj As PTR, _
                 ByRef lRet As Long) As Long
    CVstProxy_Program_get = SendRequestGeneric(WM_GETPROGRAM, 0, ByVal NULL_PTR)
    lRet = g_tSharedData(0).lArg1
End Function

'    long Program_put([in] long lValue);
Private Function CVstProxy_Program_put( _
                 ByVal pObj As PTR, _
                 ByVal lRet As Long) As Long
    CVstProxy_Program_put = SendRequestGeneric(WM_SETPROGRAM, lRet, ByVal NULL_PTR)
End Function

'    long CanMono([in, out] BOOLEAN* pValue);
Private Function CVstProxy_CanMono( _
                 ByVal pObj As PTR, _
                 ByRef bRet As Boolean) As Long
    CVstProxy_CanMono = SendRequestGeneric(WM_CANMONO, 0, ByVal NULL_PTR)
    bRet = g_tSharedData(0).lArg1
End Function

'    long BlockSize([in] long lValue);
Private Function CVstProxy_BlockSize( _
                 ByVal pObj As PTR, _
                 ByVal lValue As Long) As Long
    CVstProxy_BlockSize = SendRequestGeneric(WM_SETBLOCKSIZE, lValue, ByVal NULL_PTR)
End Function

'    long ProgramNameIndexed([in] long lCategory, [in] long lIndex, [in, out] BSTR* pValue);
Private Function CVstProxy_ProgramNameIndexed( _
                 ByVal pObj As PTR, _
                 ByVal lIndex As Long, _
                 ByRef sRet As String) As Long
    CVstProxy_ProgramNameIndexed = SendRequestGeneric(WM_GETPROGRAMNAMEINDEXED, lIndex, ByVal NULL_PTR)
    sRet = SysAllocString(ByVal g_tSharedData(0).pDataBuffer)
End Function

'    long CopyProgram([in] long lDestination);
Private Function CVstProxy_CopyProgram( _
                 ByVal pObj As PTR, _
                 ByVal lIndex As Long) As Long
    CVstProxy_CopyProgram = SendRequestGeneric(WM_COPYPROGRAM, lIndex, ByVal NULL_PTR)
End Function

'    long EffectName([in, out] BSTR* pValue);
Private Function CVstProxy_EffectName( _
                 ByVal pObj As PTR, _
                 ByRef sRet As String) As Long
    CVstProxy_EffectName = SendRequestGeneric(WM_GETEFFECTNAME, 0, ByVal NULL_PTR)
    sRet = SysAllocString(ByVal g_tSharedData(0).pDataBuffer)
End Function

'    long VendorString([in, out] BSTR* pValue);
Private Function CVstProxy_VendorString( _
                 ByVal pObj As PTR, _
                 ByRef sRet As String) As Long
    CVstProxy_VendorString = SendRequestGeneric(WM_GETVENDORSTRING, 0, ByVal NULL_PTR)
    sRet = SysAllocString(ByVal g_tSharedData(0).pDataBuffer)
End Function

'    long ProductString([in, out] BSTR* pValue);
Private Function CVstProxy_ProductString( _
                 ByVal pObj As PTR, _
                 ByRef sRet As String) As Long
    CVstProxy_ProductString = SendRequestGeneric(WM_GETPRODUCTSTRING, 0, ByVal NULL_PTR)
    sRet = SysAllocString(ByVal g_tSharedData(0).pDataBuffer)
End Function

'    long GetStateChunk([in]BOOLEAN bIsPreset, [in, out] PTR *pData, [in, out] long *pSize);
Private Function CVstProxy_GetStateChunk( _
                 ByVal pObj As PTR, _
                 ByVal bIsPreset As Boolean, _
                 ByRef ppData As PTR, _
                 ByRef lSize As Long) As Long
    Dim hr  As Long
    
    hr = SendRequestGeneric(WM_GETCHUNK, bIsPreset, ByVal NULL_PTR)
    
    If hr >= 0 Then
        
        lSize = g_tSharedData(0).lArg1
        
        If lSize > 0 Then
            
            m_pStateChunk = CoTaskMemRealloc(m_pStateChunk, lSize)
            If m_pStateChunk = NULL_PTR Then
                hr = E_OUTOFMEMORY
            Else
                If ReadProcessMemory(g_hRemoteProcess, g_tSharedData(0).pArg1, ByVal m_pStateChunk, lSize, 0) = 0 Then
                    hr = HRESULTFromWin32(GetLastError)
                Else
                    ppData = m_pStateChunk
                End If
            End If
            
        Else
            ppData = NULL_PTR
        End If
        
    End If
    
    CVstProxy_GetStateChunk = hr
    
End Function

'    long SetStateChunk([in]BOOLEAN bIsPreset, [in] PTR pData, [in] long lSize, [in, out] BOOLEAN *bRet);
Private Function CVstProxy_SetStateChunk( _
                 ByVal pObj As PTR, _
                 ByVal bIsPreset As Boolean, _
                 ByVal pData As PTR, _
                 ByVal lSize As Long, _
                 ByRef bRet As Boolean) As Long
    Dim hr      As Long
    Dim pMemory As PTR
    
    If lSize > 0 Then
        
        pMemory = VirtualAllocEx(g_hRemoteProcess, NULL_PTR, lSize, MEM_COMMIT Or MEM_RESERVE, PAGE_READWRITE)
        If pMemory = NULL_PTR Then
            hr = E_OUTOFMEMORY
            GoTo exit_proc
        End If
        
        If WriteProcessMemory(g_hRemoteProcess, pMemory, ByVal pData, lSize, 0) = 0 Then
            hr = HRESULTFromWin32(GetLastError)
            GoTo exit_proc
        End If
        
    End If
    
    g_tSharedData(0).pArg1 = pMemory
    
    hr = SendRequestGeneric(WM_SETCHUNK, bIsPreset, ByVal lSize)
    
    bRet = g_tSharedData(0).lArg1
    
exit_proc:
    
    If pMemory Then
        VirtualFreeEx g_hRemoteProcess, pMemory, 0, MEM_RELEASE
    End If
    
    CVstProxy_SetStateChunk = hr
    
End Function

'    long VendorVersion([in, out] long* pValue);
Private Function CVstProxy_VendorVersion( _
                 ByVal pObj As PTR, _
                 ByRef lRet As Long) As Long
    CVstProxy_VendorVersion = SendRequestGeneric(WM_VENDORVERSION, 0, ByVal NULL_PTR)
    lRet = g_tSharedData(0).lArg1
End Function

'    long VendorSpecific([in] long lArg1, [in] long lArg2, [in] PTR lpPtr, [in] float fArg3);
Private Function CVstProxy_VendorSpecific( _
                 ByVal pObj As PTR, _
                 ByVal lArg1 As Long, _
                 ByVal lArg2 As Long, _
                 ByVal lpPtr As PTR, _
                 ByVal fArg3 As Single) As Long
    
    With g_tSharedData(0)
    
        .lArg1 = lArg1
        .lArg2 = lArg2
        .pArg1 = lpPtr
        .fArg1 = fArg3
        CVstProxy_VendorSpecific = SendRequestGeneric(WM_VENDORSPECIFIC, 0, ByVal NULL_PTR)
        
    End With
    
End Function

'    long ProcessEvents([in] VstEvents* pEvents, [in, out] BOOLEAN *pRet);
Private Function CVstProxy_ProcessEvents( _
                 ByVal pObj As PTR, _
                 ByRef tEvents As VstEvents, _
                 ByRef bRet As Boolean) As Long
        
    With g_tSharedData(0)
    
        If tEvents.pEvents < .pEventsBuf Or (tEvents.pEvents + tEvents.numEvents * &H20) > _
                                             .pEventsBuf + EVENTS_BUFFER_SIZE * &H20 Then
            CVstProxy_ProcessEvents = E_UNEXPECTED
        Else
        
            .pArg1 = tEvents.pEvents - .pEventsBuf + .pEventsBufRemote
            .lArg1 = tEvents.numEvents
            
            CVstProxy_ProcessEvents = SendRequestGeneric(WM_PROCESSEVENTS, 0, ByVal NULL_PTR)
            
            bRet = .lArg1
            
        End If
    
    End With

End Function

'    long PlugCategory([in, out] long* pValue);
Private Function CVstProxy_PlugCategory( _
                 ByVal pObj As PTR, _
                 ByRef lRet As Long) As Long
    CVstProxy_PlugCategory = SendRequestGeneric(WM_PLUGCATEGORY, 0, ByVal NULL_PTR)
    lRet = g_tSharedData(0).lArg1
End Function

'    long SetBypass([in] BOOLEAN bValue, [in, out] BOOLEAN *pSupports);
Private Function CVstProxy_SetBypass( _
                 ByVal pObj As PTR, _
                 ByVal bValue As Boolean, _
                 ByRef bRet As Boolean) As Long
    CVstProxy_SetBypass = SendRequestGeneric(WM_SETBYPASS, bValue, ByVal NULL_PTR)
    bRet = g_tSharedData(0).lArg1
End Function

'    long StartProcess([in, out] BOOLEAN *pRet);
Private Function CVstProxy_StartProcess( _
                 ByVal pObj As PTR, _
                 ByRef bRet As Boolean) As Long
    CVstProxy_StartProcess = SendRequestGeneric(WM_STARTPROCESS, 0, ByVal NULL_PTR)
    bRet = g_tSharedData(0).lArg1
End Function

'    long StopProcess([in, out] BOOLEAN *pRet);
Private Function CVstProxy_StopProcess( _
                 ByVal pObj As PTR, _
                 ByRef bRet As Boolean) As Long
    CVstProxy_StopProcess = SendRequestGeneric(WM_STOPPROCESS, 0, ByVal NULL_PTR)
    bRet = g_tSharedData(0).lArg1
End Function

'    long ThreadId([in, out] long* pValue);
Private Function CVstProxy_ThreadId( _
                 ByVal pObj As PTR, _
                 ByRef lRet As Long) As Long
    CVstProxy_ThreadId = E_UNEXPECTED
End Function
