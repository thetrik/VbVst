
[
      odl,
      uuid(A41826AA-AB40-431F-A975-175AFB869A1E)
]
interface IVBVstEffect_dbg : IUnknown {

	long VstVersion([in, out] long* pValue);
	long NumOfParam([in, out] long* pValue);
	long NumOfInputs([in, out] long* pValue);
	long NumOfOutputs([in, out] long* pValue);
	long ParameterProperties([in, out] PTR* pProperties, [in, out] BOOLEAN* pValue);
	long Version([in, out] long* pValue);
	long UniqueId([in, out] long* pValue);
	long ParamName([in] long lIndex, [in, out] BSTR* pOut);
	long ParamLabel([in] long lIndex, [in, out] BSTR* pOut);
	long ParamDisplay([in] long lIndex, [in, out] BSTR* pOut);
	long ParamValue_put([in] long lIndex, [in] float fValue);
	long ParamValue_get([in] long lIndex, [in, out] float* pOut);
	long Suspend();
	long Resume();
	long SampleRate_put([in] float lValue);
	long SampleRate_get([in, out] float* pOut);
	long Process([in] PTR pInput, [in] PTR pOutput, [in] long sampleFrames);
	long ProcessReplacing([in] PTR pInput, [in] PTR pOutput, [in] long sampleFrames);
	long EditorRect([in, out] VBVST2X.ERect* tRet);
	long EditorOpen([in] HANDLE hWnd, [in, out] BOOLEAN* bResult);
	long EditorClose();
	long EditorIdle([in, out] SAFEARRAY(tAutomationRecord) *tRecords, [in, out] long* lNumOfRecords);
	long HasEditor([in, out] BOOLEAN* pValue);
	long ProgramsAreChunks([in, out] BOOLEAN *pValue);
	long SupportsVSTEvents([in, out] BOOLEAN *pValue);
	long CanDo([in] BSTR *ppszRequest, [in, out] BOOLEAN* pValue);
	long AudioMasterCallback_put([in] PTR pfn);
	long AudioMasterCallback_get([in, out] PTR *pfn);
	long AEffectPtr_put([in] PTR pValue);
	long AEffectPtr_get([in, out] PTR* pValue);
	long CanParameterBeAutomated([in] long lIndex, [in, out] BOOLEAN* pValue);
	long TailSize([in, out] long* pValue);
	long NumOfPrograms([in, out] long* pValue);
	long ProgramName_get([in, out] BSTR* pValue);
	long ProgramName_put([in] BSTR pValue);
	long Program_get([in, out] long* pValue);
	long Program_put([in] long lValue);
	long CanMono([in, out] BOOLEAN* pValue);
	long BlockSize([in] long lValue);
	long ProgramNameIndexed([in] long lCategory, [in] long lIndex, [in, out] BSTR* pValue);
	long CopyProgram([in] long lDestination);
	long EffectName([in, out] BSTR* pValue);
	long VendorString([in, out] BSTR* pValue);
	long ProductString([in, out] BSTR* pValue);
	long GetStateChunk([in]BOOLEAN bIsPreset, [in, out] PTR *pData, [in, out] long *pSize);
	long SetStateChunk([in]BOOLEAN bIsPreset, [in] PTR pData, [in] long lSize, [in, out] BOOLEAN *bRet);
	long VendorVersion([in, out] long* pValue);
	long VendorSpecific([in] long lArg1, [in] long lArg2, [in] PTR lpPtr, [in] float fArg3);
	long ProcessEvents([in] VstEvents* pEvents, [in, out] BOOLEAN *pRet);
	long PlugCategory([in, out] long* pValue);
	long SetBypass([in] BOOLEAN bValue, [in, out] BOOLEAN *pSupports);
	long StartProcess([in, out] BOOLEAN *pRet);
	long StopProcess([in, out] BOOLEAN *pRet);
	long ThreadId([in, out] long* pValue);

}
