
typedef struct VstEvents VstEvents;

[
      odl,
      uuid(A41826AA-AB40-431F-A975-175AFB869A1E)
]
interface IVBVstEffect : IUnknown {
	
	[propget]
	HRESULT VstVersion([out, retval] long* pValue);

	[propget]
	HRESULT NumOfParam([out, retval] long* pValue);

	[propget]
	HRESULT NumOfInputs([out, retval] long* pValue);

	[propget]
	HRESULT NumOfOutputs([out, retval] long* pValue);

	[propget]
	HRESULT ParameterProperties([in, out] PTR* pProperties, [out, retval] BOOLEAN* pValue);

	[propget]
	HRESULT Version([out, retval] long* pValue);

	[propget]
	HRESULT UniqueId([out, retval] long* pValue);

	[propget]
	HRESULT ParamName([in] long lIndex, [out, retval] BSTR* pOut);

	[propget]
	HRESULT ParamLabel([in] long lIndex, [out, retval] BSTR* pOut);

	[propget]
	HRESULT ParamDisplay([in] long lIndex, [out, retval] BSTR* pOut);

	[propput]
	HRESULT ParamValue([in] long lIndex, [in] float fValue);

	[propget]
	HRESULT ParamValue([in] long lIndex, [out, retval] float* pOut);

	HRESULT Suspend();
	HRESULT Resume();

	[propput]
	HRESULT SampleRate([in] float lValue);

	[propget]
	HRESULT SampleRate([out, retval] float* pOut);

	HRESULT Process([in] PTR pInput, [in] PTR pOutput, [in] long sampleFrames);
	HRESULT ProcessReplacing([in] PTR pInput, [in] PTR pOutput, [in] long sampleFrames);

	[propget]
	HRESULT EditorRect([out, retval] ERect* tRet);

	HRESULT EditorOpen([in] HANDLE hWnd, [out, retval] BOOLEAN* bResult);
	HRESULT EditorClose();
	HRESULT EditorIdle([in, out] SAFEARRAY(tAutomationRecord) *tRecords, [out, retval] long* lNumOfRecords);

	[propget]
	HRESULT HasEditor([out, retval] BOOLEAN* pValue);
	[propget]
	HRESULT ProgramsAreChunks([out, retval] BOOLEAN *pValue);
	[propget]
	HRESULT SupportsVSTEvents([out, retval] BOOLEAN *pValue);

	[propget]
	HRESULT CanDo([in] BSTR *ppszRequest, [out, retval] BOOLEAN* pValue);

	[propput]
	HRESULT AudioMasterCallback([in] PTR pfn);
	[propget]
	HRESULT AudioMasterCallback([out, retval] PTR *pfn);

	[propput]
	HRESULT AEffectPtr([in] PTR pValue);
	[propget]
	HRESULT AEffectPtr([out, retval] PTR* pValue);

	[propget]
	HRESULT CanParameterBeAutomated([in] long lIndex, [out, retval] BOOLEAN* pValue);
	[propget]
	HRESULT TailSize([out, retval] long* pValue);

	[propget]
	HRESULT NumOfPrograms([out, retval] long* pValue);
	[propget]
	HRESULT ProgramName([out, retval] BSTR* pValue);
	[propput]
	HRESULT ProgramName([in] BSTR pValue);

	[propget]
	HRESULT Program([out, retval] long* pValue);
	[propput]
	HRESULT Program([in] long lValue);

	[propget]
	HRESULT CanMono([out, retval] BOOLEAN* pValue);

	[propput]
	HRESULT BlockSize([in] long lValue);

	[propget]
	HRESULT ProgramNameIndexed([in] long lCategory, [in] long lIndex, [out, retval] BSTR* pValue);
	HRESULT CopyProgram([in] long lDestination);

	[propget]
	HRESULT EffectName([out, retval] BSTR* pValue);
	[propget]
	HRESULT VendorString([out, retval] BSTR* pValue);
	[propget]
	HRESULT ProductString([out, retval] BSTR* pValue);

	HRESULT GetStateChunk([in]BOOLEAN bIsPreset, [in, out] PTR *pData, [out, retval] long *pSize);
	HRESULT SetStateChunk([in]BOOLEAN bIsPreset, [in] PTR pData, [in] long lSize, [out, retval] BOOLEAN *bRet);

	[propget]
	HRESULT VendorVersion([out, retval] long* pValue);
	HRESULT VendorSpecific([in] long lArg1, [in] long lArg2, [in] PTR lpPtr, [in] float fArg3);

	HRESULT ProcessEvents([in] VstEvents* pEvents, [out, retval] BOOLEAN *pRet);

	[propget]
	HRESULT PlugCategory([out, retval] long* pValue);

	HRESULT SetBypass([in] BOOLEAN bValue, [out, retval] BOOLEAN *pSupports);
	HRESULT StartProcess([out, retval] BOOLEAN *pRet);
	HRESULT StopProcess([out, retval] BOOLEAN *pRet);

	[propget, hidden]
	HRESULT ThreadId([out, retval] long* pValue);

}
