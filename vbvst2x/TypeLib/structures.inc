
typedef struct tAutomationRecord {
	long lParamIndex;
	float fParamValue;
} tAutomationRecord;

typedef struct ERect ERect;
typedef struct AEEffect AEEffect;

typedef struct CVBVstDispatcher {

	CRITICAL_SECTION tLock;
	long lTlsIndex;

	PTR pPlugObjPtr;
	PTR pAEffect;
	PTR pHostCallback;
	HANDLE hWndContainer;
	BOOLEAN bHasEditor;
	BOOLEAN bProgAreChunks;
	BOOLEAN bSupportsEvents;
	long lVstVersion;
	long lNumOfParams;
	long lNumOfInputs;
	long lNumOfOutputs;
	PTR pParamProp;
	long lVersion;
	long lUniqueId;
	ERect tRect;
	long lNumOfPrograms;
	BOOLEAN bCanMono;
	long lPlugCategory;
} CVBVstDispatcher;

