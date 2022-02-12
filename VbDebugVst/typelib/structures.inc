
typedef struct SHARED_DATA {

	long lDebuggerThreadId;
	long lDebuggerProcessId;
	long lVstThreadId;
	long lVstProcessId;
	long lSamplesBufSize;
	long lSamplesCount;
	PTR pSamplesBuf;
	PTR pSamplesBufRemote;
	long lEventsBufSize;
	long lEventsCount;
	PTR pEventsBuf;
	PTR pEventsBufRemote;
	long lAutomationBufSize;
	long lAutomationCount;
	PTR pAutomationBuf;
	PTR pAutomationBufRemote;
	PTR pfnHostCallbackRemote;
	long lDataBufferSize;
	PTR pDataBuffer;
	PTR pDataBufferRemote;
	long lArg1;
	long lArg2;
	long lArg3;
	long lArg4;
	PTR pArg1;
	PTR pArg2;
	PTR pArg3;
	PTR pArg4;
	float fArg1;
	float fArg2;
	float fArg3;
	float fArg4;
	VstTimeInfo tCurTimeInfo;
	PTR hDllRemote;
	short pszProgId[40];
	HANDLE hWnd;
	HANDLE hWndApp;
	HANDLE hWndCallback;
	HANDLE hWndContainer;
	HANDLE hEvent;	// Remote
	HRESULT hr;

} SHARED_DATA;