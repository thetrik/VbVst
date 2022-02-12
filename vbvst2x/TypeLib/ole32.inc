
typedef struct UUID {
	long Data1;
	short Data2;
	short Data3;
	unsigned char Data4[8];
} UUID;

typedef struct SAFEARRAYBOUND {
  long cElements;
  long lLbound;
} SAFEARRAYBOUND;

typedef struct SAFEARRAY1D {
  short cDims;
  short fFeatures;
  long cbElements;
  long cLocks;
  PTR pvData;
  SAFEARRAYBOUND rgsabound[1];
} SAFEARRAY1D;

	[dllname("OLE32.DLL")]
	module ole32 {

		const int S_OK = 0;
		const int E_NOINTERFACE = 0x80004002;
		const int E_FAIL = 0x80004005;
		const int E_UNEXPECTED = 0x8000FFFF;
		const int E_OUTOFMEMORY = 0x8007000E;
		const int CLSCTX_INPROC_SERVER	= 0x1;
		const int FADF_AUTO = 1;
		const int MSHCTX_INPROC = 3;
		const int MSHLFLAGS_TABLESTRONG = 1;
		const int REGDB_E_KEYMISSING = 0x80040152;
		const int RPC_E_DISCONNECTED = 0x80010108;
		const int E_PENDING = 0x8000000A;
		const int S_FALSE = 1;

		[entry("CreateStreamOnHGlobal")]
		long CreateStreamOnHGlobal(
		  [in] HANDLE hGlobal,
		  [in] long fDeleteOnRelease,
		  [in, out] void *ppstm
		);

		[entry("CoGetCurrentProcess")]
		long CoGetCurrentProcess();

		[entry("CoCreateInstance")]
		long CoCreateInstance(
			[in] UUID* rclsid,
			[in] void *pUnkOuter,
			[in] long dwClsContext,
			[in] UUID* riid,
			[in, out] void* ppv
		);

    	[entry("CoInitialize")]
		int CoInitialize(
			[in] void* pvReserved); 

		[entry("CoUninitialize")]
			void CoUninitialize();

		[entry("CoFreeUnusedLibraries")]
			void CoFreeUnusedLibraries();

		[entry("CoTaskMemRealloc")]
		PTR CoTaskMemRealloc(
			[in] PTR pv,
			[in] long cb
		);

		[entry("CoTaskMemFree")]
		void CoTaskMemFree(
			[in] PTR pv
		);

		[entry("CoMarshalInterface")]
		long CoMarshalInterface(
			[in] PTR pStm,
			[in] UUID* riid,
			[in] void *pUnk,
			[in] long dwDestContext,
			[in] void *pvDestContext,
			[in] long mshlflags
		);

		[entry("CoMarshalInterThreadInterfaceInStream")]
		long CoMarshalInterThreadInterfaceInStream(
			[in] UUID* riid,
			[in] void *pUnk,
			[in, out] void *ppStm
		);

		[entry("CoGetInterfaceAndReleaseStream")]
		long CoGetInterfaceAndReleaseStream(
			[in] PTR pStm,
			[in] UUID* riid,
			[in, out] void *ppv
		);
		
		[entry("CoUnmarshalInterface")]
		long CoUnmarshalInterface(
			[in] PTR pStm,
			[in] UUID* riid,
			[in, out] void *ppv
		);

		[entry("CoReleaseMarshalData")]
		long CoReleaseMarshalData(
			[in] PTR pStm
		);

		[entry("IsEqualGUID")]
		long IsEqualGUID(
			[in] void *rguid1,
			[in] void *rguid2
		);

		[entry("CLSIDFromProgID")]
		long CLSIDFromProgID(
			[in] void* lpszProgID,
			[in] UUID *lpclsid
		);


	};