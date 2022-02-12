
typedef struct CRITICAL_SECTION {
	PTR DebugInfo;

	//
	//  The following three fields control entering and exiting the critical
	//  section for the resource
	//

	long LockCount;
	long RecursionCount;
	HANDLE OwningThread;        // from the thread's ClientId->UniqueThread
	HANDLE LockSemaphore;
	long SpinCount;        // force size on 64-bit systems when packed
} CRITICAL_SECTION;

typedef struct ACTCTX {
  long   cbSize;
  long   dwFlags;
  PTR  lpSource;
  short  wProcessorArchitecture;
  short  wLangId;
  PTR  lpAssemblyDirectory;
  PTR  lpResourceName;
  PTR  lpApplicationName;
  PTR hModule;
} ACTCTX;

typedef struct PROCESSENTRY32 {
  long dwSize;
  long cntUsage;
  long th32ProcessID;
  PTR th32DefaultHeapID;
  long th32ModuleID;
  long cntThreads;
  long th32ParentProcessID;
  long pcPriClassBase;
  long dwFlags;
  short szExeFile[260];
} PROCESSENTRY32;

typedef struct THREADENTRY32 {
  long dwSize;
  long cntUsage;
  long th32ThreadID;
  long th32OwnerProcessID;
  long tpBasePri;
  long tpDeltaPri;
  long dwFlags;
} THREADENTRY32;

typedef struct LARGE_INTEGER {
	long LowPart;
	long HighPart;
} LARGE_INTEGER;

[dllname("KERNEL32.DLL")]
	module kernel32 {
		
		const int NULL_PTR = 0;
		const int TLS_OUT_OF_INDEXES = 0xFFFFFFFF;
		const int GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT  = 2;
		const int GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS = 4;
		const int CP_ACP = 0;
		const int ACTCTX_FLAG_HMODULE_VALID = 128;
		const int ACTCTX_FLAG_RESOURCE_NAME_VALID = 8;
		const int INVALID_HANDLE_VALUE = -1;
		const int PAGE_EXECUTE_READWRITE=0x40;
		const int WAIT_OBJECT_0 = 0x00000000;
		const int GMEM_FIXED = 0;
		const int PAGE_READWRITE = 4;
		const int FILE_MAP_WRITE = 2;
		const int DLL_PROCESS_ATTACH = 1;
		const int DLL_PROCESS_DETACH = 0;
		const int PROCESS_CREATE_THREAD = 2;
		const int PROCESS_VM_READ = 0x10;
		const int PROCESS_VM_WRITE = 0x0020;
		const int PROCESS_VM_OPERATION = 0x08;
		const int PROCESS_QUERY_INFORMATION = 0x0400;
		const int PROCESS_DUP_HANDLE = 0x0040;
		const int THREAD_QUERY_INFORMATION = 0x40;
		const int TH32CS_SNAPPROCESS = 2;
		const int TH32CS_SNAPTHREAD = 4;
		const int MEM_COMMIT = 0x00001000;
		const int MEM_RESERVE = 0x00002000;
		const int MEM_RELEASE = 0x00008000;
		const int SYNCHRONIZE = 0x00100000;
		const int MAX_PATH  = 260;
		const int STILL_ACTIVE = 259;
		const int DUPLICATE_SAME_ACCESS = 2;
		const int GENERIC_READ = 0x80000000;
		const int GENERIC_WRITE = 0x40000000;
		const int CREATE_ALWAYS = 2;
		const int CREATE_NEW = 1;
		const int OPEN_ALWAYS = 4;
		const int OPEN_EXISTING = 3;
		const int TRUNCATE_EXISTING = 5;
		const int FILE_CURRENT = 1;
		const int FILE_END = 2;
		const int FILE_BEGIN = 0;
		const int FILE_ATTRIBUTE_NORMAL = 0x00000080;
		const int DUPLICATE_CLOSE_SOURCE = 1;

		[entry("CreateFileW")]
		HANDLE CreateFile(
			[in] LPWSTR lpFileName,
			[in] long dwDesiredAccess,
			[in] long dwShareMode,
			[in] void *lpSecurityAttributes,
			[in] long dwCreationDisposition,
			[in] long dwFlagsAndAttributes,
			[in] HANDLE hTemplateFile
		);

		[entry("ReadFile")]
		long ReadFile(
			[in] HANDLE hFile,
			[in] void *lpBuffer,
			[in] long nNumberOfBytesToRead,
			[in] long *lpNumberOfBytesRead,
			[in] void *lpOverlapped
		);

		[entry("WriteFile")]
		long WriteFile(
			[in] HANDLE hFile,
			[in] void *lpBuffer,
			[in] long nNumberOfBytesToWrite,
			[in] long *lpNumberOfBytesWritten,
			[in] void *lpOverlapped
		);

		[entry("SetFilePointerEx")]
		long SetFilePointerEx(
			[in] HANDLE hFile,
			[in] long liDistanceToMoveL,
			[in] long liDistanceToMoveH,
			[in] void* lpNewFilePointer,
			[in] long dwMoveMethod
		);

		[entry("ReadProcessMemory")]
		long ReadProcessMemory(
		  [in] HANDLE hProcess,
		  [in] PTR lpBaseAddress,
		  [in, out] void* lpBuffer,
		  [in] long nSize,
		  [in, out] long *lpNumberOfBytesRead
		);

		[entry("DuplicateHandle")]
		long DuplicateHandle(
		  [in]  HANDLE   hSourceProcessHandle,
		  [in]  HANDLE   hSourceHandle,
		  [in]  HANDLE   hTargetProcessHandle,
		  [in, out] HANDLE* lpTargetHandle,
		  [in]  long    dwDesiredAccess,
		  [in]  long     bInheritHandle,
		  [in]  long    dwOptions
		);

		[entry("GetExitCodeThread")]
		long GetExitCodeThread(
			[in] HANDLE hThread,
			[in] long* lpExitCode
		);

		[entry("GetCurrentProcessId")]
		long GetCurrentProcessId();

		[entry("GetModuleFileNameW")]
		long GetModuleFileName(
			[in] PTR hModule,
			[in] LPWSTR lpFilename,
			[in] long nSize
		);

		[entry("GlobalAlloc")]
		PTR GlobalAlloc(
			[in] long wFlags,
			[in] long dwBytes
		);

		[entry("GlobalFree")]
		long GlobalFree(
			[in] PTR hMem
		);

		[entry("FreeLibrary")]
		long FreeLibrary(
			[in] PTR hLib
		);

		[entry("MultiByteToWideChar")]
		long MultiByteToWideChar(
			[in] long CodePage,
			[in] long dwFlags,
			[in] void* lpMultiByteStr,
			[in] long cchMultiByte,
			[in] void* lpWideCharStr,
			[in] long cchWideChar
		);

		[entry("CreateThread")]
		HANDLE CreateThread(
			[in] void* lpThreadAttributes,
			[in] SIZE_T dwStackSize,
			[in] PTR lpStartAddress,
			[in] void* lpParameter,
			[in] long dwCreationFlags,
			[in] long* lpThreadId
		);

		[entry("CloseHandle")]
		long CloseHandle(
			[in] HANDLE hObject);

		[entry("WaitForSingleObject")]
		long WaitForSingleObject(
			[in] HANDLE hHandle,
			[in] long dwMilliseconds
		);

		[entry("SetEvent")]
		long SetEvent(
			[in] HANDLE hEvent
		);

		[entry("ResetEvent")]
		long ResetEvent(
			[in] HANDLE hEvent
		);

		[entry("CreateEventW")]
		HANDLE CreateEvent(
			[in] void *lpEventAttributes,
			[in] long bManualReset,
			[in] long bInitialState,
			[in] LPWSTR lpName
		);

		[entry("VirtualProtect")]
		long VirtualProtect(
			[in]  PTR lpAddress,
			[in]  SIZE_T dwSize,
			[in]  long  flNewProtect,
			[in, out] long* lpflOldProtect
		);

		[entry("lstrcmpA")]
		int lstrcmpA(
			[in] void* lpString1, 
			[in] void* lpString2
		);

		[entry("GetProcAddress")]
		long GetProcAddress(
			[in] PTR hLib, 
			[in] LPSTR pszFn);

		[entry("GetLastError")]
		long GetLastError();

		[entry("ActivateActCtx")]
		long ActivateActCtx(
		  [in]  HANDLE hActCtx,
		  [in, out] long *lpCookie
		);

		[entry("DeactivateActCtx")]
		long DeactivateActCtx(
		  [in] long dwFlags,
		  [in] long ulCookie
		);

		[entry("CreateActCtxW")]
		HANDLE CreateActCtx(
			[in, out] ACTCTX* pActCtx
		);

		[entry("ReleaseActCtx")]
		void ReleaseActCtx(
		  [in] HANDLE hActCtx
		);

		[entry("lstrcpynA")]
		int lstrcpynA(
			[in] void* lpString1, 
			[in] void* lpString2,
			[in] long iMaxLength
		);

		[entry("lstrcpynW")]
		int lstrcpyn(
			[in] void* lpString1, 
			[in] void* lpString2,
			[in] long iMaxLength
		);

		[entry("lstrcpyA")]
		int lstrcpyA(
			[in] void* lpString1, 
			[in] void* lpString2
		);

		[entry("Sleep")]
		long Sleep(
			[in] long interval
		);

		[entry("InterlockedDecrement")]
		long InterlockedDecrement(
			[in, out] void *Addend
		);

		[entry("InterlockedIncrement")]
		long InterlockedIncrement(
			[in, out] void *Addend
		);

		[entry("InterlockedCompareExchange")]
		long InterlockedCompareExchange(
			[in, out] void *Destination,
			[in] long ExChange,
			[in] long Comperand
		);

		[entry("InterlockedExchange")]
		long InterlockedExchange(
			[in, out] void *Target,
			[in] long Value
		);

		[entry("RtlMoveMemory")]
		void memcpy(
			[in] void* Destination,
			[in] void* Source,
			[in] long Length
		);

		[entry("RtlFillMemory")]
		void memset(
			[in] void* Destination,
			[in] long Length,
			[in] long Fill
		);

		[entry("GetCurrentThreadId")]
		long GetCurrentThreadId();

		[entry("GetModuleHandleExW")]
		long GetModuleHandleEx(
			[in] long dwFlags,
			[in] void* lpModuleName, 
			[in] void* phModule);

		[entry("GetModuleHandleW")]
		PTR GetModuleHandle(
			[in] LPWSTR pszLib);

		[entry("LoadLibraryW")]
		PTR LoadLibrary(
			[in] LPWSTR pszLib);

		[entry("TlsAlloc")]
		long TlsAlloc();

		[entry("TlsFree")]
		long TlsFree(
			[in] long dwTlsIndex);

		[entry("TlsSetValue")]
		long TlsSetValue(
			[in] long dwTlsIndex,
			[in] void *lpTlsValue);

		[entry("TlsGetValue")]
		long TlsGetValue(
			[in] long dwTlsIndex);

		[entry("WideCharToMultiByte")]
		long WideCharToMultiByte(
			[in] long CodePage,
			[in] long dwFlags,
			[in] void* lpWideCharStr,
			[in] long cchWideChar,
			[in] void* lpMultiByteStr,
			[in] long cchMultiByte,
			[in] void* lpDefaultChar,
			[in] long* lpUsedDefaultChar
		);

		[entry("GetUserDefaultLCID")]
		long GetUserDefaultLCID();

		[entry("InitializeCriticalSection")]
		long InitializeCriticalSection(
			[in] CRITICAL_SECTION* lpCriticalSection
		);

		[entry("DeleteCriticalSection")]
		long DeleteCriticalSection(
			[in] CRITICAL_SECTION* lpCriticalSection
		);

		[entry("EnterCriticalSection")]
		long EnterCriticalSection(
			[in] CRITICAL_SECTION* lpCriticalSection
		);

		[entry("LeaveCriticalSection")]
		long LeaveCriticalSection(
			[in] CRITICAL_SECTION* lpCriticalSection
		);

		[entry("CreateFileMappingW")]
		HANDLE CreateFileMapping(
			[in] HANDLE hFile,
			[in] void* lpFileMappingAttributes,
			[in] long flProtect,
			[in] long dwMaximumSizeHigh,
			[in] long dwMaximumSizeLow,
			[in] LPWSTR lpName
		);

		[entry("MapViewOfFile")]
		PTR MapViewOfFile(
			[in] HANDLE hFileMappingObject,
			[in] long dwDesiredAccess,
			[in] long dwFileOffsetHigh,
			[in] long dwFileOffsetLow,
			[in] long dwNumberOfBytesToMap
		);

		[entry("UnmapViewOfFile")]
		long UnmapViewOfFile(
			[in] PTR lpBaseAddress
		);

		[entry("CreateRemoteThread")]
		HANDLE CreateRemoteThread(
			[in] HANDLE hProcess,
			[in] void* lpThreadAttributes,
			[in] long dwStackSize,
			[in] PTR lpStartAddress,
			[in] void* lpParameter,
			[in] long dwCreationFlags,
			[in, out] long* lpThreadId
		);

		[entry("OpenProcess")]
		HANDLE OpenProcess(
			[in] long dwDesiredAccess,
			[in] long bInheritHandle,
			[in] long dwProcessId
		);

		[entry("OpenThread")]
		HANDLE OpenThread(
			[in] long dwDesiredAccess,
			[in] long bInheritHandle,
			[in] long dwThreadId
		);

		[entry("WriteProcessMemory")]
		BOOL WriteProcessMemory(
			[in] HANDLE hProcess,
			[in] PTR lpBaseAddress,
			[in] void* lpBuffer,
			[in] long nSize,
			[in] long* lpNumberOfBytesWritten
		);

		[entry("VirtualAllocEx")]
		PTR VirtualAllocEx(
			[in] HANDLE hProcess,
			[in] PTR lpBaseAddress,
			[in] long dwSize,
			[in] long flAllocationType,
			[in] long flProtect
		);

		[entry("VirtualFreeEx")]
		long VirtualFreeEx(
			[in] HANDLE hProcess,
			[in] PTR lpBaseAddress,
			[in] long dwSize,
			[in] long dwFreeType
		);

		[entry("CreateToolhelp32Snapshot")]
		HANDLE CreateToolhelp32Snapshot(
			[in] long dwFlags,
			[in] long th32ProcessID
		);

		[entry("Process32FirstW")]
		long Process32First(
		  [in] HANDLE hSnapshot,
		  [in, out] PROCESSENTRY32* lppe
		);

		[entry("Process32NextW")]
		long Process32Next(
		  [in] HANDLE hSnapshot,
		  [in, out] PROCESSENTRY32* lppe
		);

		[entry("Thread32First")]
		long Thread32First(
		  [in] HANDLE hSnapshot,
		  [in, out] THREADENTRY32* lppe
		);

		[entry("Thread32Next")]
		long Thread32Next(
		  [in] HANDLE hSnapshot,
		  [in, out] THREADENTRY32* lppe
		);

		[entry("lstrcatW")]
		PTR lstrcat(
		  [in] LPWSTR lpString1,
		  [in] LPWSTR lpString2
		);

		[entry("lstrcmpiW")]
		long lstrcmpi(
		  [in] void* lpString1,
		  [in] void* lpString2
		);

		[entry("FreeLibraryAndExitThread")]
		void FreeLibraryAndExitThread(
		  [in] PTR hLibModule,
		  [in] long dwExitCode
		);

		[entry("WritePrivateProfileStringW")]
		long WritePrivateProfileString(
			[in] LPWSTR lpAppName,
			[in] LPWSTR lpKeyName,
			[in] LPWSTR lpString,
			[in] LPWSTR lpFileName
		);

		[entry("GetPrivateProfileStringW")]
		long GetPrivateProfileString(
			[in] LPWSTR lpAppName,
			[in] LPWSTR lpKeyName,
			[in] LPWSTR lpDefault,
			[in] LPWSTR lpReturnedString,
			[in] long nSize,
			[in] LPWSTR lpFileName
		);

	}