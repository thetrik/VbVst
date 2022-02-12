
typedef struct CLIENT_ID {
    PTR UniqueProcess;
    PTR UniqueThread;
} CLIENT_ID;

typedef struct THREAD_BASIC_INFORMATION {
	long ExitStatus;
	PTR TebBaseAddress;
	CLIENT_ID ClientId;
	PTR AffinityMask;
	long Priority;
	long BasePriority;
} THREAD_BASIC_INFORMATION;

[dllname("NTDLL.dll")] 
module ntdll {
	
	const int ThreadBasicInformation = 0;

	[entry("NtQueryInformationThread")]
	long NtQueryInformationThread(
		[in] HANDLE ThreadHandle,
		[in] long ThreadInformationClass,
		[in] void* ThreadInformation,
		[in] long ThreadInformationLength,
		[in, out] long* ReturnLength
	);

	[entry("swprintf_s")]
	long __cdecl swprintf_s(
		[in] LPWSTR sBuffer,
		[in] long lCount,
		[in] LPWSTR sFormat,
		[in, optional, defaultvalue(0)] void* pArg1,
		[in, optional, defaultvalue(0)] void* pArg2,
		[in, optional, defaultvalue(0)] void* pArg3,
		[in, optional, defaultvalue(0)] void* pArg4,
		[in, optional, defaultvalue(0)] void* pArg5,
		[in, optional, defaultvalue(0)] void* pArg6,
		[in, optional, defaultvalue(0)] void* pArg7,
		[in, optional, defaultvalue(0)] void* pArg8
	);

	[entry("swscanf_s")]
	long __cdecl swscanf_s(
		[in] LPWSTR sBuffer,
		[in] LPWSTR sFormat,
		[in, optional, defaultvalue(0)] void* pArg1,
		[in, optional, defaultvalue(0)] void* pArg2,
		[in, optional, defaultvalue(0)] void* pArg3,
		[in, optional, defaultvalue(0)] void* pArg4,
		[in, optional, defaultvalue(0)] void* pArg5,
		[in, optional, defaultvalue(0)] void* pArg6,
		[in, optional, defaultvalue(0)] void* pArg7,
		[in, optional, defaultvalue(0)] void* pArg8
	);

}