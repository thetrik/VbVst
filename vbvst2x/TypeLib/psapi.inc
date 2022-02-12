
[dllname("PSAPI.DLL")]
module psapi {

	[entry("GetModuleFileNameExW")]
	long GetModuleFileNameEx(
	  [in] HANDLE hProcess,
	  [in] PTR hModule,
	  [in, out] LPWSTR lpFilename,
	  [in] long nSize
	);

}