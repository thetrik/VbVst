

[dllname("ADVAPI32.DLL")]
module advapi32 {
	
	const int KEY_ENUMERATE_SUB_KEYS  = 8;
	const int KEY_QUERY_VALUE = 1;
	const int KEY_READ = 0x20019;
	const int HKEY_CLASSES_ROOT = 0x80000000;

	[entry("RegOpenKeyExW")]
	long RegOpenKeyEx(
	  [in] HANDLE hKey,
	  [in] LPWSTR lpSubKey,
	  [in] long ulOptions,
	  [in] long samDesired,
	  [in, out] HANDLE* phkResult
	);

	[entry("RegCloseKey")]
	long RegCloseKey(
	  [in] HANDLE hKey
	);

	[entry("RegEnumKeyExW")]
	long RegEnumKeyEx(
	  [in] HANDLE hKey,
	  [in] long dwIndex,
	  [in, out] LPWSTR lpName,
	  [in, out] long* lpcchName,
	  [in] PTR lpReserved,
	  [in, out] LPWSTR lpClass,
	  [in, out] void* lpcchClass,
	  [in, out] void* lpftLastWriteTime
	);

	[entry("RegQueryInfoKeyW")]
	long RegQueryInfoKey(
	  [in] HANDLE hKey,
	  [in, out] LPWSTR lpClass,
	  [in, out] void* lpcchClass,
	  [in] PTR lpReserved,
	  [in, out] void* lpcSubKeys,
	  [in, out] void* lpcbMaxSubKeyLen,
	  [in, out] void* lpcbMaxClassLen,
	  [in, out] void* lpcValues,
	  [in, out] void* lpcbMaxValueNameLen,
	  [in, out] void* lpcbMaxValueLen,
	  [in, out] void* lpcbSecurityDescriptor,
	  [in, out] void* lpftLastWriteTime
	);

	[entry("RegQueryValueExW")]
	long RegQueryValueEx(
	  [in] HANDLE hKey,
	  [in] LPWSTR lpValueName,
	  [in] PTR lpReserved,
	  [in, out] void* lpType,
	  [in, out] void* lpData,
	  [in, out] void* lpcbData
	);

	[entry("RegEnumValueW")]
	long RegEnumValue(
	  [in] HANDLE hKey,
	  [in] long dwIndex,
	  [in, out] LPWSTR lpValueName,
	  [in, out] long* lpcchValueName,
	  [in] PTR lpReserved,
	  [in, out] void* lpType,
	  [in, out] void* lpData,
	  [in, out] void* lpcbData
	);


}