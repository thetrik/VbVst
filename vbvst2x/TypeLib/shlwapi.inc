	[dllname("SHLWAPI.DLL")]
	module shlwapi {
		
		const int STIF_SUPPORT_HEX = 1;

		[entry("IStream_Reset")]
		long IStream_Reset(
		  [in] PTR pstm
		);

		[entry("StrToIntExW")]
		long StrToIntEx(
		  [in] LPWSTR pszString,
		  [in] long dwFlags,
		  [in, out] int *piRet
		);

		[entry("PathFindFileNameW")]
		PTR PathFindFileName(
		  [in] LPWSTR pszPath
		);

}