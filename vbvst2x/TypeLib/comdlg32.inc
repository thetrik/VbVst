
typedef struct OPENFILENAME {
	long lStructSize;
	HANDLE hwndOwner;
	PTR hInstance;
	PTR lpstrFilter;
	PTR lpstrCustomFilter;
	long nMaxCustFilter;
	long nFilterIndex;
	PTR lpstrFile;
	long nMaxFile;
	PTR lpstrFileTitle;
	long nMaxFileTitle;
	PTR lpstrInitialDir;
	PTR lpstrTitle;
	long Flags;
	short nFileOffset;
	short nFileExtension;
	PTR lpstrDefExt;
	PTR lCustData;
	PTR lpfnHook;
	PTR lpTemplateName;
} OPENFILENAME;

[dllname("COMDLG32.DLL")]
module comdlg32 {

	const int OFN_READONLY = 0x00000001;
	const int OFN_OVERWRITEPROMPT = 0x00000002;
	const int OFN_HIDEREADONLY = 0x00000004;
	const int OFN_NOCHANGEDIR = 0x00000008;
	const int OFN_SHOWHELP = 0x00000010;
	const int OFN_ENABLEHOOK = 0x00000020;
	const int OFN_ENABLETEMPLATE = 0x00000040;
	const int OFN_ENABLETEMPLATEHANDLE = 0x00000080;
	const int OFN_NOVALIDATE = 0x00000100;
	const int OFN_ALLOWMULTISELECT = 0x00000200;
	const int OFN_EXTENSIONDIFFERENT = 0x00000400;
	const int OFN_PATHMUSTEXIST = 0x00000800;
	const int OFN_FILEMUSTEXIST = 0x00001000;
	const int OFN_CREATEPROMPT = 0x00002000;
	const int OFN_SHAREAWARE = 0x00004000;
	const int OFN_NOREADONLYRETURN = 0x00008000;
	const int OFN_NOTESTFILECREATE = 0x00010000;
	const int OFN_NONETWORKBUTTON = 0x00020000;
	const int OFN_NOLONGNAMES = 0x00040000;
	const int OFN_EXPLORER = 0x00080000;
	const int OFN_NODEREFERENCELINKS = 0x00100000;
	const int OFN_LONGNAMES = 0x00200000;
	const int OFN_ENABLEINCLUDENOTIFY = 0x00400000;
	const int OFN_ENABLESIZING = 0x00800000;
	const int OFN_DONTADDTORECENT = 0x02000000;
	const int OFN_FORCESHOWHIDDEN = 0x10000000;

	[entry("GetOpenFileNameW")]
	long GetOpenFileName(
		[in, out] OPENFILENAME* lpOFN
	);

	[entry("GetSaveFileNameW")]
	long GetSaveFileName(
		[in] OPENFILENAME* pOpenfilename
	);
};