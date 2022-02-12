
typedef enum REGKIND {
  REGKIND_DEFAULT,
  REGKIND_REGISTER,
  REGKIND_NONE
} REGKIND;

typedef struct NUMPARSE {
  long cDig;
  long dwInFlags;
  long dwOutFlags;
  long cchUsed;
  long nBaseShift;
  long nPwr10;
} NUMPARSE;

	[dllname("OLEAUT32.DLL")]
	module oleaut32 {
		
		const int CC_CDECL = 1;
		const int CC_STDCALL = 4;

		const int NUMPRS_LEADING_WHITE  = 0x00001;
		const int NUMPRS_TRAILING_WHITE = 0x00002;
		const int NUMPRS_LEADING_PLUS   = 0x00004;
		const int NUMPRS_TRAILING_PLUS  = 0x00008;
		const int NUMPRS_LEADING_MINUS  = 0x00010;
		const int NUMPRS_TRAILING_MINUS = 0x00020;
		const int NUMPRS_HEX_OCT        = 0x00040;
		const int NUMPRS_PARENS         = 0x00080;
		const int NUMPRS_DECIMAL        = 0x00100;
		const int NUMPRS_THOUSANDS      = 0x00200;
		const int NUMPRS_CURRENCY       = 0x00400;
		const int NUMPRS_EXPONENT       = 0x00800;
		const int NUMPRS_USE_ALL        = 0x01000;
		const int NUMPRS_STD            = 0x01FFF;
		const int NUMPRS_NEG            = 0x10000;
		const int NUMPRS_INEXACT        = 0x20000;

		const int VTBIT_I1 = 1 << 16;
		const int VTBIT_UI1 = 1 << 17;
		const int VTBIT_I2 = 1 << 2;
		const int VTBIT_UI2 = 1 << 18;
		const int VTBIT_I4 = 1 << 3;
		const int VTBIT_UI4 = 1 << 19;
		const int VTBIT_I8 = 1 << 20;
		const int VTBIT_UI8 = 1 << 21;
		const int VTBIT_R4 = 1 << 4;
		const int VTBIT_R8 = 1 << 5;
		const int VTBIT_CY = 1 << 6;
		const int VTBIT_DECIMAL = 1 << 14;

		[entry("VarNumFromParseNum")]
		long VarNumFromParseNum(
		  [in] NUMPARSE *pnumprs,
		  [in] void *rgbDig,
		  [in] long dwVtBits,
		  [in, out] VARIANT *pvar
		);

		[entry("VarParseNumFromStr")]
		long VarParseNumFromStr(
		  [in] LPWSTR strIn,
		  [in] long lcid,
		  [in] long dwFlags,
		  [in, out] NUMPARSE *pnumprs,
		  [in, out] void *rgbDig
		);

		[entry("DispCallFunc")]
		long DispCallFunc(
			[in] void *pvInstance,
			[in] PTR oVft,
			[in] long cc,
			[in] long vtReturn,
			[in] long cActuals,
			[in] void *prgvt,
			[in] void *prgpvarg,
			[in] void *pvargResult
		);

		[entry("SysAllocStringByteLen")]
		BSTR SysAllocStringByteLen(
		  [in] void *psz,
		  [in] long len
		);

		[entry("SysAllocString")]
		BSTR SysAllocString(
		  [in] void *psz
		);

		[entry("LoadTypeLibEx")]
		long LoadTypeLibEx(
		   [in] LPWSTR szFile,
		   [in] REGKIND regkind,
		   [in] IUnknown **pptlib
		);

	}