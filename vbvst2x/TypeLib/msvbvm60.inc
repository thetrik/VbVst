	[dllname("MSVBVM60.DLL")]
	module msvbvm60 {

		[entry("VarPtr")]
		PTR ArrPtr(
			[in] SAFEARRAY(void)* pSrc);

		[entry("GetMem1")]
		void GetMem1(
			[in] void *pSrc,
			[in, out] void *pDst);

		[entry("GetMem2")]
		void GetMem2(
			[in] void *pSrc,
			[in, out] void *pDst);

		[entry("GetMem4")]
		void GetMem4(
			[in] void *pSrc,
			[in, out] void *pDst);

		[entry("GetMem4")]
		void GetMemPtr(
			[in] void *pSrc,
			[in, out] void *pDst);

		[entry("GetMem8")]
		void GetMem8(
			[in] void *pSrc,
			[in, out] void *pDst);

		[entry("PutMem1")]
		void PutMem1(
			[in, out] void *pDst,
			[in] unsigned char v);

		[entry("PutMem2")]
		void PutMem2(
			[in, out] void *pDst,
			[in] short v);

		[entry("PutMem4")]
		void PutMem4(
			[in, out] void *pDst,
			[in] long v);

		[entry("PutMem4")]
		void PutMemPtr(
			[in, out] void *pDst,
			[in] PTR v);

		[entry("PutMem8")]
		void PutMem8(
			[in, out] void *pDst,
			[in] long vL,
			[in] long vH);

		[entry("__vbaAryMove")]
		void vbaAryMove(
			[in] void* pDst,
			[in] void* Source);

		[entry("__vbaObjSetAddref")]
		void vbaObjSetAddref(
			[in, out] void *pDstObj,
			[in] void *pSrcObj);

		[entry("__vbaObjSet")]
		void vbaObjSet(
			[in, out] void *pDstObj,
			[in] void *pSrcObj);

		[entry("__vbaObjAddref")]
		void vbaObjAddref(
			[in] void *pDstObj);

		[entry("__vbaCopyBytesZero")]
		void vbaCopyBytesZero(
			[in] long lSize,
			[in] void *pDst,
			[in] void *pSrc);

		[entry("__vbaStrToAnsi")]
		void vbaStrToAnsi(
			[in] BSTR* ppDest,
			[in] BSTR pSrc);

	}