
[dllname("MSIMG32.DLL")]
module msimg32 {

	[entry("AlphaBlend")]
	long AlphaBlend(
		[in] HANDLE hdcDest,
		[in] long xDest,
		[in] long yDest,
		[in] long WidthDest,
		[in] long HeightDest,
		[in] HANDLE hdcSrc,
		[in] long xSrc,
		[in] long ySrc,
		[in] long WidthSrc,
		[in] long HeightSrc,
		[in] long Blendfunc
	);


}