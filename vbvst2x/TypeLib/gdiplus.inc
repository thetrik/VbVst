typedef struct GdiplusStartupInput {
	long GdiplusVersion;
	PTR DebugEventCallback;
	long SuppressBackgroundThread;
	long SuppressExternalCodecs;
} GdiplusStartupInput;

[dllname("GDIPLUS.DLL")]
module gdiplus {

	const int SmoothingModeAntiAlias = 4;
	const int UnitPixel = 2;

	const int MatrixOrderPrepend = 0;
	const int MatrixOrderAppend = 1;

	[entry("GdiplusStartup")]
	long GdiplusStartup(
      [in] HANDLE* phToken,
	  [in] GdiplusStartupInput* pInputBuf,
	  [in, optional, defaultvalue(0)] HANDLE outputbuf
	);

	[entry("GdipCreateFromHDC")]
	long GdipCreateFromHDC(
      [in] HANDLE hdc,
	  [in, out] HANDLE *phGraphics
	);

	[entry("GdipDeleteGraphics")]
	long GdipDeleteGraphics(
      [in] HANDLE hGraphics
	);

	[entry("GdiplusShutdown")]
	long GdiplusShutdown(
      [in] HANDLE hToken
	);

	[entry("GdipSetSmoothingMode")]
	long GdipSetSmoothingMode(
      [in] HANDLE hGraphics,
	  [in] long lSmoothingMd
	);

	[entry("GdipCreateSolidFill")]
	long GdipCreateSolidFill(
      [in] long lARGB,
	  [in, out] HANDLE *phBrush
	);

	[entry("GdipDeleteBrush")]
	long GdipDeleteBrush(
      [in] HANDLE hBrush
	);

	[entry("GdipSetSolidFillColor")]
	long GdipSetSolidFillColor(
      [in] HANDLE hBrush,
	  [in] long lARGB
	);

	[entry("GdipCreatePen1")]
	long GdipCreatePen1(
      [in] long lColor,
	  [in] float fWidth,
	  [in] long lUnit,
	  [in, out] HANDLE *phPen
	);

	[entry("GdipDeletePen")]
	long GdipDeletePen(
      [in] HANDLE hPen
	);

	[entry("GdipSetPenWidth")]
	long GdipSetPenWidth(
      [in] HANDLE hPen,
	  [in] float fWidth
	);

	[entry("GdipSetPenColor")]
	long GdipSetPenColor(
      [in] HANDLE hPen,
	  [in] long lARGB
	);

	[entry("GdipDrawLine")]
	long GdipDrawLine(
      [in] HANDLE hGraphics,
	  [in] HANDLE hPen,
	  [in] float fx1,
	  [in] float fy1,
	  [in] float fx2,
	  [in] float fy2
	);

	[entry("GdipFillEllipse")]
	long GdipFillEllipse(
      [in] HANDLE hGraphics,
	  [in] HANDLE hBrush,
	  [in] float fx,
	  [in] float fy,
	  [in] float fWidth,
	  [in] float fHeight
	);

	[entry("GdipGraphicsClear")]
	long GdipGraphicsClear(
      [in] HANDLE hGraphics,
	  [in] long lColor
	);

	[entry("GdipDrawPie")]
	long GdipDrawPie(
      [in] HANDLE hGraphics,
	  [in] HANDLE hPen,
	  [in] float fx,
	  [in] float fy,
	  [in] float fWidth,
	  [in] float fHeight,
	  [in] float fStartAngle,
	  [in] float fSweepAngle
	);

	[entry("GdipDrawArc")]
	long GdipDrawArc(
      [in] HANDLE hGraphics,
	  [in] HANDLE hPen,
	  [in] float fx,
	  [in] float fy,
	  [in] float fWidth,
	  [in] float fHeight,
	  [in] float fStartAngle,
	  [in] float fSweepAngle
	);

	[entry("GdipResetWorldTransform")]
	long GdipResetWorldTransform(
      [in] HANDLE hGraphics
	);

	[entry("GdipTranslateWorldTransform")]
	long GdipTranslateWorldTransform(
      [in] HANDLE hGraphics,
	  [in] float fdx,
	  [in] float fdy,
	  [in] long lorder
	);

	[entry("GdipRotateWorldTransform")]
	long GdipRotateWorldTransform(
      [in] HANDLE hGraphics,
	  [in] float fAngle,
	  [in] long lorder
	);



}