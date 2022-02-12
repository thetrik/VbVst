
typedef struct POINT {
	long x;
	long y;
} POINT;

typedef struct RECT {
	long Left;
	long Top;
	long Right;
	long Bottom;
} RECT;

typedef struct RGBQUAD {
	unsigned char rgbBlue;
	unsigned char rgbGreen;
	unsigned char rgbRed;
	unsigned char rgbReserved;
} RGBQUAD;

typedef struct BITMAPINFOHEADER {
	long biSize;
	long biWidth;
	long biHeight;
	short biPlanes;
	short biBitCount;
	long biCompression;
	long biSizeImage;
	long biXPelsPerMeter;
	long biYPelsPerMeter;
	long biClrUsed;
	long biClrImportant;
} BITMAPINFOHEADER;

typedef struct BITMAPINFO {
	BITMAPINFOHEADER bmiHeader;
	RGBQUAD  bmiColors[1];
} BITMAPINFO;

typedef struct SIZE {
	long cx;
	long cy;
} SIZE;


[dllname("GDI32.DLL")]
module gdi32 {

	const int DC_PEN = 19;
	const int DC_BRUSH = 18;
	const int NULL_BRUSH = 5;
	const int DT_CENTER = 0x1;
	const int DT_END_ELLIPSIS = 0x8000;
	const int DT_CALCRECT = 0x400;
	const int DT_SINGLELINE = 0x20;
	const int DT_VCENTER = 0x4;
	const int DT_RIGHT = 0x2;

	[entry("TextOutW")]
	long TextOut(
	  [in] HANDLE hdc,
	  [in] int x,
	  [in] int y,
	  [in] LPWSTR lpString,
	  [in] int c
	);

	[entry("GetTextExtentPoint32W")]
	long GetTextExtentPoint32(
	  [in] HANDLE hdc,
	  [in] LPWSTR lpString,
	  [in] int cbString,
	  [in] SIZE *lpSize
	);

	[entry("IntersectClipRect")]
	int IntersectClipRect(
	  [in] HANDLE hdc,
	  [in] int left,
	  [in] int top,
	  [in] int right,
	  [in] int bottom
	);

	[entry("SelectClipRgn")]
	int SelectClipRgn(
	  [in] HANDLE hdc,
	  [in] HANDLE hrgn
	);

	[entry("CreateCompatibleBitmap")]
	HANDLE CreateCompatibleBitmap(
	  [in] HANDLE hdc,
	  [in] long nWidth,
	  [in] long nHeight
	);

	[entry("CreateCompatibleDC")]
	HANDLE CreateCompatibleDC(
	  [in] HANDLE hdc
	);

	[entry("SelectObject")]
	HANDLE SelectObject(
	  [in] HANDLE hdc,
	  [in] HANDLE hObject
	);

	[entry("DeleteObject")]
	long DeleteObject(
	  [in] HANDLE hObject
	);

	[entry("DeleteDC")]
	long DeleteDC(
	  [in] HANDLE hdc
	);

	[entry("CreateDIBSection")]
	HANDLE CreateDIBSection(
	  [in] HANDLE hdc,
	  [in] BITMAPINFO* pBitmapInfo,
	  [in] long usage,
	  [in] void* ppvBits,
	  [in] HANDLE hSection,
	  [in] long offset
	);

	[entry("LineTo")]
	long LineTo(
	  [in] HANDLE hdc,
	  [in] long x,
	  [in] long y
	);

	[entry("MoveToEx")]
	long MoveToEx(
	  [in] HANDLE hdc,
	  [in] long x,
	  [in] long y,
	  [in] void* lpPoint
	);

	[entry("Rectangle")]
	long Rectangle(
	  [in] HANDLE hdc,
	  [in] long x1,
	  [in] long y1,
	  [in] long x2,
	  [in] long y2
	);

	[entry("SetTextColor")]
	long SetTextColor(
	  [in] HANDLE hdc,
	  [in] long crColor
	);

	[entry("GetStockObject")]
	HANDLE GetStockObject(
	  [in] long lObject
	);

	[entry("SetDCPenColor")]
	HANDLE SetDCPenColor(
	  [in] HANDLE hdc,
	  [in] long crColor
	);

	[entry("SetDCBrushColor")]
	HANDLE SetDCBrushColor(
	  [in] HANDLE hdc,
	  [in] long crColor
	);

	[entry("PatBlt")]
	long PatBlt(
	  [in] HANDLE hdc,
	  [in] long x,
	  [in] long y,
	  [in] long nWidth,
	  [in] long nHeight,
	  [in] long dwRop
	);

	[entry("SaveDC")]
	long SaveDC(
	  [in] HANDLE hdc
	);

	[entry("RestoreDC")]
	long RestoreDC(
	  [in] HANDLE hdc,
	  [in] long nSavedDC
	);

	[entry("RoundRect")]
	long RoundRect(
	  [in] HANDLE hdc,
	  [in] long x1,
	  [in] long y1,
	  [in] long x2,
	  [in] long y2,
	  [in] long x3,
	  [in] long y3
	);

	[entry("CreateRoundRectRgn")]
	HANDLE CreateRoundRectRgn(
	  [in] long x1,
	  [in] long y1,
	  [in] long x2,
	  [in] long y2,
	  [in] long x3,
	  [in] long y3
	);

	[entry("OffsetClipRgn")]
	long OffsetClipRgn(
	  [in] HANDLE hdc,
	  [in] long x,
	  [in] long y
	);

	[entry("BitBlt")]
	long BitBlt(
		[in] HANDLE hdcDest,
		[in] long xDest,
		[in] long yDest,
		[in] long WidthDest,
		[in] long HeightDest,
		[in] HANDLE hdcSrc,
		[in] long xSrc,
		[in] long ySrc,
		[in] long rop
	);

}