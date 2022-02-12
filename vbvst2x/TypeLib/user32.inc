
typedef struct MSG {
	HANDLE hwnd;
	long message;
	PTR wParam;
	PTR lParam;
	long time;
	POINT pt;
	long lPrivate;
} MSG;

typedef struct WNDCLASSEX {
	long cbSize;
	long style;
	PTR lpfnWndProc;
	long cbClsExtra;
	long cbWndExtra;
	PTR hInstance;
	HANDLE hIcon;
	HANDLE hCursor;
	HANDLE hbrBackground;
	PTR lpszMenuName;
	PTR lpszClassName;
	HANDLE hIconSm;
} WNDCLASSEX;

typedef struct PAINTSTRUCT {
	HANDLE  hdc;
	long fErase;
	RECT rcPaint;
	long fRestore;
	long fIncUpdate;
	unsigned char rgbReserved[32];
} PAINTSTRUCT;

typedef struct CWPSTRUCT {
  PTR lParam;
  PTR wParam;
  long message;
  HANDLE hwnd;
} CWPSTRUCT;

[dllname("USER32.DLL")]
module user32 {
		
		const int WM_USER = 0x0400;
		const int PM_NOREMOVE = 0;
		const int HWND_MESSAGE = -3;
		const int WS_CHILD = 0x40000000;
		const int WS_CLIPCHILDREN = 0x02000000;
		const int WS_CLIPSIBLINGS = 0x04000000;
		const int WS_SYSMENU = 0x00080000;
		const int WS_OVERLAPPED = 0;
		const int WS_MINIMIZEBOX = 0x00020000;
		const int WS_EX_NOPARENTNOTIFY = 0x00000004;
		const int WS_CAPTION = 0x00C00000;
		const int WS_EX_TOOLWINDOW = 0x00000080;
		const int SW_HIDE = 0;
		const int SW_SHOW = 5;
		const int SW_SHOWMINNOACTIVE = 7;
		const int GWLP_USERDATA = -21;
		const int GWLP_WNDPROC = -4;
		const int GWL_STYLE = -16;
		const int GWL_EXSTYLE = -20;
		const int WM_NCCREATE = 0x081;
		const int WM_PAINT  = 15;
		const int WM_CLOSE = 16;
		const int WM_ACTIVATE = 0x0006;
		const int WA_INACTIVE = 0;
		const int CWP_SKIPDISABLED = 2;
		const int CWP_SKIPINVISIBLE = 1;
		const int SM_CYSCREEN = 1;
		const int SM_CXSCREEN = 0;
		const int RDW_UPDATENOW = 0x100;
		const int RDW_INVALIDATE = 1;
		const int WH_GETMESSAGE = 3;
		const int WH_CALLWNDPROC = 4;
		const int HC_ACTION = 0;
		const int SMTO_BLOCK = 1;
		const int SMTO_NOTIMEOUTIFNOTHUNG = 8;
		const int SMTO_ERRORONEXIT = 0x20;
		const int HWND_TOP = 0;
		const int QS_KEY = 0x0001;
		const int QS_MOUSEMOVE = 0x0002;
		const int QS_MOUSEBUTTON = 0x0004;
		const int QS_MOUSE = 6;
		const int QS_POSTMESSAGE = 0x0008;
		const int QS_TIMER = 0x0010;
		const int QS_PAINT = 0x0020;
		const int QS_SENDMESSAGE = 0x0040;
		const int QS_HOTKEY = 0x0080;
		const int QS_INPUT = 0x407;
		const int QS_ALLEVENTS = 0x04BF;
		const int QS_ALLINPUT = 0x04FF;
		const int SWP_SHOWWINDOW = 0x40;
		const int SWP_NOZORDER = 0x4;

		[entry("SetWindowTextW")]
		long SetWindowText(
		  [in] HANDLE hWnd,
		  [in] LPWSTR lpString
		);

		[entry("SetForegroundWindow")]
		long SetForegroundWindow(
		  [in] HANDLE hWnd
		);

		[entry("SetWindowPos")]
		long SetWindowPos(
		  [in] HANDLE hWnd,
		  [in] HANDLE hWndInsertAfter,
		  [in] int  X,
		  [in] int  Y,
		  [in] int  cx,
		  [in] int  cy,
		  [in] long uFlags
		);

		[entry("MessageBoxW")]
		long MessageBox(
			[in] HANDLE hWnd,
			[in] LPWSTR lpText,
			[in] LPWSTR lpCaption, 
			[in] long uType
		);

		[entry("AdjustWindowRectEx")]
		long AdjustWindowRectEx(
			[in] RECT* lpRect,
			[in] long dwStyle,
			[in] long bMenu,
			[in] long dwExStyle
		);

		[entry("AdjustWindowRect")]
		long AdjustWindowRect(
			[in] RECT* lpRect,
			[in] long dwStyle,
			[in] long bMenu
		);

		[entry("FindWindowExW")]
		HANDLE FindWindowEx(
			[in] HANDLE hWndParent,
			[in] HANDLE hWndChildAfter,
			[in] LPWSTR lpszClass,
			[in] LPWSTR lpszWindow
		);

		[entry("GetWindowThreadProcessId")]
		long GetWindowThreadProcessId(
			[in] HANDLE hWnd,
			[in] void* lpdwProcessId
		);

		[entry("CreateWindowExW")]
		HANDLE CreateWindowEx(
			[in] long dwExStyle,
			[in] LPWSTR lpClassName,
			[in] LPWSTR lpWindowName,
			[in] long dwStyle,
			[in] long X,
			[in] long Y,
			[in] long nWidth,
			[in] long nHeight,
			[in] HANDLE hWndParent,
			[in] HANDLE hMenu,
			[in] PTR hInstance,
			[in] void* lpParam
		);

		[entry("DestroyWindow")]
		long DestroyWindow(
			[in] HANDLE hWnd
		);

		[entry("RegisterClassExW")]
		short RegisterClassEx(
			[in] WNDCLASSEX *pCls
		);

		[entry("GetClassInfoExW")]
		long GetClassInfoEx(
			[in] PTR hInstance,
			[in] LPWSTR lpszClass,
			[in, out] WNDCLASSEX *lpwcx
		);

		[entry("UnregisterClassW")]
		long UnregisterClass(
			[in] LPWSTR lpClassName,
			[in] PTR hInstance
		);

		[entry("DefWindowProcW")]
		PTR DefWindowProc(
			[in] HANDLE hWnd,
			[in] long Msg,
			[in] PTR wParam,
			[in] void* lParam
		);

		[entry("CallWindowProcW")]
		PTR CallWindowProc(
			[in] PTR lpPrevWndFunc,
			[in] HANDLE hWnd,
			[in] long Msg,
			[in] PTR wParam,
			[in] void* lParam
		);

		[entry("PostThreadMessageW")]
		long PostThreadMessage(
		  [in] long idThread,
		  [in] long Msg,
		  [in] PTR wParam,
		  [in] void *lParam
		);

		[entry("SendMessageTimeoutW")]
		long SendMessageTimeout(
		  [in] HANDLE hWnd,
		  [in] long Msg,
		  [in] PTR wParam,
		  [in] void *lParam,
		  [in] long fuFlags,
		  [in] long uTimeout,
		  [in] PTR* lpdwResult
		);

		[entry("SendMessageCallbackW")]
		long SendMessageCallback(
		  [in] HANDLE hWnd,
		  [in] long Msg,
		  [in] PTR wParam,
		  [in] void *lParam,
		  [in] PTR lpResultCallBack,
		  [in] PTR* dwData
		);

		[entry("SendMessageW")]
		PTR SendMessage(
		  [in] HANDLE hWnd,
		  [in] long Msg,
		  [in] PTR wParam,
		  [in] void *lParam
		);

		[entry("PostMessageW")]
		PTR PostMessage(
		  [in] HANDLE hWnd,
		  [in] long Msg,
		  [in] PTR wParam,
		  [in] void *lParam
		);

		[entry("MsgWaitForMultipleObjects")]
		long MsgWaitForMultipleObjects(
		  [in] long nCount, 
		  [in] HANDLE *pHandles,
		  [in] long fWaitAll,
		  [in] long dwMilliseconds,
		  [in] long dwWakeMask
		);

		[entry("PeekMessageW")]
		long PeekMessage(
		  [in] MSG *lpMsg,
		  [in] HANDLE hWnd,
		  [in] long wMsgFilterMin,
		  [in] long wMsgFilterMax,
		  [in] long wRemoveMsg
		);

		[entry("PostQuitMessage")]
		void PostQuitMessage(
		  [in] int nExitCode
		);

		[entry("GetMessageW")]
		long GetMessage(
		  [in] MSG *lpMsg,
		  [in] HANDLE hWnd,
		  [in] long wMsgFilterMin,
		  [in] long wMsgFilterMax
		);

		[entry("TranslateMessage")]
		long TranslateMessage(
			[in] MSG *lpMsg
		);

		[entry("DispatchMessageW")]
		long DispatchMessage(
			[in] MSG *lpMsg
		);

		[entry("ShowWindow")]
		long ShowWindow(
		  [in] HANDLE hWnd,
		  [in] long nCmdShow
		);

		[entry("MoveWindow")]
		long MoveWindow(
		  [in] HANDLE hWnd,
		  [in] int X,
		  [in] int Y,
		  [in] int nWidth,
		  [in] int nHeight,
		  [in] long bRepaint
		);

		[entry("SetWindowLongW")]
		PTR SetWindowLongPtr(
		  [in] HANDLE hWnd,
		  [in] int nIndex,
		  [in] PTR dwNewLong
		);

		[entry("GetWindowLongW")]
		PTR GetWindowLongPtr(
		  [in] HANDLE hWnd,
		  [in] int nIndex
		);

		[entry("SetParent")]
		HANDLE SetParent(
		  [in] HANDLE hWndChild,
		  [in] HANDLE  hWndNewParent
		);

		[entry("BeginPaint")]
		HANDLE BeginPaint(
		  [in]  HANDLE hWnd,
		  [in, out] PAINTSTRUCT* lpPaint
		);

		[entry("EndPaint")]
		long EndPaint(
		  [in] HANDLE hWnd,
		  [in, out] PAINTSTRUCT* lpPaint
		);

		[entry("SetCursorPos")]
		long SetCursorPos(
		  [in] long lx,
		  [in] long ly
		);

		[entry("GetCursorPos")]
		long GetCursorPos(
		  [in, out] POINT* lpPoint
		);

		[entry("ChildWindowFromPointEx")]
		HANDLE ChildWindowFromPointEx(
		  [in] HANDLE hWndParent,
		  [in] long lx,
		  [in] long ly,
		  [in] long uFlags
		);

		[entry("ScreenToClient")]
		long ScreenToClient(
		  [in] HANDLE hWnd,
		  [in] POINT* lpPoint
		);

		[entry("ClientToScreen")]
		long ClientToScreen(
		  [in] HANDLE hWnd,
		  [in] POINT* lpPoint
		);

		[entry("PtInRect")]
		long PtInRect(
		  [in] RECT *lpRect,
		  [in] long lx,
		  [in] long ly
		);

		[entry("SetWindowRgn")]
		long SetWindowRgn(
		  [in] HANDLE hWnd,
		  [in] HANDLE hRgn, 
		  [in] long bRedraw
		);

		[entry("InflateRect")]
		long InflateRect(
		  [in] RECT *lpRect,
		  [in] long lx,
		  [in] long ly
		);

		[entry("SetCapture")]
		HANDLE SetCapture(
		  [in] HANDLE hWnd
		);

		[entry("ReleaseCapture")]
		long ReleaseCapture();

		[entry("GetCapture")]
		HANDLE GetCapture();

		[entry("GetSystemMetrics")]
		HANDLE GetSystemMetrics(
		  [in] long nIndex
		);

		[entry("GetFocus")]
		HANDLE GetFocus();

		[entry("InvalidateRect")]
		long InvalidateRect(
		  [in] HANDLE hWnd,
		  [in] void* lpRect, 
		  [in] long bErase
		);

		[entry("SetRect")]
		long SetRect(
		  [in] RECT *lpRect,
		  [in] long x1,
		  [in] long y1,
		  [in] long x2,
		  [in] long y2
		);

		[entry("GetDC")]
		HANDLE GetDC(
		  [in] HANDLE hWnd
		);

		[entry("ReleaseDC")]
		long ReleaseDC(
		  [in] HANDLE hWnd,
		  [in] HANDLE hdc
		);

		[entry("RedrawWindow")]
		long RedrawWindow(
		  [in] HANDLE hWnd,
		  [in] void* lprcUpdate,
		  [in] HANDLE hrgnUpdate,
		  [in] long fuRedraw
		);

		[entry("ScrollDC")]
		long ScrollDC(
		  [in] HANDLE hdc,
		  [in] long dx,
		  [in] long dy,
		  [in] RECT* lprcScroll,
		  [in] RECT* lprcClip,
		  [in] HANDLE hrgnUpdate,
		  [in] RECT* lprcUpdate
		);

		[entry("GetParent")]
		HANDLE GetParent(
		  [in] HANDLE hWnd
		);

		[entry("OffsetRect")]
		long OffsetRect(
		  [in] RECT *lpRect,
		  [in] long lx,
		  [in] long ly
		);

		[entry("IntersectRect")]
		long IntersectRect(
		  [in] RECT *lpRect,
		  [in] RECT *lprcSrc1,
		  [in] RECT *lprcSrc2
		);

		[entry("SubtractRect")]
		long SubtractRect(
		  [in] RECT *lpRect,
		  [in] RECT *lprcSrc1,
		  [in] RECT *lprcSrc2
		);

		[entry("UnionRect")]
		long UnionRect(
		  [in] RECT *lpRect,
		  [in] RECT *lprcSrc1,
		  [in] RECT *lprcSrc2
		);

		[entry("IsRectEmpty")]
		long IsRectEmpty(
		  [in] RECT *lpRect
		);

		[entry("DrawTextW")]
		long DrawText(
		  [in] HANDLE hdc,
		  [in] LPWSTR lpStr,
		  [in] long nCount,
		  [in] RECT* lpRect,
		  [in] long wFormat
		);

		[entry("SetWindowsHookExW")]
		HANDLE SetWindowsHookEx(
		  [in] long idHook,
		  [in] PTR lpfn,
		  [in] PTR hmod,
		  [in] long dwThreadId
		);

		[entry("UnhookWindowsHookEx")]
		long UnhookWindowsHookEx(
		  [in] HANDLE hhk
		);

		[entry("CallNextHookEx")]
		long CallNextHookEx(
		  [in] HANDLE hhk,
		  [in] long nCode,
		  [in] PTR wParam,
		  [in] PTR lParam
		);

		[entry("RegisterWindowMessageW")]
		long RegisterWindowMessage(
		  [in] LPWSTR lpString
		);


}
