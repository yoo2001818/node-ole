#include "olewindow.h"
#include <tchar.h>
#include "dummyClientSite.h"

namespace node_ole {
	LPUNKNOWN constLpUnk = NULL;
	void setOleWindowIUnknown(LPUNKNOWN lpunk) {
		constLpUnk = lpunk;
	}
	void createOleWindow(HWND * output) {
		const wchar_t CLASS_NAME[] = L"node-ole";

		WNDCLASS wc = {};

		wc.lpfnWndProc = oleWindowProc;
		wc.hInstance = GetModuleHandle(NULL);
		wc.lpszClassName = CLASS_NAME;

		RegisterClass(&wc);

		// Create the window.

		HWND hwnd = CreateWindowEx(
			0,                              // Optional window styles.
			CLASS_NAME,                     // Window class
			L"node-ole window",    // Window text
			WS_OVERLAPPEDWINDOW,            // Window style

											// Size and position
			CW_USEDEFAULT, CW_USEDEFAULT, 640, 480,
			NULL,       // Parent window    
			NULL,       // Menu
			GetModuleHandle(NULL),  // Instance handle
			NULL        // Additional application data
		);

		*output = hwnd;

		if (hwnd == NULL) return;

		ShowWindow(hwnd, SW_SHOWNORMAL);
	}
	int count = 0;
	BOOL checkContinue(ULONG_PTR dwContinue) {
		return TRUE;
	}
	LRESULT CALLBACK oleWindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam) {
		switch (uMsg) {
		case WM_DESTROY:
			PostQuitMessage(0);
			return 0;

		case WM_PAINT:
		{
			PAINTSTRUCT ps;
			HDC hdc = BeginPaint(hwnd, &ps);
			RECT rect = { 0, 0, 640, 480 };
			TCHAR greeting[] = _T(L"Hello, world!");
			count++;
			TextOut(hdc,
				count, 5,
				greeting, 13);
			if (constLpUnk != NULL) {
				LPVIEWOBJECT viewObj;
				constLpUnk->QueryInterface(&viewObj);
				HRESULT result = viewObj->Draw(DVASPECT_CONTENT, 0, NULL, NULL, NULL,
					hdc, (LPCRECTL) &rect, NULL, checkContinue, 1);
				printf("%X\n", result);
				// OleDraw(constLpUnk, DVASPECT_CONTENT, hdc, &rect);
			}

			EndPaint(hwnd, &ps);
		}
		return 0;

		}
		return DefWindowProc(hwnd, uMsg, wParam, lParam);
	}
}