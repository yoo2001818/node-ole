#pragma once
#include "stdafx.h"

namespace node_ole {
	void setOleWindowIUnknown(LPUNKNOWN lpunk);
	void createOleWindow(HWND * output);
	LRESULT CALLBACK oleWindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
}