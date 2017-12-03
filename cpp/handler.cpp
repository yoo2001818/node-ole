#include "stdafx.h"
#include "handler.h"
#include "environment.h"

namespace node_ole {
	DWORD WINAPI workerHandler(LPVOID lpParam) {
		Environment * env = (Environment *)lpParam;
		HANDLE handles[] = {
			env->workerHandle,
			env->workerExitHandle
		};
		HRESULT hresult;
		// Start COM environment
		hresult = CoInitializeEx(NULL, COINIT_MULTITHREADED);
		if FAILED(hresult) return 1;
		while (TRUE) {
			switch (MsgWaitForMultipleObjectsEx(2, handles, INFINITE,
				QS_ALLINPUT | QS_ALLPOSTMESSAGE, MWMO_INPUTAVAILABLE)) {
			case WAIT_OBJECT_0: {
				// TODO Handle event queue
				break;
			}
			case WAIT_OBJECT_0 + 1: {
				// TODO Handle exit
				break;
			}
			case WAIT_OBJECT_0 + 2: {
				MSG msg;
				while (PeekMessage(&msg, NULL, 0, 0, PM_REMOVE)) {
					TranslateMessage(&msg);
					DispatchMessage(&msg);
				}
				break;
			}
			case WAIT_TIMEOUT:
			case WAIT_FAILED:
			default:
				break;
			}
		}
	}

	void nodeHandler(uv_async_t * handle) {
		Environment * env = (Environment *)handle->data;
		// TODO Handle event queue
	}
}