#include "stdafx.h"
#include "handler.h"
#include "environment.h"
#include "event.h"
#include "comutil.h"

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
		printf("COM Thread initialized; %d\n", GetCurrentThreadId());
		while (TRUE) {
			switch (MsgWaitForMultipleObjectsEx(2, handles, INFINITE,
				QS_ALLINPUT | QS_ALLPOSTMESSAGE, MWMO_INPUTAVAILABLE)) {
			case WAIT_OBJECT_0: {
				// Handle event queue
				std::cout << "Something received!!!" << std::endl;
				handleWorkerEvent(env);
				break;
			}
			case WAIT_OBJECT_0 + 1: {
				printf("COM Thread shutdown; %d\n", GetCurrentThreadId());
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

	void handleWorkerEvent(Environment * env) {
		while (true) {
			std::unique_ptr<Request> req;
			{
				std::lock_guard<std::mutex> lock(env->reqGuard);
				if (env->reqQueue.empty()) break;
				req = std::move(env->reqQueue.front());
				env->reqQueue.pop();
			}
			switch (req->getType()) {
			case RequestType::Create: {
				RequestCreate * r = static_cast<RequestCreate*>(req.get());
				std::wcout << r->name << std::endl;
				// Initialize COM
				LPUNKNOWN lpunk;
				DispatchInfo * dispatchInfo;
				initObject(r->name.data(), &lpunk);
				// Retrieve type information.
				getObjectInfo(lpunk, &dispatchInfo);
				// Initialize event listeners.
				break;
			}
			case RequestType::Invoke: {
				RequestInvoke * r = static_cast<RequestInvoke*>(req.get());
				break;
			}
			case RequestType::GC: {
				RequestGC * r = static_cast<RequestGC*>(req.get());
				break;
			}
			}
		}
	}

	void nodeHandler(uv_async_t * handle) {
		Environment * env = (Environment *)handle->data;
		// TODO Handle event queue
	}
}