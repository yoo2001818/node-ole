#include "stdafx.h"
#include "handler.h"
#include "environment.h"
#include "event.h"
#include "comutil.h"
#include "oleobject.h"

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
				std::cout << "COM received!!!" << std::endl;
				handleWorkerEvent(env);
				break;
			}
			case WAIT_OBJECT_0 + 1: {
				printf("COM Thread shutdown; %d\n", GetCurrentThreadId());
				CoUninitialize();
				return 0;
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
				DispatchInfo * dispatchInfo = nullptr;
				HRESULT result;
				result = initObject(r->name.data(), &lpunk);
				if SUCCEEDED(result) {
					// Retrieve type information.
					result = getObjectInfo(lpunk, &dispatchInfo);
					// Initialize event listeners. TODO
				}

				// Send the response back to the node.
				std::unique_ptr<ResponseCreate> res = std::make_unique<ResponseCreate>();
				res->deferred = r->deferred;
				res->info = dispatchInfo;
				res->result = result;
				
				env->pushResponse(std::move(res));
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
		printf("Node received!! \n");
		while (true) {
			std::unique_ptr<Response> res;
			{
				std::lock_guard<std::mutex> lock(env->resGuard);
				if (env->resQueue.empty()) break;
				res = std::move(env->resQueue.front());
				env->resQueue.pop();
			}
			switch (res->getType()) {
			case ResponseType::Create: {
				Nan::HandleScope scope;
				ResponseCreate * r = static_cast<ResponseCreate*>(res.get());
				printf("Resolving ResponseCreate\n");
				auto resolver = Nan::New(*(r->deferred));
				if (r->result == S_OK) {
					// Unwrap the information object into an object - the object should wrap a
					// DispatchObject, while the function should contain FuncInfo.
					OLEObject * oleObj = new OLEObject(env, r->info);
					auto tpl = Nan::New<v8::ObjectTemplate>();
					tpl->SetInternalFieldCount(1);
					auto object = tpl->NewInstance();
					oleObj->bake(object);
					resolver->Resolve(object);
				} else {
					// Throw an error.
					_com_error err(r->result);
					printf("%X\n", err.Error());
					resolver->Reject(Nan::Error(Nan::New((uint16_t *) err.ErrorMessage()).ToLocalChecked()));
				}
				v8::Isolate::GetCurrent()->RunMicrotasks();
				r->deferred->Reset();
				delete r->deferred;
				break;
			}
			case ResponseType::Invoke: {
				ResponseInvoke * r = static_cast<ResponseInvoke*>(res.get());
				break;
			}
			case ResponseType::Event: {
				ResponseEvent * r = static_cast<ResponseEvent*>(res.get());
				break;
			}
			}
		}
	}
}