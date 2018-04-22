#include "stdafx.h"
#include "handler.h"
#include "environment.h"
#include "event.h"
#include "comutil.h"
#include "oleobject.h"
#include "nodeutil.h"

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
				// Handle event queue
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
				// Initialize COM
				LPUNKNOWN lpunk;
				DispatchInfo * dispatchInfo = nullptr;
				HRESULT result;
				result = initObject(r->name.data(), &lpunk);
				if SUCCEEDED(result) {
					// Retrieve type information.
					result = getObjectInfo(env, lpunk, &dispatchInfo, true);
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
				// Invoke the dispatch object
				DISPPARAMS * params = r->params;
				VARIANT * output = new VARIANT();
				EXCEPINFO excepInfo;
				UINT argErr;
				HRESULT result = S_OK;
				result = r->dispatch->Invoke(r->funcInfo->dispId, IID_NULL, NULL,
					r->funcInfo->invokeKind, r->params, output, &excepInfo, &argErr);
				// Parse variant information.
				parseVariantTypeInfo(env, output);
				// Send the response back to the node.
				std::unique_ptr<ResponseInvoke> res = std::make_unique<ResponseInvoke>();
				res->deferred = r->deferred;
				res->returnValue = output;
				res->params = params;
				res->funcInfo = r->funcInfo;
				res->result = result;

				env->pushResponse(std::move(res));
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
				auto resolver = Nan::New(*(r->deferred));
				if (r->result == S_OK) {
					// Unwrap the information object into an object - the object should wrap a
					// DispatchObject, while the function should contain FuncInfo.
					OLEObject * oleObj = new OLEObject(env, r->info);
					auto tpl = Nan::New<v8::ObjectTemplate>();
					tpl->SetInternalFieldCount(1);
					auto object = tpl->NewInstance();
					r->info->persistent = new Nan::Persistent<v8::Object>(object);
					oleObj->bake(object);
					resolver->Resolve(object);
				} else {
					// Throw an error.
					_com_error err(r->result);
					resolver->Reject(Nan::Error(Nan::New((uint16_t *) err.ErrorMessage()).ToLocalChecked()));
				}
				v8::Isolate::GetCurrent()->RunMicrotasks();
				r->deferred->Reset();
				delete r->deferred;
				break;
			}
			case ResponseType::Invoke: {
				Nan::HandleScope scope;
				ResponseInvoke * r = static_cast<ResponseInvoke*>(res.get());
				auto resolver = Nan::New(*(r->deferred));
				if (r->result == S_OK) {
					// Resolve the returned object, and free the params
					resolver->Resolve(readVariant(r->returnValue, *env));
					freeVariant(r->returnValue);
					freeDispParams(r->params);
				} else {
					// Throw an error.
					_com_error err(r->result);
					resolver->Reject(Nan::Error(Nan::New((uint16_t *)err.ErrorMessage()).ToLocalChecked()));
				}
				v8::Isolate::GetCurrent()->RunMicrotasks();
				r->deferred->Reset();
				delete r->deferred;
				break;
			}
			case ResponseType::Event: {
				Nan::HandleScope scope;
				ResponseEvent * r = static_cast<ResponseEvent*>(res.get());
				if (r->info->persistent == NULL) break;
				auto object = Nan::New(*(r->info->persistent));
				OLEObject * oleObj = OLEObject::Unwrap<OLEObject>(object);
				oleObj->handleEvent(object, r->funcInfo, r->params);
				break;
			}
			}
		}
	}
}