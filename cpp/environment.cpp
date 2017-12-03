#pragma once
#include "stdafx.h"
#include "environment.h"
#include "handler.h"

namespace node_ole {
	Environment::Environment() {
		// Create both side of handle.
		uv_loop_t * loop = uv_default_loop();
		uv_async_init(loop, &nodeHandle, nodeHandler);
		workerHandle = CreateEvent(NULL, FALSE, FALSE, NULL);
		workerExitHandle = CreateEvent(NULL, TRUE, FALSE, NULL);
		// Initialize worker thread.
		workerThread = CreateThread(NULL, 0, workerHandler, (void *)this, 0, &workerId);
	}

	Environment::~Environment() {
		// Shutdown
		uv_close((uv_handle_t *) &nodeHandle, NULL);
		SetEvent(workerExitHandle);
		WaitForSingleObject(workerThread, INFINITE);
		CloseHandle(workerHandle);
		CloseHandle(workerExitHandle);
	}

	void Environment::pushRequest(Request req) {

	}

	void Environment::pushResponse(Response req) {

	}

	NAN_METHOD(Environment::New) {
		if (!info.IsConstructCall()) {
			Nan::ThrowTypeError("This function must be called with 'new'.");
			return;
		}
		// Start the environment
		Environment * env = new Environment();
		env->Wrap(info.This());
		info.GetReturnValue().Set(info.This());
	}

	NAN_METHOD(Environment::Create) {
		if (info.Length() < 1 || !info[0]->IsString()) {
			Nan::ThrowTypeError("This function accepts a string");
			return;
		}
		Environment * env = Unwrap<Environment>(info.Holder());

		// NOTE: utf16_t -> wchar_t will break in other environments; however, this library only supports
		// Windows anyway.
		auto value = new v8::String::Value(info[0]);

		// Create Promise object.
		v8::Local<v8::Promise::Resolver> resolver = v8::Promise::Resolver::New(info.GetIsolate());
		v8::Local<v8::Promise> promise = resolver->GetPromise();
		ResolverPersistent * deferred = new ResolverPersistent(resolver);

		// Create event object, then push it.
		Request req;
		req.type = RequestType::Create;
		req.create.deferred = deferred;
		req.create.name = std::wstring((wchar_t *) **value);

		env->pushRequest(std::move(req));
		delete value;
	}

	NAN_MODULE_INIT(Environment::Init) {
		v8::Local<v8::FunctionTemplate> tpl = Nan::New<v8::FunctionTemplate>(New);
		tpl->SetClassName(Nan::New("Environment").ToLocalChecked());
		tpl->InstanceTemplate()->SetInternalFieldCount(1);

		Nan::SetPrototypeMethod(tpl, "create", Create);

		Nan::Set(target, Nan::New("Environment").ToLocalChecked(), Nan::GetFunction(tpl).ToLocalChecked());
	}
}
