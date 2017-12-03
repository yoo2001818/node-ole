#pragma once
#include "stdafx.h"
#include "environment.h"
#include "handler.h"

static std::vector<node_ole::Environment *> instances;

namespace node_ole {
	Environment::Environment() {
		// Create both side of handle.
		uv_loop_t * loop = uv_default_loop();
		uv_async_init(loop, &nodeHandle, nodeHandler);
		workerHandle = CreateEvent(NULL, FALSE, FALSE, NULL);
		workerExitHandle = CreateEvent(NULL, TRUE, FALSE, NULL);
		// Initialize worker thread.
		workerThread = CreateThread(NULL, 0, workerHandler, (void *)this, 0, &workerId);
		instances.push_back(this);
	}

	Environment::~Environment() {
		close();
		instances.erase(std::remove(instances.begin(), instances.end(), this), instances.end());
	}

	void Environment::close() {
		// Shutdown
		uv_close((uv_handle_t *)&nodeHandle, NULL);
		SetEvent(workerExitHandle);
		WaitForSingleObject(workerThread, INFINITE);
		CloseHandle(workerHandle);
		CloseHandle(workerExitHandle);
	}

	void Environment::pushRequest(std::unique_ptr<Request> req) {
		{
			std::lock_guard<std::mutex> lock(reqGuard);
			reqQueue.push(std::move(req));
		}
		SetEvent(workerHandle);
	}

	void Environment::pushResponse(std::unique_ptr<Response> res) {
		{
			std::lock_guard<std::mutex> lock(resGuard);
			resQueue.push(std::move(res));
		}
		uv_async_send(&nodeHandle);
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
		std::unique_ptr<RequestCreate> req = std::make_unique<RequestCreate>();
		req->deferred = deferred;
		req->name = std::wstring((wchar_t *) **value);

		env->pushRequest(std::move(req));
		delete value;

		info.GetReturnValue().Set(promise);
	}

	NAN_MODULE_INIT(Environment::Init) {
		v8::Local<v8::FunctionTemplate> tpl = Nan::New<v8::FunctionTemplate>(New);
		tpl->SetClassName(Nan::New("Environment").ToLocalChecked());
		tpl->InstanceTemplate()->SetInternalFieldCount(1);

		Nan::SetPrototypeMethod(tpl, "create", Create);

		Nan::Set(target, Nan::New("Environment").ToLocalChecked(), Nan::GetFunction(tpl).ToLocalChecked());
	}

	void Environment::Cleanup() {
		for (auto it = instances.begin(); it != instances.end(); it++) {
			(*it)->close();
		}
	}
}
