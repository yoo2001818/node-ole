#pragma once
#include "stdafx.h"
#include "info.h"

namespace node_ole {
	using ResolverPersistent = Nan::Persistent<v8::Promise::Resolver>;
	using FunctionPersistent = Nan::Persistent<v8::Function>;

	typedef struct ResponseCreate {
		dispatch_info * info;
		ResolverPersistent * deferred;
	} ResponseCreate;

	typedef struct ResponseInvoke {
		func_info * funcInfo;
		std::vector<VARIANT> outValues;
		VARIANT returnValue;
		ResolverPersistent * deferred;
	} ResponseInvoke;

	typedef struct ResponseEvent {
		func_info * funcInfo;
		std::vector<VARIANT> params;
		FunctionPersistent * eventCallback;
	} ResponseEvent;

	enum class ResponseType {
		Create, Invoke, Event
	};

	typedef struct Response {
		ResponseType type;
		union {
			ResponseCreate create;
			ResponseInvoke invoke;
			ResponseEvent event;
		};
		Response() {}
		Response(Response&& o): type(std::move(o.type)) {
			switch (type) {
			case ResponseType::Create:
				create = std::move(o.create);
				break;
			case ResponseType::Invoke:
				invoke = std::move(o.invoke);
				break;
			case ResponseType::Event:
				event = std::move(o.event);
				break;
			}
		}
		~Response() {
			switch (type) {
			case ResponseType::Create:
				create.~ResponseCreate();
				break;
			case ResponseType::Invoke:
				invoke.~ResponseInvoke();
				break;
			case ResponseType::Event:
				event.~ResponseEvent();
				break;
			}
		}
	} Response;

	typedef struct RequestCreate {
		std::wstring name;
		ResolverPersistent * deferred;
		FunctionPersistent * eventCallback;
	} RequestCreate;

	typedef struct RequestInvoke {
		LPUNKNOWN dispatch;
		func_info * funcInfo;
		std::vector<VARIANT> params;
		ResolverPersistent * deferred;
	} RequestInvoke;

	typedef struct RequestGC {
		LPUNKNOWN dispatch;
	} RequestGC;

	enum class RequestType {
		Create, Invoke, GC
	};

	typedef struct Request {
		RequestType type;
		union {
			RequestCreate create;
			RequestInvoke invoke;
			RequestGC gc;
		};
		Request() {}
		Request(Request&& o) : type(std::move(o.type)) {
			switch (type) {
			case RequestType::Create:
				create = std::move(o.create);
				break;
			case RequestType::Invoke:
				invoke = std::move(o.invoke);
				break;
			case RequestType::GC:
				gc = std::move(o.gc);
				break;
			}
		}
		~Request() {
			switch (type) {
			case RequestType::Create:
				create.~RequestCreate();
				break;
			case RequestType::Invoke:
				invoke.~RequestInvoke();
				break;
			case RequestType::GC:
				gc.~RequestGC();
				break;
			}
		}
	} Request;
}