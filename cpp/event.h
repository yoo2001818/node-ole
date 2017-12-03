#pragma once
#include "stdafx.h"
#include "info.h"

namespace node_ole {
	using ResolverPersistent = Nan::Persistent<v8::Promise::Resolver>;
	using FunctionPersistent = Nan::Persistent<v8::Function>;

	class ResponseCreate {
	public:
		dispatch_info * info;
		ResolverPersistent * deferred;
	};

	class ResponseInvoke {
	public:
		func_info * funcInfo;
		std::vector<VARIANT> outValues;
		VARIANT returnValue;
		ResolverPersistent * deferred;
	};

	class ResponseEvent {
	public:
		func_info * funcInfo;
		std::vector<VARIANT> params;
		FunctionPersistent * eventCallback;
	};

	enum class ResponseType {
		Create, Invoke, Event
	};

	class Response {
	public:
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

		ResponseType type;
		union {
			ResponseCreate create;
			ResponseInvoke invoke;
			ResponseEvent event;
		};
	};

	class RequestCreate {
	public:
		std::wstring name;
		ResolverPersistent * deferred;
		FunctionPersistent * eventCallback;
	};

	class RequestInvoke {
	public:
		LPUNKNOWN dispatch;
		func_info * funcInfo;
		std::vector<VARIANT> params;
		ResolverPersistent * deferred;
	};

	class RequestGC {
	public:
		LPUNKNOWN dispatch;
	};

	enum class RequestType {
		Create, Invoke, GC
	};

	class Request {
	public:
		Request(RequestType type) : type(type) {
			switch (type) {
			case RequestType::Create:
				new (&create) RequestCreate();
				break;
			case RequestType::Invoke:
				new (&invoke) RequestInvoke();
				break;
			case RequestType::GC:
				new (&gc) RequestGC();
				break;
			}
		}
		Request(Request& o) : type(o.type) {
			switch (type) {
			case RequestType::Create:
				new (&create) RequestCreate(o.create);
				break;
			case RequestType::Invoke:
				new (&invoke) RequestInvoke(o.invoke);
				break;
			case RequestType::GC:
				new (&gc) RequestGC(o.gc);
				break;
			}
		}
		Request(Request&& o) : type(std::move(o.type)) {
			switch (type) {
			case RequestType::Create:
				new (&create) RequestCreate(std::move(o.create));
				break;
			case RequestType::Invoke:
				new (&invoke) RequestInvoke(std::move(o.invoke));
				break;
			case RequestType::GC:
				new (&gc) RequestGC(std::move(o.gc));
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

		RequestType type;
		union {
			RequestCreate create;
			RequestInvoke invoke;
			RequestGC gc;
		};
	};
}