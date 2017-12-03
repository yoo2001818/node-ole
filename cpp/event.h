#pragma once
#include "stdafx.h"
#include "info.h"

namespace node_ole {
	using ResolverPersistent = Nan::Persistent<v8::Promise::Resolver>;
	using FunctionPersistent = Nan::Persistent<v8::Function>;

	enum class ResponseType {
		Create, Invoke, Event
	};

	class Response {
	public:
		virtual ~Response() = default;
		virtual ResponseType getType() = 0;
	};

	class ResponseCreate: public Response {
	public:
		virtual ~ResponseCreate() = default;
		virtual ResponseType getType() { return ResponseType::Create; };
		dispatch_info * info;
		ResolverPersistent * deferred;
	};

	class ResponseInvoke: public Response {
	public:
		virtual ~ResponseInvoke() = default;
		virtual ResponseType getType() { return ResponseType::Invoke; };
		func_info * funcInfo;
		std::vector<VARIANT> outValues;
		VARIANT returnValue;
		ResolverPersistent * deferred;
	};

	class ResponseEvent: public Response {
	public:
		virtual ~ResponseEvent() = default;
		virtual ResponseType getType() { return ResponseType::Event; };
		func_info * funcInfo;
		std::vector<VARIANT> params;
		FunctionPersistent * eventCallback;
	};

	enum class RequestType {
		Create, Invoke, GC
	};

	class Request {
	public:
		virtual ~Request() = default;
		virtual RequestType getType() = 0;
	};

	class RequestCreate: public Request {
	public:
		virtual ~RequestCreate() = default;
		virtual RequestType getType() { return RequestType::Create; };
		std::wstring name;
		ResolverPersistent * deferred;
		FunctionPersistent * eventCallback;
	};

	class RequestInvoke : public Request {
	public:
		virtual ~RequestInvoke() = default;
		virtual RequestType getType() { return RequestType::Invoke; };
		LPUNKNOWN dispatch;
		func_info * funcInfo;
		std::vector<VARIANT> params;
		ResolverPersistent * deferred;
	};

	class RequestGC : public Request {
	public:
		virtual ~RequestGC() = default;
		virtual RequestType getType() { return RequestType::GC; };
		LPUNKNOWN dispatch;
	};
}