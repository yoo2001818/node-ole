#pragma once
#include "stdafx.h"
#include "info.h"

namespace node_ole {
	typedef struct response_create {
		dispatch_info * info;
		napi_deferred * deferred;
	} response_create;

	typedef struct response_invoke {
		func_info * funcInfo;
		std::vector<VARIANT> outValues;
		VARIANT returnValue;
		napi_deferred * deferred;
	} response_invoke;

	typedef struct response_event {
		func_info * funcInfo;
		std::vector<VARIANT> params;
	} response_event;

	enum class response_type {
		create, invoke, event
	};

	typedef struct response {
		response_type type;
		union {
			response_create create;
			response_invoke invoke;
			response_event event;
		};
	} response;

	typedef struct request_create {
		wchar_t * prgid;
		CLSID clsid;
		napi_deferred * deferred;
	} request_create;

	typedef struct request_invoke {
		LPDISPATCH * dispatch;
		func_info * funcInfo;
		std::vector<VARIANT> params;
		napi_deferred * deferred;
	} request_invoke;

	typedef struct request_gc {
		LPDISPATCH * dispatch;
	} request_gc;

	enum class request_type {
		create, invoke, gc
	};

	typedef struct request {
		request_type type;
		union {
			request_create create;
			request_invoke invoke;
			request_gc gc;
		};
	} request;
}