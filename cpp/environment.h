#pragma once
#include "stdafx.h"
#include "event.h"

namespace node_ole {
	class Environment : public Nan::ObjectWrap {
	public:
		Environment();
		virtual ~Environment();

		static NAN_MODULE_INIT(Init);
		static NAN_METHOD(New);
		static NAN_METHOD(Create);

		void pushRequest(Request req);
		void pushResponse(Response res);

		HANDLE workerThread;
		DWORD workerId;

		std::vector<Request> reqQueue;
		std::vector<Response> resQueue;

		HANDLE workerHandle;
		uv_async_t nodeHandle;

		HANDLE workerExitHandle;
	};
}