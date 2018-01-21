#pragma once
#include "stdafx.h"
#include "event.h"

namespace node_ole {
	class Environment : public Nan::ObjectWrap {
	public:
		Environment();
		virtual ~Environment();
		
		void close();

		static NAN_MODULE_INIT(Init);
		static NAN_METHOD(New);
		static NAN_METHOD(Create);
		static NAN_METHOD(Close);
		static NAN_METHOD(Cleanup);

		void pushRequest(std::unique_ptr<Request> req);
		void pushResponse(std::unique_ptr<Response> res);

		void ref();
		void unref();

		HANDLE workerThread;
		DWORD workerId;

		std::queue<std::unique_ptr<Request>> reqQueue;
		std::queue<std::unique_ptr<Response>> resQueue;

		std::mutex reqGuard;
		std::mutex resGuard;

		HANDLE workerHandle;
		uv_async_t nodeHandle;

		HANDLE workerExitHandle;
	};
}