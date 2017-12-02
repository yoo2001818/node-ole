#pragma once
#include "stdafx.h"
#include "event.h"

namespace node_ole {
	class Environment {
	public:
		Environment(napi_env env);
		~Environment();

		void pushRequest(request req);
		void pushResponse(response res);

		napi_env env;
		HANDLE workerThread;
		DWORD workerId;

		std::vector<request> reqQueue;
		std::vector<response> resQueue;

		HANDLE workerHandle;
		uv_async_t nodeHandle;

		HANDLE workerExitHandle;
	};
}