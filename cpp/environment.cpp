#pragma once
#include "stdafx.h"
#include "environment.h"
#include "handler.h"

namespace node_ole {
	Environment::Environment(napi_env env):env(env) {
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

	void Environment::pushRequest(request req) {

	}

	void Environment::pushResponse(response req) {

	}
}
