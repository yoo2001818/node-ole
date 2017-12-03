#pragma once
#include "stdafx.h"
#include "environment.h"

namespace node_ole {
	DWORD WINAPI workerHandler(LPVOID lpParam);

	void handleWorkerEvent(Environment * env);

	void nodeHandler(uv_async_t * handle);

}