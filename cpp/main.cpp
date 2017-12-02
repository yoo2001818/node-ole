#include "stdafx.h"
#include "environment.h"
#include "handler.h"

namespace node_ole {

	napi_value init(napi_env env, napi_value exports) {
		// Start OLE environment.
		Environment * environment = new Environment(env);
		// Add an entry point - this is used to start all the OLE objects.
		napi_value func;
		napi_status status;
		status = napi_create_function(env, u8"createOLEObject",
			NAPI_AUTO_LENGTH, nodeInitHandler, environment, &func);
		if (status != napi_ok) return nullptr;
		return func;
	}

}

NAPI_MODULE(node_ole, node_ole::init);

BOOL WINAPI DllMain(HINSTANCE hinstDll, DWORD fdwReason, LPVOID lpvReserved) {
	switch (fdwReason) {
	case DLL_PROCESS_ATTACH:
		break;
	case DLL_PROCESS_DETACH:
		// Stop worker thread
		break;
	case DLL_THREAD_ATTACH:
		break;
	case DLL_THREAD_DETACH:
		break;
	}
	return TRUE;
}