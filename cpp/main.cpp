#include "stdafx.h"

namespace node_ole {

	napi_value Init(napi_env env, napi_value exports) {
		// Start OLE worker thread - use MTA if possible.
		return exports;
	}

	NAPI_MODULE(node_ole, Init);

}

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