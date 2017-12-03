#include "stdafx.h"
#include "environment.h"
#include "handler.h"

namespace node_ole {

	NAN_MODULE_INIT(initAll) {
		Environment::Init(target);
	}

	NODE_MODULE(node_ole, initAll);
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