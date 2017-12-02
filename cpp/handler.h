#pragma once
#include "stdafx.h"

DWORD WINAPI workerHandler(LPVOID lpParam);

void nodeHandler(uv_async_t * handle);