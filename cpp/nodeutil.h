#pragma once
#include "stdafx.h"
#include "info.h"

namespace node_ole {
	void constructFuncInfo(FuncInfo& funcInfo, v8::Local<v8::Object>& output);
}