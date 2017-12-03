#pragma once
#include "stdafx.h"
#include "info.h"

namespace node_ole {
	HRESULT initObject(const wchar_t * name, LPUNKNOWN * output);
	HRESULT getObjectInfo(LPUNKNOWN lpunk, DispatchInfo ** output);
	HRESULT readTypeInfo(LPTYPEINFO typeInfo, DispatchInfo * output);
	TypeInfo readElemDesc(LPELEMDESC elemDesc);
}