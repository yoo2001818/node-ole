#pragma once
#include "stdafx.h"
#include "info.h"
#include "environment.h"

namespace node_ole {
	HRESULT initObject(const wchar_t * name, LPUNKNOWN * output);
	HRESULT getObjectInfo(Environment * env, LPUNKNOWN lpunk, DispatchInfo ** output);
	HRESULT readTypeInfo(LPTYPEINFO typeInfo, DispatchInfo * output, bool isEvent);
	HRESULT readFuncInfo(LPTYPEINFO typeInfo, LPFUNCDESC funcDesc, FuncInfo * output, bool isEvent);
	TypeInfo readElemDesc(LPELEMDESC elemDesc);
	void parseVariantTypeInfo(Environment * env, LPVARIANT input);
	HRESULT createEventObject(Environment * env, LPTYPEINFO typeInfo, DispatchInfo * dispInfo, LPUNKNOWN * output);
}