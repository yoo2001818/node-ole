#pragma once
#include "stdafx.h"
#include "info.h"
#include "environment.h"

namespace node_ole {
	HRESULT initObject(const wchar_t * name, LPUNKNOWN * output);
	HRESULT getObjectInfo(Environment * env, LPUNKNOWN lpunk, DispatchInfo ** output, bool registerEvents);
	HRESULT readTypeInfo(LPTYPEINFO typeInfo, DispatchInfo * output, bool isEvent);
	HRESULT readFuncInfo(LPTYPEINFO typeInfo, LPFUNCDESC funcDesc, FuncInfo * output);
	HRESULT readVarInfo(LPTYPEINFO typeInfo, LPVARDESC varDesc, FuncInfo * output, bool isWrite);
	TypeInfo readElemDesc(LPELEMDESC elemDesc);
	void parseVariantTypeInfo(Environment * env, LPVARIANT input);
	void copyVariant(LPVARIANT input, LPVARIANT output);
	HRESULT createEventObject(Environment * env, LPTYPEINFO typeInfo, DispatchInfo * dispInfo, LPUNKNOWN * output);
}