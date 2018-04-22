#pragma once
#include "stdafx.h"
#include "info.h"
#include "environment.h"

namespace node_ole {
	class EventListener : public IDispatch {
	public:
		EventListener(Environment * env, DispatchInfo * info, LPTYPEINFO typeInfo, REFIID iid);
		virtual ~EventListener();

		HRESULT registerTypeInfo(LPTYPEINFO typeInfo);

		Environment * env;
		DispatchInfo * info;
		IID iid;
		LPTYPEINFO typeInfo;
		std::map<MEMBERID, FuncInfo> funcInfoMap;
		ULONG cRef = 0;

		STDMETHOD_(ULONG, AddRef)(void);
		STDMETHOD_(ULONG, Release)(void);
		STDMETHOD(QueryInterface)(REFIID riid, void ** ppvObject);

		STDMETHOD(GetIDsOfNames)(REFIID riid, LPOLESTR * rgszNames, UINT cNames,
			LCID lcid, DISPID * rgDispId);
		STDMETHOD(GetTypeInfo)(UINT iTInfo, LCID lcid, ITypeInfo ** ppTInfo);
		STDMETHOD(GetTypeInfoCount)(UINT *pctinfo);
		STDMETHOD(Invoke)(DISPID dispIdMember, REFIID riid, LCID lcid,
			WORD wFlags, DISPPARAMS * pDispParams, VARIANT * pVarResult,
			EXCEPINFO * pExcepInfo, UINT * puArgErr);
	};
}