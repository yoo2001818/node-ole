#include "stdafx.h"
#include "info.h"
#include "eventListener.h"
#include "comutil.h"

namespace node_ole {
	EventListener::EventListener(Environment * env, LPTYPEINFO typeInfo, REFIID iid) {
		this->env = env;
		this->typeInfo = typeInfo;
		typeInfo->AddRef();
		
		this->iid = iid;
		this->registerTypeInfo(typeInfo);
	}
	EventListener::~EventListener() {
		this->typeInfo->Release();
	}
	HRESULT EventListener::registerTypeInfo(LPTYPEINFO typeInfo) {
		HRESULT result;
		LPTYPEATTR typeAttr;
		result = typeInfo->GetTypeAttr(&typeAttr);
		if FAILED(result) return result;

		// Read type flags; ignore if non-dispatchable, or hidden or restricted.
		WORD typeFlags = typeAttr->wTypeFlags;
		if (!(typeFlags & TYPEFLAG_FDISPATCHABLE)) return S_OK;

		// Read each defined functions.
		LPFUNCDESC funcDesc;
		for (UINT funcId = 0; funcId < typeAttr->cFuncs; ++funcId) {
			result = typeInfo->GetFuncDesc(funcId, &funcDesc);
			if FAILED(result) continue;

			FuncInfo funcInfo;
			result = readFuncInfo(typeInfo, funcDesc, &funcInfo, true);
			typeInfo->ReleaseFuncDesc(funcDesc);
			if FAILED(result) continue;

			this->funcInfoMap->insert_or_assign(funcDesc->memid, funcDesc);
		}

		// Read implemented types.
		for (UINT implId = 0; implId < typeAttr->cImplTypes; ++implId) {
			HREFTYPE handle;
			LPTYPEINFO implTypeInfo;
			result = typeInfo->GetRefTypeOfImplType(implId, &handle);
			if FAILED(result) return result;
			result = typeInfo->GetRefTypeInfo(handle, &implTypeInfo);
			if FAILED(result) return result;
			registerTypeInfo(implTypeInfo);
			implTypeInfo->Release();
		}

		return S_OK;
	}
	STDMETHODIMP_(ULONG)
	EventListener::AddRef(void) {
		return ++cRef;
	}
	STDMETHODIMP_(ULONG)
	EventListener::Release(void) {
		if (--cRef == 0) {
			delete this;
			return 0;
		}
		return cRef;
	}
	STDMETHODIMP
	EventListener::QueryInterface(REFIID riid, void ** ppvObject) {
		if (ppvObject == NULL) return E_INVALIDARG;

		*ppvObject = NULL;
		if (riid == iid || riid == IID_IUnknown || riid == IID_IDispatch) {
			*ppvObject = this;
		} else {
			return E_NOINTERFACE;
		}
		AddRef();
		return S_OK;
	}
	STDMETHODIMP
	EventListener::GetIDsOfNames(REFIID riid, LPOLESTR * rgszNames,
		UINT cNames, LCID lcid, DISPID * rgDispId
	) {
		return DispGetIDsOfNames(typeInfo, rgszNames, cNames, rgDispId);
	}
	STDMETHODIMP
	EventListener::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo ** ppTInfo) {
		if (iTInfo != 0) return DISP_E_BADINDEX;
		*ppTInfo = typeInfo;
		typeInfo->AddRef();
		return S_OK;
	}
	STDMETHODIMP
	EventListener::GetTypeInfoCount(UINT *pctinfo) {
		*pctinfo = 0;
		return S_OK;
	}
	STDMETHODIMP
	EventListener::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid,
		WORD wFlags, DISPPARAMS * pDispParams, VARIANT * pVarResult,
		EXCEPINFO * pExcepInfo, UINT * puArgErr
	) {
		printf("Invoked!!!! %d\n", dispIdMember);
		return S_OK;
	}
}