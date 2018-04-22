#include "stdafx.h"
#include "info.h"
#include "eventListener.h"

namespace node_ole {
	EventListener::EventListener(Environment * env, LPTYPEINFO typeInfo, REFIID iid) {
		this->env = env;
		this->typeInfo = typeInfo;
		typeInfo->AddRef();
		this->iid = iid;
	}
	EventListener::~EventListener() {
		this->typeInfo->Release();
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