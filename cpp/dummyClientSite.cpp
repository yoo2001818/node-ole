#include "stdafx.h"
#include "dummyClientSite.h"

namespace node_ole {
	DummyClientSite::DummyClientSite() {
	}
	DummyClientSite::~DummyClientSite() {
	}
	STDMETHODIMP_(ULONG)
	DummyClientSite::AddRef(void) {
		return ++cRef;
	}
	STDMETHODIMP_(ULONG)
	DummyClientSite::Release(void) {
		if (--cRef == 0) {
			delete this;
			return 0;
		}
		return cRef;
	}
	STDMETHODIMP
	DummyClientSite::QueryInterface(REFIID riid, void ** ppvObject) {
		if (ppvObject == NULL) return E_INVALIDARG;

		*ppvObject = NULL;
		if (riid == IID_IUnknown || riid == IID_IOleClientSite) {
			*ppvObject = this;
		}
		else {
			return E_NOINTERFACE;
		}
		AddRef();
		return S_OK;
	}
	STDMETHODIMP
	DummyClientSite::GetContainer(LPOLECONTAINER * ppContainer) {
		*ppContainer = (LPOLECONTAINER) 1111;
		return S_OK;
	}
	STDMETHODIMP
	DummyClientSite::GetMoniker(DWORD dwAssign, DWORD dwWhichMoniker, LPMONIKER * ppmk) {
		*ppmk = NULL;
		return E_NOTIMPL;
	}
	STDMETHODIMP
	DummyClientSite::OnShowWindow(BOOL fShow) {
		return S_OK;
	}
	STDMETHODIMP
	DummyClientSite::RequestNewObjectLayout() {
		return E_NOTIMPL;
	}
	STDMETHODIMP
	DummyClientSite::SaveObject() {
		return S_OK;
	}
	STDMETHODIMP
	DummyClientSite::ShowObject() {
		return S_OK;
	}
}