#include "stdafx.h"
#include "dummyClientSite.h"

namespace node_ole {

	DummyClientSite::DummyClientSite(HWND hwnd, RECT rect) {
		this->hwnd = hwnd;
		this->rect = rect;
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
		if (riid == IID_IUnknown || riid == IID_IOleClientSite || riid == IID_IOleInPlaceSite ||
			riid == IID_IOleContainer || riid == IID_IParseDisplayName ||
			riid == IID_IOleInPlaceFrame || riid == IID_IOleInPlaceUIWindow ||
			riid == IID_IOleWindow) {
			*ppvObject = this;
		} else {
			return E_NOINTERFACE;
		}
		AddRef();
		return S_OK;
	}
	STDMETHODIMP
	DummyClientSite::GetContainer(LPOLECONTAINER * ppContainer) {
		*ppContainer = this;
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
		return S_OK;
	}
	STDMETHODIMP
	DummyClientSite::SaveObject() {
		return S_OK;
	}
	STDMETHODIMP
	DummyClientSite::ShowObject() {
		return S_OK;
	}
	STDMETHODIMP
	DummyClientSite::ParseDisplayName(IBindCtx * pbc, LPOLESTR pszDisplayname,
		ULONG * pchEaten, IMoniker ** ppmkOut
	) {
		*ppmkOut = NULL;
		return MK_E_NOOBJECT;
	}
	STDMETHODIMP
	DummyClientSite::EnumObjects(DWORD grfFlags, IEnumUnknown ** ppenum) {
		*ppenum = NULL;
		return E_NOTIMPL;
	}
	STDMETHODIMP
	DummyClientSite::LockContainer(BOOL fLock) {
		return S_OK;
	}
	STDMETHODIMP
	DummyClientSite::CanInPlaceActivate() {
		return S_OK;
	}
	STDMETHODIMP
	DummyClientSite::DeactivateAndUndo() {
		return S_OK;
	}
	STDMETHODIMP
	DummyClientSite::DiscardUndoState() {
		return S_OK;
	}
	STDMETHODIMP
	DummyClientSite::GetWindowContext(IOleInPlaceFrame **ppFrame, IOleInPlaceUIWindow **ppDoc,
		LPRECT lprcPosRect, LPRECT lprcClipRect, LPOLEINPLACEFRAMEINFO lpFrameInfo
	) {
		if (ppFrame != NULL) *ppFrame = this;
		if (ppDoc != NULL) *ppDoc = NULL;
		if (lprcPosRect != NULL) CopyRect(&(this->rect), lprcPosRect);
		if (lpFrameInfo != NULL) {
			lpFrameInfo->fMDIApp = FALSE;
			lpFrameInfo->hwndFrame = hwnd;
			lpFrameInfo->cAccelEntries = 0;
		}
		return S_OK;
	}
	STDMETHODIMP
	DummyClientSite::OnInPlaceActivate() {
		return S_OK;
	}
	STDMETHODIMP
	DummyClientSite::OnInPlaceDeactivate() {
		return S_OK;
	}
	STDMETHODIMP
	DummyClientSite::OnPosRectChange(LPCRECT lprcPosRect) {
		return S_OK;
	}
	STDMETHODIMP
	DummyClientSite::OnUIActivate() {
		return S_OK;
	}
	STDMETHODIMP
	DummyClientSite::OnUIDeactivate(BOOL fUndoable) {
		return S_OK;
	}
	STDMETHODIMP
	DummyClientSite::Scroll(SIZE scrollExtant) {
		return S_OK;
	}
	STDMETHODIMP
	DummyClientSite::EnableModeless(BOOL fEnable) {
		return S_OK;
	}
	STDMETHODIMP
	DummyClientSite::InsertMenus(HMENU hmenuShared, LPOLEMENUGROUPWIDTHS lpMenuWidths) {
		return S_OK;
	}
	STDMETHODIMP
	DummyClientSite::RemoveMenus(HMENU hmenuShared) {
		return S_OK;
	}
	STDMETHODIMP
	DummyClientSite::SetMenu(HMENU hmenuShared, HOLEMENU holemenu, HWND hwndActiveObject) {
		return S_OK;
	}
	STDMETHODIMP
	DummyClientSite::SetStatusText(LPCOLESTR pszStatusText) {
		return S_OK;
	}
	STDMETHODIMP
	DummyClientSite::TranslateAccelerator(LPMSG lpmsg, WORD wID) {
		return S_FALSE;
	}
	STDMETHODIMP
	DummyClientSite::GetBorder(LPRECT lprectBorder) {
		CopyRect(&(this->rect), lprectBorder);
		return S_OK;
	}
	STDMETHODIMP
	DummyClientSite::RequestBorderSpace(LPCBORDERWIDTHS pborderwidths) {
		return INPLACE_E_NOTOOLSPACE;
	}
	STDMETHODIMP
	DummyClientSite::SetActiveObject(IOleInPlaceActiveObject *pActiveObject, LPCOLESTR pszObjName) {
		return S_OK;
	}
	STDMETHODIMP
	DummyClientSite::SetBorderSpace(LPCBORDERWIDTHS pborderwidths) {
		return S_OK;
	}
	STDMETHODIMP
	DummyClientSite::ContextSensitiveHelp(BOOL fEnterMode) {
		return S_OK;
	}
	STDMETHODIMP
	DummyClientSite::GetWindow(HWND * phwnd) {
		*phwnd = this->hwnd;
		return S_OK;
	}
}