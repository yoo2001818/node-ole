#pragma once
#include "stdafx.h"

namespace node_ole {

	class DummyClientSite : public IOleInPlaceSite, public IOleClientSite, public IOleContainer,
		public IOleInPlaceFrame {
	public:
		DummyClientSite(HWND hwnd, RECT rect);
		virtual ~DummyClientSite();

		HWND hwnd;
		RECT rect;
		ULONG cRef = 0;

		STDMETHOD_(ULONG, AddRef)(void);
		STDMETHOD_(ULONG, Release)(void);
		STDMETHOD(QueryInterface)(REFIID riid, void ** ppvObject);

		STDMETHOD(GetContainer)(LPOLECONTAINER * ppContainer);
		STDMETHOD(GetMoniker)(DWORD dwAssign, DWORD dwWhichMoniker, LPMONIKER * ppmk);
		STDMETHOD(OnShowWindow)(BOOL fShow);
		STDMETHOD(RequestNewObjectLayout)();
		STDMETHOD(SaveObject)();
		STDMETHOD(ShowObject)();

		STDMETHOD(ParseDisplayName)(IBindCtx * pbc, LPOLESTR pszDisplayname,
			ULONG * pchEaten, IMoniker ** ppmkOut);

		STDMETHOD(EnumObjects)(DWORD grfFlags, IEnumUnknown ** ppenum);
		STDMETHOD(LockContainer)(BOOL fLock);

		STDMETHOD(CanInPlaceActivate)();
		STDMETHOD(DeactivateAndUndo)();
		STDMETHOD(DiscardUndoState)();
		STDMETHOD(GetWindowContext)(IOleInPlaceFrame **ppFrame, IOleInPlaceUIWindow **ppDoc,
			LPRECT lprcPosRect, LPRECT lprcClipRect, LPOLEINPLACEFRAMEINFO lpFrameInfo);
		STDMETHOD(OnInPlaceActivate)();
		STDMETHOD(OnInPlaceDeactivate)();
		STDMETHOD(OnPosRectChange)(LPCRECT lprcPosRect);
		STDMETHOD(OnUIActivate)();
		STDMETHOD(OnUIDeactivate)(BOOL fUndoable);
		STDMETHOD(Scroll)(SIZE scrollExtant);

		STDMETHOD(EnableModeless)(BOOL fEnable);
		STDMETHOD(InsertMenus)(HMENU hmenuShared, LPOLEMENUGROUPWIDTHS lpMenuWidths);
		STDMETHOD(RemoveMenus)(HMENU hmenuShared);
		STDMETHOD(SetMenu)(HMENU hmenuShared, HOLEMENU holemenu, HWND hwndActiveObject);
		STDMETHOD(SetStatusText)(LPCOLESTR pszStatusText);
		STDMETHOD(TranslateAccelerator)(LPMSG lpmsg, WORD wID);

		STDMETHOD(GetBorder)(LPRECT lprectBorder);
		STDMETHOD(RequestBorderSpace)(LPCBORDERWIDTHS pborderwidths);
		STDMETHOD(SetActiveObject)(IOleInPlaceActiveObject *pActiveObject, LPCOLESTR pszObjName);
		STDMETHOD(SetBorderSpace)(LPCBORDERWIDTHS pborderwidths);

		STDMETHOD(ContextSensitiveHelp)(BOOL fEnterMode);
		STDMETHOD(GetWindow)(HWND * phwnd);
	};
}