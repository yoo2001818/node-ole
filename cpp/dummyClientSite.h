#pragma once
#include "stdafx.h"

namespace node_ole {
	class DummyClientSite : public IOleClientSite {
	public:
		DummyClientSite();
		virtual ~DummyClientSite();

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
	};
}