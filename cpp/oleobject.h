#pragma once
#include "stdafx.h"
#include "info.h"
#include "environment.h"

namespace node_ole {
	class OLEObject : public Nan::ObjectWrap {
		friend Environment;
	public:
		OLEObject(Environment * env, DispatchInfo * dispInfo);
		virtual ~OLEObject();

		static NAN_METHOD(Invoke);

		void bake(v8::Local<v8::Object> object);

		Environment * env;
		DispatchInfo * dispInfo;
	};
}