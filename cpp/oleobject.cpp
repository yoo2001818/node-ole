#include "oleobject.h"
#include "nodeutil.h"

namespace node_ole {
	OLEObject::OLEObject(Environment * env, DispatchInfo * dispInfo) : env(env), dispInfo(dispInfo) {
		env->ref();
	}
	OLEObject::~OLEObject() {
		env->unref();
	}
	NAN_METHOD(OLEObject::Invoke) {

	}
	void OLEObject::bake(v8::Local<v8::Object> object) {
		Wrap(object);
		// Read out dispatch info, and put it to the object
		for (auto iter = dispInfo->info.begin(); iter != dispInfo->info.end(); iter++) {
			auto tpl = Nan::New<v8::FunctionTemplate>(Invoke, Nan::New<v8::External>(&(*iter)));
			auto func = Nan::GetFunction(tpl).ToLocalChecked();
			// Set the data of func
			auto typeArr = Nan::New<v8::Array>();
			auto& list = iter->second;
			int count = 0;
			for (auto iter = list.begin(); iter != list.end(); iter++) {
				// Set type information
				auto obj = Nan::New<v8::Object>();
				constructFuncInfo(*iter, obj);
				typeArr->Set(count++, obj);
			}
			func->Set(Nan::New("types").ToLocalChecked(), typeArr);
			object->Set(Nan::New((uint16_t *)iter->first.data()).ToLocalChecked(), func);
		}
	}
}