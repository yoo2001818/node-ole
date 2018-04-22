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
		auto data = v8::Handle<v8::External>::Cast(info.Data());
		auto funcInfos = (std::vector<FuncInfo> *) data->Value();
		OLEObject * ole = Unwrap<OLEObject>(info.Holder());
		Environment * env = ole->env;
		DispatchInfo * dispInfo = ole->dispInfo;
		auto args = getArgsType(info);
		bool matched = false;
		// Search for the matching function entry.
		for (auto iter = funcInfos->begin(); iter != funcInfos->end(); iter++) {
			// Match all 'input' entries. Iterate through current entry and exit if failed.
			if (!isFuncCompatiable(args, *iter)) continue;
			matched = true;

			// Generate DISPPARAMS information using args and type info.
			DISPPARAMS * params = new DISPPARAMS();
			constructDispParams(info, *iter, params);

			// Create Promise object.
			v8::Local<v8::Promise::Resolver> resolver = v8::Promise::Resolver::New(info.GetIsolate());
			v8::Local<v8::Promise> promise = resolver->GetPromise();
			ResolverPersistent * deferred = new ResolverPersistent(resolver);

			// Create event object, then push it.
			std::unique_ptr<RequestInvoke> req = std::make_unique<RequestInvoke>();
			req->deferred = deferred;
			req->params = params;
			req->funcInfo = &(*iter);
			req->dispatch = dispInfo->dispPtr;

			env->pushRequest(std::move(req));

			info.GetReturnValue().Set(promise);
			break;
		}
		if (!matched) {
			Nan::ThrowTypeError("Cannot find corresponding function declaration.");
		}
	}
	void OLEObject::bake(v8::Local<v8::Object> object) {
		Wrap(object);
		auto typeNames = Nan::New<v8::Array>();
		int typeNameCount = 0;
		for (auto iter = dispInfo->typeNames.begin(); iter != dispInfo->typeNames.end(); iter++) {
			typeNames->Set(typeNameCount++, Nan::New((uint16_t *)iter->data()).ToLocalChecked());
		}
		object->Set(Nan::New("typeNames").ToLocalChecked(), typeNames);
		// Read out dispatch info, and put it to the object
		for (auto iter = dispInfo->info.begin(); iter != dispInfo->info.end(); iter++) {
			auto tpl = Nan::New<v8::FunctionTemplate>(Invoke, Nan::New<v8::External>(&(iter->second)));
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
		// Read event info too
		auto eventsInfoObj = Nan::New<v8::Object>();
		for (auto iter = dispInfo->eventInfo.begin(); iter != dispInfo->eventInfo.end(); iter++) {
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
			eventsInfoObj->Set(Nan::New((uint16_t *)iter->first.data()).ToLocalChecked(), typeArr);
		}
		object->Set(Nan::New("eventsInfo").ToLocalChecked(), eventsInfoObj);
	}
}