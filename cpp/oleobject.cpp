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
		auto args = getArgsType(info);
		bool matched = false;
		// Search for the matching function entry.
		for (auto iter = funcInfos->begin(); iter != funcInfos->end(); iter++) {
			// Match all 'input' entries. Iterate through current entry and exit if failed.
			bool matchFailed = false;
			auto funcArgs = iter->args;
			{
				auto i = args.begin();
				auto j = funcArgs.begin();
				while (j != funcArgs.end()) {
					// Skip if current flag is not 'in'
					auto flags = j->flags;
					if (!(flags & PARAMFLAG_FIN)) {
						j++;
						continue;
					}
					// Handle if args has ended
					if (i == args.end()) {
						if (flags & PARAMFLAG_FOPT) {
							j++;
							continue;
						} else {
							// Failed
							matchFailed = true;
							break;
						}
					}
					bool passed =
						// Allow 'null' if the type is optional.
						((flags & PARAMFLAG_FOPT) && i->type == VT_NULL) ||
						// Run actual check
						isTypeCompatiable(*i, j->type);
					if (passed) {
						if (i != args.end()) i++;
						j++;
						continue;
					} else {
						matchFailed = true;
						break;
					}
				}
			}
			if (matchFailed) continue;
			matched = true;
			printf("%S\n", (wchar_t *)iter->name.data());
			info.GetReturnValue().Set(Nan::New(true));
			break;
		}
		if (!matched) {
			Nan::ThrowTypeError("Cannot find corresponding function declaration.");
		}
	}
	void OLEObject::bake(v8::Local<v8::Object> object) {
		Wrap(object);
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
	}
}