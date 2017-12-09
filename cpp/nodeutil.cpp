#include "nodeutil.h"

namespace node_ole {
	char * getInvokeKind(INVOKEKIND invKind) {
		switch (invKind) {
		case INVOKE_FUNC: return "function";
		case INVOKE_PROPERTYGET: return "get";
		case INVOKE_PROPERTYPUT: return "put";
		case INVOKE_PROPERTYPUTREF: return "putRef";
		default: return "unknown";
		}
	}
	char * getVarType(VARTYPE varType) {
		switch (varType) {
		case VT_EMPTY: return "empty";
		case VT_NULL: return "null";
		case VT_I2: return "i2";
		case VT_I4: return "i4";
		case VT_R4: return "r4";
		case VT_R8: return "r8";
		case VT_CY: return "CURRENCY";
		case VT_DATE: return "DATE";
		case VT_BSTR: return "BSTR";
		case VT_DISPATCH: return "IDispatch";
		case VT_ERROR: return "SCODE";
		case VT_BOOL: return "bool";
		case VT_VARIANT: return "VARIANT";
		case VT_UNKNOWN: return "IUnknown";
		case VT_DECIMAL: return "decimal";
		case VT_I1: return "i1";
		case VT_UI1: return "ui1";
		case VT_UI2: return "ui2";
		case VT_UI4: return "ui4";
		case VT_I8: return "i8";
		case VT_UI8: return "ui8";
		case VT_INT: return "int";
		case VT_UINT: return "uint";
		case VT_VOID: return "void";
		case VT_HRESULT: return "HRESULT";
		case VT_PTR: return "pointer";
		case VT_SAFEARRAY: return "safearray";
		case VT_CARRAY: return "carray";
		case VT_USERDEFINED: return "user defined";
		case VT_LPSTR: return "char *";
		case VT_LPWSTR: return "w_str *";
		case VT_RECORD: return "record";
		case VT_INT_PTR: return "int *";
		case VT_UINT_PTR: return "uint *";
		/*
		case VT_FILETIME: return "empty";
		case VT_BLOB: return "empty";
		case VT_STREAM: return "empty";
		case VT_STORAGE: return "empty";
		case VT_STREAMED_OBJECT: return "empty";
		case VT_STORED_OBJECT: return "empty";
		case VT_BLOB_OBJECT: return "empty";
		*/
		case VT_CF: return "empty";
		case VT_CLSID: return "empty";
		case VT_VERSIONED_STREAM: return "empty";
		case VT_VECTOR: return "empty";
		case VT_ARRAY: return "empty";
		case VT_BYREF: return "empty";
		default: return "unknown";
		}
	}
	void constructTypeInfo(TypeInfo& typeInfo, v8::Local<v8::Object>& output) {
		output->Set(Nan::New("type").ToLocalChecked(),
			Nan::New(getVarType(typeInfo.type)).ToLocalChecked());
	}
	void constructArgInfo(ArgInfo& argInfo, v8::Local<v8::Object>& output) {
		constructTypeInfo(argInfo.type, output);
		output->Set(Nan::New("name").ToLocalChecked(),
			Nan::New((uint16_t *) argInfo.name.data()).ToLocalChecked());
		output->Set(Nan::New("in").ToLocalChecked(),
			Nan::New<v8::Boolean>(argInfo.flags & PARAMFLAG_FIN));
		output->Set(Nan::New("out").ToLocalChecked(),
			Nan::New<v8::Boolean>(argInfo.flags & PARAMFLAG_FOUT));
		output->Set(Nan::New("retval").ToLocalChecked(),
			Nan::New<v8::Boolean>(argInfo.flags & PARAMFLAG_FRETVAL));
		output->Set(Nan::New("optional").ToLocalChecked(),
			Nan::New<v8::Boolean>(argInfo.flags & PARAMFLAG_FOPT));
		output->Set(Nan::New("hasDefault").ToLocalChecked(),
			Nan::New<v8::Boolean>(argInfo.flags & PARAMFLAG_FHASDEFAULT));
	}
	void constructFuncInfo(FuncInfo& funcInfo, v8::Local<v8::Object>& output) {
		output->Set(Nan::New("dispId").ToLocalChecked(), Nan::New(funcInfo.dispId));
		output->Set(Nan::New("name").ToLocalChecked(),
			Nan::New((uint16_t *) funcInfo.name.data()).ToLocalChecked());
		output->Set(Nan::New("description").ToLocalChecked(),
			Nan::New((uint16_t *) funcInfo.description.data()).ToLocalChecked());
		output->Set(Nan::New("type").ToLocalChecked(), 
			Nan::New(getInvokeKind(funcInfo.invokeKind)).ToLocalChecked());
		{
			auto obj = Nan::New<v8::Object>();
			constructTypeInfo(funcInfo.returnType, obj);
			output->Set(Nan::New("returns").ToLocalChecked(), obj);
		}
		{
			auto arr = Nan::New<v8::Array>();
			int index = 0;
			for (auto it = funcInfo.args.begin(); it != funcInfo.args.end(); it++) {
				auto obj = Nan::New<v8::Object>();
				constructArgInfo(*it, obj);
				arr->Set(index, obj);
				index++;
			}
			output->Set(Nan::New("args").ToLocalChecked(), arr);
		}
	}
}