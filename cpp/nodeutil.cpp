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
		case VT_EMPTY:
		case VT_NULL: return "null";
		case VT_CY: return "currency";
		case VT_DATE: return "date";
		case VT_BSTR: return "string";
		case VT_UNKNOWN:
		case VT_DISPATCH: return "object";
		case VT_ERROR: return "error";
		case VT_BOOL: return "boolean";
		case VT_VARIANT: return "any";
		case VT_DECIMAL:
		case VT_R4:
		case VT_R8: return "number";
		case VT_I2:
		case VT_I4:
		case VT_I1:
		case VT_UI1:
		case VT_UI2:
		case VT_UI4:
		case VT_I8:
		case VT_UI8:
		case VT_INT:
		case VT_UINT: return "integer";
		case VT_VOID: return "void";
		case VT_HRESULT: return "error";
		case VT_PTR: return "pointer";
		case VT_SAFEARRAY: return "array";
		case VT_CARRAY: return "array";
		case VT_USERDEFINED: return "any";
		case VT_LPSTR: return "string";
		case VT_LPWSTR: return "string";
		case VT_RECORD: return "record";
		case VT_INT_PTR:
		case VT_UINT_PTR: return "integer";
		case VT_ARRAY: return "array";
		case VT_BYREF: return "byref";
		default: return "unregistered";
		}
	}
	char * getVarCType(VARTYPE varType) {
		switch (varType) {
		case VT_EMPTY: return "EMPTY";
		case VT_NULL: return "NULL";
		case VT_I2: return "I2";
		case VT_I4: return "I4";
		case VT_R4: return "R4";
		case VT_R8: return "R8";
		case VT_CY: return "CURRENCY";
		case VT_DATE: return "DATE";
		case VT_BSTR: return "BSTR";
		case VT_DISPATCH: return "DISPATCH";
		case VT_ERROR: return "ERROR";
		case VT_BOOL: return "BOOL";
		case VT_VARIANT: return "VARIANT";
		case VT_UNKNOWN: return "UNKNOWN";
		case VT_DECIMAL: return "DECIMAL";
		case VT_I1: return "I1";
		case VT_UI1: return "UI1";
		case VT_UI2: return "UI2";
		case VT_UI4: return "UI4";
		case VT_I8: return "I8";
		case VT_UI8: return "UI8";
		case VT_INT: return "INT";
		case VT_UINT: return "UINT";
		case VT_VOID: return "VOID";
		case VT_HRESULT: return "HRESULT";
		case VT_PTR: return "PTR";
		case VT_SAFEARRAY: return "SAFEARRAY";
		case VT_CARRAY: return "CARRAY";
		case VT_USERDEFINED: return "USERDEFINED";
		case VT_LPSTR: return "LPSTR";
		case VT_LPWSTR: return "LPWSTR";
		case VT_RECORD: return "RECORD";
		case VT_INT_PTR: return "INT_PTR";
		case VT_UINT_PTR: return "UINT_PTR";
		/*
		case VT_FILETIME: return "empty";
		case VT_BLOB: return "empty";
		case VT_STREAM: return "empty";
		case VT_STORAGE: return "empty";
		case VT_STREAMED_OBJECT: return "empty";
		case VT_STORED_OBJECT: return "empty";
		case VT_BLOB_OBJECT: return "empty";
		*/
		case VT_CF: return "CF";
		case VT_CLSID: return "CLSID";
		case VT_VERSIONED_STREAM: return "VERSIONED_STREAM";
		case VT_VECTOR: return "VECTOR";
		case VT_ARRAY: return "ARRAY";
		case VT_BYREF: return "BYREF";
		default: return "unregistered";
		}
	}
	void constructTypeInfo(TypeInfo& typeInfo, v8::Local<v8::Object>& output) {
		auto arrayDimensions = Nan::New<v8::Array>();
		int arrayIndex = 0;
		bool hasArray = false;
		bool hasPointer = false;
		// Parse pointer information - since Node.js side doesn't have any pointers,
		// it'd be better to greatly simplify these.
		// We'll just present the user with 1) whether a pointer is available
		// 2) dimensions of an array, if any
		// 3) arrayType
		for (auto it = typeInfo.ptrs.begin(); it != typeInfo.ptrs.end(); ++it) {
			switch ((*it).type) {
			case PtrType::CArray: {
				hasArray = true;
				auto bounds = &((*it).bounds);
				for (auto j = bounds->begin(); j != bounds->end(); ++j) {
					arrayDimensions->Set(arrayIndex, Nan::New((uint32_t) *j));
					arrayIndex++;
				}
				break;
			}
			case PtrType::Pointer:
				hasPointer = true;
				break;
			}
		}
		output->Set(Nan::New("cType").ToLocalChecked(),
			Nan::New(getVarCType(typeInfo.type)).ToLocalChecked());
		output->Set(Nan::New("pointer").ToLocalChecked(), Nan::New(hasPointer));
		if (hasArray) {
			output->Set(Nan::New("type").ToLocalChecked(),
				Nan::New("array").ToLocalChecked());
			output->Set(Nan::New("arrayType").ToLocalChecked(),
				Nan::New(getVarType(typeInfo.type)).ToLocalChecked());
			output->Set(Nan::New("dimensions").ToLocalChecked(), arrayDimensions);
		} else {
			output->Set(Nan::New("type").ToLocalChecked(),
				Nan::New(getVarType(typeInfo.type)).ToLocalChecked());
		}
	}
	void constructArgInfo(ArgInfo& argInfo, v8::Local<v8::Object>& output) {
		constructTypeInfo(argInfo.type, output);
		output->Set(Nan::New("name").ToLocalChecked(),
			Nan::New((uint16_t *) argInfo.name.data()).ToLocalChecked());
		output->Set(Nan::New("in").ToLocalChecked(),
			Nan::New<v8::Boolean>(argInfo.flags & PARAMFLAG_FIN));
		output->Set(Nan::New("out").ToLocalChecked(),
			Nan::New<v8::Boolean>(argInfo.flags & PARAMFLAG_FOUT));
		output->Set(Nan::New("retVal").ToLocalChecked(),
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
		bool hasRetVal = false;
		{
			auto arr = Nan::New<v8::Array>();
			auto arrC = Nan::New<v8::Array>();
			int index = 0;
			int indexC = 0;
			for (auto it = funcInfo.args.begin(); it != funcInfo.args.end(); it++) {
				auto obj = Nan::New<v8::Object>();
				constructArgInfo(*it, obj);
				USHORT flags = (*it).flags;
				if (flags & PARAMFLAG_FRETVAL) {
					hasRetVal = true;
					output->Set(Nan::New("returns").ToLocalChecked(), obj);
				}
				arrC->Set(indexC, obj);
				indexC++;
				if (flags & PARAMFLAG_FIN || !(flags & PARAMFLAG_FOUT || flags & PARAMFLAG_FRETVAL)) {
					arr->Set(index, obj);
					index++;
				}
			}
			output->Set(Nan::New("args").ToLocalChecked(), arr);
			output->Set(Nan::New("cArgs").ToLocalChecked(), arrC);
		}
		{
			auto obj = Nan::New<v8::Object>();
			constructTypeInfo(funcInfo.returnType, obj);
			if (!hasRetVal) output->Set(Nan::New("returns").ToLocalChecked(), obj);
			output->Set(Nan::New("cReturns").ToLocalChecked(), obj);
		}
	}
}