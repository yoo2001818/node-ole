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
				if (flags & PARAMFLAG_FIN) {
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
	void getArgType(v8::Local<v8::Value>& value, JSTypeInfo& output) {
		// Determine type of argument. Let's ignore BooleanObject, etc.
		// Actual conversion is done later.
		if (value->IsNullOrUndefined()) {
			output.type = VT_NULL;
		}
		else if (value->IsArray()) {
			auto arr = v8::Handle<v8::Array>::Cast(value);
			output.dimensions += 1;
			getArgType(arr->Get(0), output);
		}
		else if (value->IsBoolean()) {
			output.type = VT_BOOL;
		}
		else if (value->IsDate()) {
			output.type = VT_DATE;
		}
		else if (value->IsNumber()) {
			output.type = VT_R8;
		}
		else if (value->IsString()) {
			output.type = VT_BSTR;
		}
		else {
			Nan::ThrowTypeError("Unsupported type passed to node-ole.");
		}
	}
	std::vector<JSTypeInfo> getArgsType(Nan::NAN_METHOD_ARGS_TYPE& funcArgs) {
		// Convert the arguments to type information for pattern matching.
		std::vector<JSTypeInfo> output;
		for (int i = 0; i < funcArgs.Length(); ++i) {
			JSTypeInfo type;
			getArgType(funcArgs[i], type);
			output.push_back(type);
		}
		return output;
	}
	bool isTypeCompatiable(JSTypeInfo& jsInfo, TypeInfo& type) {
		// Handle arrays first - except for VARIANT and SAFEARRAY, we must check
		// if array dimension matches
		if (type.type == VT_VARIANT) return true;
		if (type.type == VT_SAFEARRAY) return jsInfo.dimensions > 0;
		auto remainingDimensions = jsInfo.dimensions;
		for (auto i = type.ptrs.begin(); i != type.ptrs.end(); ++i) {
			if (i->type != PtrType::CArray) continue;
			remainingDimensions -= i->bounds.size();
			if (remainingDimensions < 0) return false;
		}
		if (remainingDimensions != 0) return false;
		switch (type.type) {
			case VT_EMPTY:
			case VT_NULL:
			case VT_VOID:
				return jsInfo.type == VT_NULL;
			case VT_DATE:
				return jsInfo.type == VT_DATE;
			case VT_BSTR:
			case VT_LPSTR:
			case VT_LPWSTR:
				return jsInfo.type == VT_BSTR;
			// case VT_UNKNOWN:
			// case VT_DISPATCH: return "object";
			// case VT_ERROR: return "error";
			case VT_BOOL:
				return jsInfo.type == VT_BOOL;
			case VT_DECIMAL:
			case VT_CY:
			case VT_R4:
			case VT_R8:
			case VT_I2:
			case VT_I4:
			case VT_I1:
			case VT_UI1:
			case VT_UI2:
			case VT_UI4:
			case VT_I8:
			case VT_UI8:
			case VT_INT:
			case VT_UINT:
				return jsInfo.type == VT_R8;
			default:
				return false;
		}
	}
	bool isFuncCompatiable(std::vector<JSTypeInfo>& args, FuncInfo& funcInfo) {
		auto i = args.begin();
		for (auto j = funcInfo.args.begin(); j != funcInfo.args.end(); j++) {
			// Skip if current flag is not 'in'
			auto flags = j->flags;
			if (!(flags & PARAMFLAG_FIN)) continue;
			// Handle if args has ended
			if (i == args.end()) {
				if (flags & PARAMFLAG_FOPT) continue;
				else return false;
			}
			bool passed =
				// Allow 'null' if the type is optional.
				((flags & PARAMFLAG_FOPT) && i->type == VT_NULL) ||
				// Run actual check
				isTypeCompatiable(*i, j->type);
			if (passed) {
				if (i != args.end()) i++;
				continue;
			}
			else {
				return false;
			}
		}
		return true;
	}
	void writeDispParam(v8::Local<v8::Value>& value, VARTYPE type,
		VARIANT * output
	) {
		// We shouldn't create other VARIANT unless VT_BYREF is specified. (right?)
		// For the time being, we'll only support non-array objects.
		output->vt = type;
		switch (type & (~VT_BYREF)) {
		case VT_EMPTY:
		case VT_NULL:
			output->vt = VT_NULL;
			break;
		case VT_CY: {
			// Although COM supports currency format, we don't - we just convert
			// it to 64bit integer.
			if (type & VT_BYREF) output->pcyVal = new CY{ value->IntegerValue };
			else output->cyVal = value->IntegerValue;
			break;
		}
		case VT_DATE: {
			// What the heck??
			// DATE is a double, and a day is represented with number '1', and
			// it starts at 1899-12-30.
			// This is a complete mess.... however, JS Date uses Unix
			// timestamp * 1000.
			// Furthermore, DATE is not UTC - it is tied to local timezone!!
			// TODO Implement this mess
			double timestamp = 0;
			if (type & VT_BYREF) output->pdate = new DATE{ timestamp };
			else output->date = timestamp;
			break;
		}
		case VT_BSTR: {
			// Construct BSTR object from String
			v8::String::Value str(value);
			BSTR names = SysAllocString((LPOLESTR)* str);
			if (type & VT_BYREF) output->pbstrVal = new BSTR{ names };
			else output->bstrVal = names;
			break;
		}
					  // case VT_UNKNOWN:
					  // case VT_DISPATCH: return "object";
		case VT_BOOL: {
			VARIANT_BOOL boolVal = value->IsTrue() ? VARIANT_TRUE : VARIANT_FALSE;
			if (type & VT_BYREF) output->pboolVal = new VARIANT_BOOL{ boolVal };
			else output->boolVal = boolVal;
			break;
		}
		case VT_VARIANT: {
			// Determine the type from JS side
			if (value->IsNullOrUndefined()) {
				writeDispParam(value, VT_NULL, output);
			}
			else if (value->IsArray()) {
				// TODO
			}
			else if (value->IsBoolean()) {
				writeDispParam(value, VT_BOOL, output);
			}
			else if (value->IsDate()) {
				writeDispParam(value, VT_DATE, output);
			}
			else if (value->IsInt32()) {
				writeDispParam(value, VT_I4, output);
			}
			else if (value->IsUint32()) {
				writeDispParam(value, VT_UI4, output);
			}
			else if (value->IsNumber()) {
				writeDispParam(value, VT_R8, output);
			}
			else if (value->IsString()) {
				writeDispParam(value, VT_BSTR, output);
			}
			else {
				// Dunno
			}
		}
		case VT_DECIMAL: {
			if (type & VT_BYREF) output->pdecVal = new DECIMAL(value->NumberValue);
			else output->decVal = DECIMAL(value->NumberValue);
			break;
		}
		case VT_R4: {
			if (type & VT_BYREF) output->pfltVal = new float{ (float)value->NumberValue };
			else output->fltVal = (float)value->NumberValue;
			break;
		}
		case VT_R8: {
			if (type & VT_BYREF)output->pdblVal = new double{ value->NumberValue };
			else output->dblVal = value->NumberValue;
			break;
		}
		case VT_I1: {
			if (type & VT_BYREF)output->pcVal = new char{ value->Int32Value };
			else output->cVal = value->Int32Value;
			break;
		}
		case VT_I2: {
			if (type & VT_BYREF)output->piVal = new short{ value->Int32Value };
			else output->iVal = value->Int32Value;
			break;
		}
		case VT_I4: {
			if (type & VT_BYREF)output->plVal = new long{ value->Int32Value };
			else output->lVal = value->Int32Value;
			break;
		}
		case VT_I8: {
			if (type & VT_BYREF)output->pllVal = new long long{ value->IntegerValue };
			else output->llVal = value->IntegerValue;
			break;
		}
		case VT_UI1: {
			if (type & VT_BYREF)output->pbVal = new unsigned char{ value->Uint32Value };
			else output->bVal = value->Int32Value;
			break;
		}
		case VT_UI2: {
			if (type & VT_BYREF)output->puiVal = new unsigned short{ value->Uint32Value };
			else output->uiVal = value->Int32Value;
			break;
		}
		case VT_UI4: {
			if (type & VT_BYREF)output->pulVal = new unsigned long{ value->Uint32Value };
			else output->ulVal = value->IntegerValue;
			break;
		}
		case VT_UI8: {
			if (type & VT_BYREF)output->pullVal = new unsigned long long{ value->IntegerValue };
			else output->ullVal = value->IntegerValue;
			break;
		}
		case VT_INT: {
			if (type & VT_BYREF)output->pintVal = new INT{ value->IntegerValue };
			else output->intVal = value->IntegerValue;
			break;
		}
		case VT_UINT: {
			if (type & VT_BYREF)output->puintVal = new UINT{ value->IntegerValue };
			else output->uintVal = value->IntegerValue;
			break;
		}
		case VT_SAFEARRAY: {
			// Do nothing for now....
		}
		}
	}
	void constructDispParam(v8::Local<v8::Value>& value, TypeInfo& type,
		VARIANTARG * output
	) {
		// The internal format supports many formats, however,
		// VARIANT only supports one level of pointer indirection -
		// we just check for presence of pointer in types.
		// (and set VT_BYREF)
		// For the array, it isn't possible to use arrays again -
		// so we just have to be sure if there is a pointer after the array.
		// (and set VT_ARRAY)
		bool hasPointer = false;
		bool hasArray = false;
		bool hasArrayPointer = false;
		
		for (auto i = type.ptrs.begin(); i != type.ptrs.end(); i++) {
			switch (i->type) {
				case PtrType::CArray:
					hasArray = true;
					break;
				case PtrType::Pointer:
					if (hasArray) hasArrayPointer = true;
					else hasPointer = true;
					break;
			}
		}
		if (type.type == VT_SAFEARRAY) hasArray = true;
		// We shouldn't create other VARIANT unless VT_BYREF is specified. (right?)
		// For the time being, we'll only support non-array objects.
		writeDispParam(value, type.type & (hasPointer ? VT_BYREF : 0), output);
	}
	void constructDispParams(Nan::NAN_METHOD_ARGS_TYPE& args, FuncInfo& funcInfo,
		DISPPARAMS& output
	) {

	}
}