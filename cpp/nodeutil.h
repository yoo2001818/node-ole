#pragma once
#include "stdafx.h"
#include "info.h"

namespace node_ole {
	void constructFuncInfo(FuncInfo& funcInfo, v8::Local<v8::Object>& output);

	class JSTypeInfo {
	public:
		VARTYPE type;
		uint32_t dimensions = 0;
	};
	
	void getArgType(v8::Local<v8::Value>& value, JSTypeInfo& output);
	std::vector<JSTypeInfo> getArgsType(Nan::NAN_METHOD_ARGS_TYPE& funcArgs);
	bool isTypeCompatiable(JSTypeInfo& jsInfo, TypeInfo& type);
	bool isFuncCompatiable(std::vector<JSTypeInfo>& args, FuncInfo& funcInfo);
	void writeVariant(v8::Local<v8::Value>& value, VARTYPE type,
		VARIANT * output);
	void constructVariant(v8::Local<v8::Value>& value, TypeInfo& type,
		VARIANTARG * output);
	void constructEmptyVariant(TypeInfo& type, VARIANTARG * output);
	void constructDispParams(Nan::NAN_METHOD_ARGS_TYPE& args, FuncInfo& funcInfo,
		DISPPARAMS * output);
	v8::Local<v8::Value> readVariant(VARIANT * input);
}