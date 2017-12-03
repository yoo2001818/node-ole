#pragma once
#include "stdafx.h"

namespace node_ole {
	enum class PtrType {
		Pointer,
		CArray,
		SafeArray
	};

	class PtrInfo {
	public:
		PtrType type;
		std::vector<ULONG> bounds;
	};

	class TypeInfo {
	public:
		VARTYPE type;
		std::vector<PtrInfo> ptrs;
	};

	class ArgInfo {
	public:
		std::wstring name;
		USHORT flags;
		TypeInfo type;
	};

	class FuncInfo {
	public:
		MEMBERID memId;
		DISPID dispId;
		INVOKEKIND invokeKind;
		std::wstring name;
		std::wstring description;
		TypeInfo returnType;
		std::vector<ArgInfo> args;
		std::vector<ArgInfo> outs;
	};

	class DispatchInfo {
	public:
		IUnknown * ptr;
		std::list<std::wstring> typeNames;
		std::map<std::wstring, std::vector<FuncInfo>> info;
		std::map<std::wstring, std::vector<FuncInfo>> eventInfo;
	};
}