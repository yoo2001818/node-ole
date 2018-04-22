#pragma once
#include "stdafx.h"

namespace node_ole {

	class DispatchInfo;

	enum class PtrType {
		Pointer,
		CArray
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
		std::shared_ptr<DispatchInfo> ref;
	};

	class ArgInfo {
	public:
		std::wstring name;
		USHORT flags;
		TypeInfo type;
	};

	class FuncInfo {
	public:
		DISPID dispId;
		SHORT vftId;
		INVOKEKIND invokeKind;
		std::wstring name;
		std::wstring description;
		TypeInfo returnType;
		std::vector<ArgInfo> args;
		std::vector<ArgInfo> outs;
	};

	class ConnectionPointConnection {
	public:
		LPCONNECTIONPOINT point;
		DWORD cookie;
		void * listener;
	};

	class DispatchInfo {
	public:
		IUnknown * ptr;
		IDispatch * dispPtr;
		std::list<std::wstring> typeNames;
		std::map<std::wstring, std::vector<FuncInfo>> info;
		std::map<std::wstring, std::vector<FuncInfo>> eventInfo;
		std::list<ConnectionPointConnection> connections;
	};
}