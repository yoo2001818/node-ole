#pragma once
#include "stdafx.h"

namespace node_ole {
	enum ptr_type {
		pointer,
		carray,
		safearray
	};

	typedef struct ptr_info {
		ptr_type type;
		std::vector<ULONG> bounds;
	} ptr_info;

	typedef struct type_info {
		VARTYPE type;
		std::vector<ptr_info> ptrs;
	} type_info;

	typedef struct arg_info {
		BSTR name;
		USHORT flags;
		type_info type;
	} arg_info;

	typedef struct func_info {
		MEMBERID memId;
		DISPID dispId;
		BSTR name;
		BSTR description;
		type_info returnType;
		std::vector<arg_info> args;
		std::vector<arg_info> outs;
	} func_info;

	typedef struct dispatch_info {
		IUnknown * ptr;
		std::map<BSTR, std::vector<func_info>> info;
	} dispatch_info;
}