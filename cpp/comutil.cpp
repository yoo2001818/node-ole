#include "comutil.h"

namespace node_ole {
	HRESULT initObject(const wchar_t * name, LPUNKNOWN * output) {
		HRESULT result;
		CLSID clsId;
		// Retrieve CLSID
		result = CLSIDFromProgID(name, &clsId);
		if FAILED(result) {
			result = CLSIDFromString(name, &clsId);
			if FAILED(result) return CO_E_CLASSSTRING;
		}
		// Initialize OLE object
		LPUNKNOWN lpunk;
		result = GetActiveObject(clsId, NULL, &lpunk);
		if FAILED(result) {
			result = CoCreateInstance(clsId, NULL, CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER,
				IID_IUnknown, (LPVOID*) &lpunk);
			if FAILED(result) return result;
		}
		// All done, for now.
		*output = lpunk;
		return S_OK;
	}

	HRESULT getObjectInfo(LPUNKNOWN lpunk, DispatchInfo ** output) {
		HRESULT result;
		LPDISPATCH disp;
		result = lpunk->QueryInterface(&disp);
		if FAILED(result) return result;
		// Read type information and copy it into DispatchInfo.
		DispatchInfo * info = new DispatchInfo();
		info->ptr = lpunk;
		UINT numTypes;
		disp->GetTypeInfoCount(&numTypes);
		for (UINT i = 0; i < numTypes; ++i) {
			LPTYPEINFO typeInfo;
			disp->GetTypeInfo(i, NULL, &typeInfo);
			readTypeInfo(typeInfo, info);
			typeInfo->Release();
		}
		return S_OK;
	}

	HRESULT readTypeInfo(LPTYPEINFO typeInfo, DispatchInfo * output) {
		HRESULT result;
		LPTYPEATTR typeAttr;
		result = typeInfo->GetTypeAttr(&typeAttr);
		if FAILED(result) return result;

		// Read type flags; ignore if non-dispatchable, or hidden or restricted.
		WORD typeFlags = typeAttr->wTypeFlags;
		if (!(typeFlags & TYPEFLAG_FDISPATCHABLE)) return;
		if (typeFlags & TYPEFLAG_FRESTRICTED) return;
		if (typeFlags & TYPEFLAG_FHIDDEN) return;

		// Read type name.
		BSTR typeName;
		result = typeInfo->GetDocumentation(MEMBERID_NIL, &typeName, NULL, NULL, NULL);
		if FAILED(result) return result;
		output->typeNames.push_back((std::wstring) _bstr_t(typeName, false));

		// Read each defined functions.
		LPFUNCDESC funcDesc;
		for (UINT funcId = 0; funcId < typeAttr->cFuncs; ++funcId) {
			result = typeInfo->GetFuncDesc(funcId, &funcDesc);
			if FAILED(result) continue;

			// Stop if the function is hidden or restricted.
			WORD funcFlags = funcDesc->wFuncFlags;
			if (funcFlags & FUNCFLAG_FHIDDEN) continue;
			if (funcFlags & FUNCFLAG_FRESTRICTED) continue;

			FuncInfo funcInfo;
			// Read documentation.
			BSTR funcName;
			BSTR funcDoc;
			typeInfo->GetDocumentation(funcDesc->memid, &funcName, &funcDoc, NULL, NULL);
			funcInfo.name = (std::wstring) _bstr_t(funcName, false);
			funcInfo.description = (std::wstring) _bstr_t(funcDoc, false);
			funcInfo.invokeKind = funcDesc->invkind;

			// Read return type.
			funcInfo.returnType = readElemDesc(&(funcDesc->elemdescFunc));

			// Get names.
			// Before calling GetNames, allocate BSTR array first.
			UINT paramSize = funcDesc->cParams + 1;
			BSTR * namesArr = (BSTR *)malloc(sizeof(BSTR) * paramSize);
			typeInfo->GetNames(funcDesc->memid, namesArr, paramSize, &paramSize);

			// Process each arguments.
			for (UINT k = 1; k < paramSize; ++k) {
				ArgInfo argInfo;
				LPELEMDESC elemDesc = &(funcDesc->lprgelemdescParam[k - 1]);
				argInfo.flags = elemDesc->paramdesc.wParamFlags;
				argInfo.name = (std::wstring) _bstr_t(namesArr[k], false);
				argInfo.type = readElemDesc(elemDesc);
				funcInfo.args.push_back(argInfo);
			}

			// Free names
			free(namesArr);
			typeInfo->ReleaseFuncDesc(funcDesc);

			// TODO Put into info, etc.
		}

		// Read implemented types.
		for (UINT implId = 0; implId < typeAttr->cImplTypes; ++implId) {
			HREFTYPE handle;
			LPTYPEINFO implTypeInfo;
			result = typeInfo->GetRefTypeOfImplType(implId, &handle);
			if FAILED(result) return result;
			result = typeInfo->GetRefTypeInfo(handle, &implTypeInfo);
			if FAILED(result) return result;
			readTypeInfo(implTypeInfo, output);
			implTypeInfo->Release();
		}

		typeInfo->ReleaseTypeAttr(typeAttr);

		return S_OK;
	}

	TypeInfo readElemDesc(LPELEMDESC elemDesc) {
		// TODO
		return TypeInfo();
	}
}