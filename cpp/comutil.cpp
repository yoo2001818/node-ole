#include "comutil.h"

namespace node_ole {
	HRESULT initObject(const wchar_t * name, LPUNKNOWN * output) {
		HRESULT result;
		CLSID clsId;
		// Retrieve CLSID
		result = CLSIDFromString(name, &clsId);
		if FAILED(result) {
			result = CLSIDFromProgID(name, &clsId);
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
		std::vector<LPTYPELIB> typeLibs;
		info->ptr = lpunk;
		info->dispPtr = disp;
		UINT numTypes;
		disp->GetTypeInfoCount(&numTypes);
		for (UINT i = 0; i < numTypes; ++i) {
			LPTYPEINFO typeInfo;
			disp->GetTypeInfo(i, NULL, &typeInfo);
			LPTYPELIB typeLib;
			UINT typePos;
			typeInfo->GetContainingTypeLib(&typeLib, &typePos);
			typeLibs.push_back(typeLib);
			readTypeInfo(typeInfo, info, false);
			typeInfo->Release();
		}
		// Read event listener type info.
		LPCONNECTIONPOINTCONTAINER connPointContainer;
		result = lpunk->QueryInterface(&connPointContainer);
		if SUCCEEDED(result) {
			LPENUMCONNECTIONPOINTS connPoints;
			connPointContainer->EnumConnectionPoints(&connPoints);
			LPCONNECTIONPOINT connPoint;
			ULONG connFetched = 0;
			connPoints->Next(1, &connPoint, &connFetched);
			while (connFetched > 0) {
				IID connIID;
				connPoint->GetConnectionInterface(&connIID);
				// Fetch ITypeInfo fromTypeInfo above.
				LPTYPEINFO typeInfo = NULL;
				for (auto iter = typeLibs.begin(); iter != typeLibs.end(); iter++) {
					LPTYPELIB typeLib = *iter;
					typeLib->GetTypeInfoOfGuid(connIID, &typeInfo);
					if (typeInfo != NULL) break;
				}
				if (typeInfo != NULL) {
					printf("Found a type info\n");
					// Read typeInfo to an object.
					readTypeInfo(typeInfo, info, true);
					typeInfo->Release();
				}
				connPoint->Release();
				connPoints->Next(1, &connPoint, &connFetched);
			}
			connPoints->Release();
			connPointContainer->Release();
		}
		for (auto iter = typeLibs.front(); iter != typeLibs.back(); iter++) {
			iter->Release();
		}
		*output = info;
		return S_OK;
	}

	HRESULT readTypeInfo(LPTYPEINFO typeInfo, DispatchInfo * output, bool isEvent) {
		HRESULT result;
		LPTYPEATTR typeAttr;
		result = typeInfo->GetTypeAttr(&typeAttr);
		if FAILED(result) return result;

		// Read type flags; ignore if non-dispatchable, or hidden or restricted.
		WORD typeFlags = typeAttr->wTypeFlags;
		if (!(typeFlags & TYPEFLAG_FDISPATCHABLE)) return S_OK;
		// if (typeFlags & TYPEFLAG_FRESTRICTED) return S_OK;
		// if (typeFlags & TYPEFLAG_FHIDDEN) return S_OK;

		// Read type name.
		BSTR typeName;
		result = typeInfo->GetDocumentation(MEMBERID_NIL, &typeName, NULL, NULL, NULL);
		if FAILED(result) return result;
		output->typeNames.push_back(std::move((std::wstring) _bstr_t(typeName, false)));

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
			if (funcName != NULL) funcInfo.name = (std::wstring) _bstr_t(funcName, false);
			if (funcDoc != NULL) funcInfo.description = (std::wstring) _bstr_t(funcDoc, false);
			funcInfo.invokeKind = funcDesc->invkind;

			funcInfo.dispId = funcDesc->memid;

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

			auto infoMap = &(output->info);
			if (isEvent) infoMap = &(output->eventInfo);
			// Put the info into the map.
			auto found = infoMap->find(funcInfo.name);
			if (found != infoMap->end()) {
				found->second.push_back(std::move(funcInfo));
			} else {
				(*infoMap)[funcInfo.name] = { funcInfo };
			}
		}

		// Read implemented types.
		for (UINT implId = 0; implId < typeAttr->cImplTypes; ++implId) {
			HREFTYPE handle;
			LPTYPEINFO implTypeInfo;
			result = typeInfo->GetRefTypeOfImplType(implId, &handle);
			if FAILED(result) return result;
			result = typeInfo->GetRefTypeInfo(handle, &implTypeInfo);
			if FAILED(result) return result;
			readTypeInfo(implTypeInfo, output, isEvent);
			implTypeInfo->Release();
		}

		typeInfo->ReleaseTypeAttr(typeAttr);

		return S_OK;
	}

	TypeInfo readElemDesc(LPELEMDESC elemDesc) {
		TypeInfo info;
		// We read the arrays, and pointers, and put them in a serialized std::vector.
		// All pointers are recorded in pre-order, i.e. it gets added to vector at the end.
		TYPEDESC * typeDesc = &(elemDesc->tdesc);
		while (true) {
			switch (typeDesc->vt) {
			case VT_CARRAY: {
				// CArray
				PtrInfo ptrInfo;
				ptrInfo.type = PtrType::CArray;
				ARRAYDESC * arrayDesc = typeDesc->lpadesc;
				USHORT dims = arrayDesc->cDims;
				for (USHORT i = 0; i < dims; ++i) {
					// TODO Store cElements
					ptrInfo.bounds.push_back(arrayDesc->rgbounds[i].lLbound);
				}
				info.ptrs.push_back(std::move(ptrInfo));
				typeDesc = &(arrayDesc->tdescElem);
				continue;
			}
			case VT_ARRAY:
			case VT_SAFEARRAY:
				// SafeArray doesn't have fixed array types - just skip it.
				info.type = VT_SAFEARRAY;
				break;
			case VT_PTR:
				// Pointer
				info.ptrs.push_back(PtrInfo{ PtrType::Pointer });
				typeDesc = typeDesc->lptdesc;
				continue;
			case VT_USERDEFINED:
				// Why
			case VT_INT_PTR:
				info.type = VT_INT;
				info.ptrs.push_back(PtrInfo{ PtrType::Pointer });
				break;
			case VT_UINT_PTR:
				info.type = VT_UINT;
				info.ptrs.push_back(PtrInfo{ PtrType::Pointer });
				break;
			case VT_BYREF:
				// TODO I have no idea what this does
			default:
				info.type = typeDesc->vt;
				break;
			}
			return info;
		}
	}
	void parseVariantTypeInfo(LPVARIANT input) {
		DispatchInfo * info = nullptr;
		if (input->vt == (VT_DISPATCH | VT_BYREF) || input->vt == (VT_UNKNOWN | VT_BYREF)) {
			getObjectInfo(*(input->ppdispVal), &info);
			input->vt = VT_DISPATCH | VT_BYREF;
			input->pdispVal = (LPDISPATCH) info;
		}
		if (input->vt == VT_DISPATCH || input->vt == VT_UNKNOWN) {
			getObjectInfo(input->pdispVal, &info);
			input->vt = VT_DISPATCH | VT_BYREF;
			input->pdispVal = (LPDISPATCH) info;
		}
	}
}