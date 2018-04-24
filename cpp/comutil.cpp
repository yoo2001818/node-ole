#include "comutil.h"
#include "eventListener.h"
#include "dummyClientSite.h"

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

	HRESULT getObjectInfo(Environment * env, LPUNKNOWN lpunk, DispatchInfo ** output, bool registerEvents) {
		HRESULT result;
		LPDISPATCH disp;
		if (lpunk == NULL) return E_FAIL;
		result = lpunk->QueryInterface(&disp);
		if FAILED(result) return result;
		if (disp == NULL) return E_FAIL;
		// Read type information and copy it into DispatchInfo.
		DispatchInfo * info = new DispatchInfo();
		std::vector<LPTYPELIB> typeLibs;
		info->ptr = lpunk;
		info->dispPtr = disp;
		UINT numTypes;
		disp->GetTypeInfoCount(&numTypes);
		for (UINT i = 0; i < numTypes; ++i) {
			LPTYPEINFO typeInfo = NULL;
			disp->GetTypeInfo(i, NULL, &typeInfo);
			if (typeInfo == NULL) continue;
			LPTYPELIB typeLib;
			UINT typePos;
			typeInfo->GetContainingTypeLib(&typeLib, &typePos);
			if (typeLib == NULL) continue;
			typeLibs.push_back(typeLib);
			readTypeInfo(typeInfo, info, false);
			typeInfo->Release();
		}
		// Read event listener type info.
		LPCONNECTIONPOINTCONTAINER connPointContainer;
		if (registerEvents) {
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
						// Read typeInfo to an object.
						readTypeInfo(typeInfo, info, true);
						LPUNKNOWN eventObj;
						result = createEventObject(env, typeInfo, info, &eventObj);
						if SUCCEEDED(result) {
							DWORD cookie;
							connPoint->Advise(eventObj, &cookie);
							// All good! Create connection point object for removing it later.
							ConnectionPointConnection connection;
							connection.cookie = cookie;
							connection.listener = (EventListener *)eventObj;
							eventObj->AddRef();
							connection.point = connPoint;
							connection.point->AddRef();
							info->connections.push_back(connection);
						}
						typeInfo->Release();
					}
					connPoint->Release();
					connPoints->Next(1, &connPoint, &connFetched);
				}
				connPoints->Release();
				connPointContainer->Release();
			}
		}
		// Initialize OLE Object, if it's necessary.
		if (registerEvents) {
			LPOLEINPLACEOBJECT lpOleInPlace;
			result = lpunk->QueryInterface(&lpOleInPlace);
			if SUCCEEDED(result) {
				HWND hwnd;
				lpOleInPlace->GetWindow(&hwnd);

				LPOLEOBJECT lpOle;
				if SUCCEEDED(lpunk->QueryInterface(&lpOle)) {
					LPOLECLIENTSITE lpOleClientSite = new DummyClientSite();
					result = lpOle->SetClientSite(lpOleClientSite);
					result = lpOle->DoVerb(OLEIVERB_INPLACEACTIVATE,
						NULL,
						lpOleClientSite,
						0,
						NULL,
						NULL);
				}
			}
		}
		for (auto iter = typeLibs.begin(); iter != typeLibs.end(); iter++) {
			(*iter)->Release();
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

		// Read each defined variables.
		LPVARDESC varDesc;
		for (UINT varId = 0; varId < typeAttr->cVars; ++varId) {
			result = typeInfo->GetVarDesc(varId, &varDesc);
			if FAILED(result) continue;

			FuncInfo funcInfo[2];
			result = readVarInfo(typeInfo, varDesc, funcInfo + 0, false);
			result = readVarInfo(typeInfo, varDesc, funcInfo + 1, true);
			typeInfo->ReleaseVarDesc(varDesc);
			if FAILED(result) continue;

			auto infoMap = &(output->info);
			if (isEvent) infoMap = &(output->eventInfo);
			// Put the info into the map.
			auto found = infoMap->find(funcInfo[0].name);
			if (found != infoMap->end()) {
				found->second.push_back(std::move(funcInfo[0]));
				found->second.push_back(std::move(funcInfo[1]));
			}
			else {
				(*infoMap)[funcInfo[0].name] = { funcInfo[0], funcInfo[1] };
			}
		}

		// Read each defined functions.
		LPFUNCDESC funcDesc;
		for (UINT funcId = 0; funcId < typeAttr->cFuncs; ++funcId) {
			result = typeInfo->GetFuncDesc(funcId, &funcDesc);
			if FAILED(result) continue;

			FuncInfo funcInfo;
			result = readFuncInfo(typeInfo, funcDesc, &funcInfo);
			typeInfo->ReleaseFuncDesc(funcDesc);
			if FAILED(result) continue;

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

	HRESULT readFuncInfo(LPTYPEINFO typeInfo, LPFUNCDESC funcDesc, FuncInfo * output) {
		// Stop if the function is hidden or restricted.
		WORD funcFlags = funcDesc->wFuncFlags;
		if (funcFlags & FUNCFLAG_FHIDDEN) return E_FAIL;
		if (funcFlags & FUNCFLAG_FRESTRICTED) return E_FAIL;

		FuncInfo * funcInfo = output;
		// Read documentation.
		BSTR funcName;
		BSTR funcDoc;
		typeInfo->GetDocumentation(funcDesc->memid, &funcName, &funcDoc, NULL, NULL);
		if (funcName != NULL) funcInfo->name = (std::wstring) _bstr_t(funcName, false);
		if (funcDoc != NULL) funcInfo->description = (std::wstring) _bstr_t(funcDoc, false);
		funcInfo->invokeKind = funcDesc->invkind;

		funcInfo->dispId = funcDesc->memid;
		funcInfo->vftId = funcDesc->oVft;

		// Read return type.
		funcInfo->returnType = readElemDesc(&(funcDesc->elemdescFunc));

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
			funcInfo->args.push_back(argInfo);
		}

		// Free names
		free(namesArr);
		return S_OK;
	}

	HRESULT readVarInfo(LPTYPEINFO typeInfo, LPVARDESC varDesc, FuncInfo * output, bool isWrite) {
		// Stop if the variable is hidden or restricted.
		WORD varFlags = varDesc->wVarFlags;
		if (varFlags & VARFLAG_FHIDDEN) return E_FAIL;
		if (varFlags & VARFLAG_FRESTRICTED) return E_FAIL;

		FuncInfo * funcInfo = output;
		// Read documentation.
		BSTR funcName;
		BSTR funcDoc;
		typeInfo->GetDocumentation(varDesc->memid, &funcName, &funcDoc, NULL, NULL);
		if (funcName != NULL) funcInfo->name = (std::wstring) _bstr_t(funcName, false);
		if (funcDoc != NULL) funcInfo->description = (std::wstring) _bstr_t(funcDoc, false);
		funcInfo->invokeKind = isWrite ? INVOKE_PROPERTYPUT : INVOKE_PROPERTYGET;

		funcInfo->dispId = varDesc->memid;
		funcInfo->vftId = 0;

		// Read return type.
		funcInfo->returnType = readElemDesc(&(varDesc->elemdescVar));

		if (isWrite) {
			ArgInfo argInfo;
			argInfo.flags = PARAMFLAG_FIN;
			argInfo.name = L"value";
			argInfo.type = funcInfo->returnType;
			funcInfo->args.push_back(argInfo);
		}

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
	void parseVariantTypeInfo(Environment * env, LPVARIANT input) {
		DispatchInfo * info = nullptr;
		if (input->vt == (VT_DISPATCH | VT_BYREF) || input->vt == (VT_UNKNOWN | VT_BYREF)) {
			getObjectInfo(env, *(input->ppdispVal), &info, false);
			input->vt = VT_DISPATCH | VT_BYREF;
			input->pdispVal = (LPDISPATCH) info;
		}
		if (input->vt == VT_DISPATCH || input->vt == VT_UNKNOWN) {
			getObjectInfo(env, input->pdispVal, &info, false);
			input->vt = VT_DISPATCH | VT_BYREF;
			input->pdispVal = (LPDISPATCH) info;
		}
	}
	void copyVariant(LPVARIANT input, LPVARIANT output) {
		VARTYPE type = input->vt;
		output->vt = type & (~VT_BYREF);
		switch (type & (~VT_BYREF)) {
		case VT_NULL:
			return;
		case VT_CY: {
			if (type & VT_BYREF) output->cyVal = *(input->pcyVal);
			else output->cyVal = input->cyVal;
			return;
		}
		case VT_DATE: {
			if (type & VT_BYREF) output->date = *(input->pdate);
			else output->date = input->date;
			return;
		}
		case VT_BSTR: {
			BSTR str;
			if (type & VT_BYREF) {
				str = SysAllocString(*(input->pbstrVal));
			} else {
				str = SysAllocString(input->bstrVal);
			}
			output->bstrVal = str;
			return;
		}
		case VT_UNKNOWN:
		case VT_DISPATCH: {
			output->vt = input->vt;
			output->punkVal = input->punkVal;
			// output->punkVal = NULL;
			return;
		}
		case VT_BOOL: {
			if (type & VT_BYREF) output->boolVal = *(input->pboolVal);
			else output->boolVal = input->boolVal;
			return;
		}
		case VT_VARIANT: {
			if (type & VT_BYREF) {
				copyVariant(input->pvarVal, output);
			}
			return;
		}
		case VT_DECIMAL: {
			if (type & VT_BYREF) output->decVal = *(input->pdecVal);
			else output->decVal = input->decVal;
			return;
		}
		case VT_R4: {
			if (type & VT_BYREF) output->fltVal = *(input->pfltVal);
			else output->fltVal = input->fltVal;
			return;
		}
		case VT_R8: {
			if (type & VT_BYREF) output->dblVal = *(input->pdblVal);
			else output->dblVal = input->dblVal;
			return;
		}
		case VT_UI1:
		case VT_I1: {
			if (type & VT_BYREF) output->cVal = *(input->pcVal);
			else output->cVal = input->cVal;
			return;
		}
		case VT_UI2:
		case VT_I2: {
			if (type & VT_BYREF) output->iVal = *(input->piVal);
			else output->iVal = input->iVal;
			return;
		}
		case VT_UI4:
		case VT_I4: {
			if (type & VT_BYREF) output->lVal = *(input->plVal);
			else output->lVal = input->lVal;
			return;
		}
		case VT_UI8:
		case VT_I8: {
			if (type & VT_BYREF) output->llVal = *(input->pllVal);
			else output->llVal = input->llVal;
			return;
		}
		case VT_INT: {
			if (type & VT_BYREF) output->intVal = *(input->pintVal);
			else output->intVal = input->intVal;
			return;
		}
		case VT_UINT: {
			if (type & VT_BYREF) output->uintVal = *(input->puintVal);
			else output->uintVal = input->uintVal;
			return;
		}
		case VT_SAFEARRAY: {
			// TODO
		}
		case VT_ERROR:
		case VT_HRESULT: {

		}
		}
	}
	HRESULT createEventObject(Environment * env, LPTYPEINFO typeInfo, DispatchInfo * dispInfo,
		LPUNKNOWN * output
	) {
		HRESULT result;
		LPTYPEATTR typeAttr;
		result = typeInfo->GetTypeAttr(&typeAttr);
		if FAILED(result) return result;

		EventListener * listener = new EventListener(env, dispInfo, typeInfo, typeAttr->guid);
		listener->AddRef();
		*output = (LPUNKNOWN)listener;

		typeInfo->ReleaseTypeAttr(typeAttr);

		return S_OK;
	}
}