#include "stdafx.h"
#include "globaldata.h"
#include "script.h"
#include "script_object.h"
#include "script_com.h"
#include "script_func_impl.h"
#include <DispEx.h>


// IID__IObject -- .NET's System.Object:
const IID IID__Object = {0x65074F7F, 0x63C0, 0x304E, 0xAF, 0x0A, 0xD5, 0x17, 0x41, 0xCB, 0x4A, 0x8D};

// Identifies an AutoHotkey object which was passed to a COM API and back again:
const IID IID_IObjectComCompatible = { 0x619f7e25, 0x6d89, 0x4eb4, 0xb2, 0xfb, 0x18, 0xe7, 0xc7, 0x3c, 0xe, 0xa6 };



BIF_DECL(ComObject_Call) // Formerly BIF_ComObjCreate.
{
	++aParam; // Exclude `this`
	--aParamCount;

	HRESULT hr;
	CLSID clsid, iid;
	for (;;)
	{
#ifdef UNICODE
		LPTSTR cls = TokenToString(*aParam[0]);
#else
		CStringWCharFromTChar cls = TokenToString(*aParam[0]);
#endif
		// It has been confirmed on Windows 10 that both CLSIDFromString and CLSIDFromProgID
		// were unable to resolve a ProgID starting with '{', like "{Foo", though "Foo}" works.
		// There are probably also guidelines and such that prohibit it.
		if (cls[0] == '{')
			hr = CLSIDFromString(cls, &clsid);
		else
			// CLSIDFromString is known to be able to resolve ProgIDs via the registry,
			// but fails on registration-free classes such as "Microsoft.Windows.ActCtx".
			// CLSIDFromProgID works for that, but fails when given a CLSID string
			// (consistent with VBScript and JScript in both cases).
			hr = CLSIDFromProgID(cls, &clsid);
		if (FAILED(hr)) break;

		__int64 punk = 0;
		if (aParamCount > 1)
		{
			hr = CLSIDFromString(CStringWCharFromTCharIfNeeded(TokenToString(*aParam[1])), &iid);
			if (FAILED(hr)) break;
		}
		else
			iid = IID_IDispatch;

		hr = CoCreateInstance(clsid, NULL, CLSCTX_SERVER, iid, (void **)&punk);
		if (FAILED(hr)) break;

		_f_return(new ComObject(punk, iid == IID_IDispatch ? VT_DISPATCH : VT_UNKNOWN));
	}
	ComError(hr, aResultToken);
}


BIF_DECL(BIF_ComObjGet)
{
	HRESULT hr;
	IDispatch *pdisp;
	hr = CoGetObject(CStringWCharFromTCharIfNeeded(TokenToString(*aParam[0])), NULL, IID_IDispatch, (void **)&pdisp);
	if (SUCCEEDED(hr))
	{
		if (aResultToken.object = new ComObject(pdisp))
		{
			aResultToken.symbol = SYM_OBJECT;
			return;
		}
		hr = E_OUTOFMEMORY;
		pdisp->Release();
	}
	_f_set_retval_p(_T(""), 0);
	ComError(hr, aResultToken);
}


BIF_DECL(BIF_ComObj) // Handles both ComObjFromPtr and ComValue.Call.
{
	if (!TokenIsNumeric(*aParam[0]))
		_f_throw_param(0, _T("Number"));
		
	VARTYPE vt;
	__int64 llVal;
	USHORT flags = 0;

	if (aParamCount > 1)
	{
		if (!TokenIsNumeric(*aParam[1]))
			_f_throw_param(1, _T("Number"));
		// ComValue(vt, value [, flags])
		vt = (VARTYPE)TokenToInt64(*aParam[0]);
		llVal = TokenToInt64(*aParam[1]);
		if (aParamCount > 2)
			flags = (USHORT)TokenToInt64(*aParam[2]);
	}
	else
	{
		// ComObjFromPtr(pdisp)
		vt = VT_DISPATCH;
		llVal = TokenToInt64(*aParam[0]);
		if (!(IUnknown *)llVal)
			_f_throw_param(0); // Require ComValue(9,0) for null; assume ComObjFromPtr(0) is an error.
	}

	if (vt == VT_DISPATCH || vt == VT_UNKNOWN)
	{
		IUnknown *punk = (IUnknown *)llVal;
		if (punk)
		{
			if (aParamCount == 1) // Implies above set vt = VT_DISPATCH.
			{
				IDispatch *pdisp;
				if (SUCCEEDED(punk->QueryInterface(IID_IDispatch, (void **)&pdisp)))
				{
					// Replace caller-specified interface pointer with pdisp.  v2: Caller is expected
					// to have called AddRef() if they want to keep the pointer.  In v1, nearly all
					// users of this function either passed the F_OWNVALUE flag or simply failed to
					// release their copy of the pointer, so now "taking ownership" is the default
					// behaviour and the flag has no effect on VT_DISPATCH/VT_UNKNOWN.
					punk->Release();
					llVal = (__int64)pdisp;
				}
				// Otherwise interpret it as IDispatch anyway, since caller has requested it and
				// there are known cases where it works (such as some CLR COM callable wrappers).
			}
		}
		// Otherwise, NULL may have some meaning, so allow it.  If the
		// script tries to invoke the object, it'll get a warning then.
	}

	ComObject *obj;
	if (  !(obj = new ComObject(llVal, vt, flags))  )
	{
		aResultToken.MemoryError();
		// Do any cleanup that the object may have been expected to do:
		if (vt == VT_DISPATCH || vt == VT_UNKNOWN)
			((IUnknown *)llVal)->Release();
		else if ((vt & (VT_BYREF | VT_ARRAY)) == VT_ARRAY && (flags & ComObject::F_OWNVALUE))
			SafeArrayDestroy((SAFEARRAY *)llVal);
		return;
	}
	aResultToken.symbol = SYM_OBJECT;
	aResultToken.object = obj;
}

BIF_DECL(ComValue_Call)
{
	++aParam; // Exclude `this`
	--aParamCount;
	BIF_ComObj(aResultToken, aParam, aParamCount);
}


BIF_DECL(BIF_ComObjActive)
{
	_f_set_retval_p(_T(""), 0);

	HRESULT hr;
	CLSID clsid;
	IUnknown *punk;
	hr = CLSIDFromString(CStringWCharFromTCharIfNeeded(TokenToString(*aParam[0])), &clsid);
	if (SUCCEEDED(hr))
	{
		hr = GetActiveObject(clsid, NULL, &punk);
		if (SUCCEEDED(hr))
		{
			IDispatch *pdisp;
			hr = punk->QueryInterface(IID_IDispatch, (void **)&pdisp);
			punk->Release();
			if (SUCCEEDED(hr))
			{
				if (ComObject *obj = new ComObject(pdisp))
				{
					aResultToken.symbol = SYM_OBJECT;
					aResultToken.object = obj;
					return;
				}
				hr = E_OUTOFMEMORY;
				pdisp->Release();
			}
		}
	}
	ComError(hr, aResultToken);
}


ITypeInfo *GetClassTypeInfo(IUnknown *aUnk)
{
	BOOL found = false;
	ITypeInfo *ptinfo;
	TYPEATTR *typeattr;

	// Find class typeinfo via IProvideClassInfo if available.
	// Testing shows this works in some cases where the IDispatch method fails,
	// such as for an HTMLDivElement object from an IE11 WebBrowser control.
	IProvideClassInfo *ppci;
	if (SUCCEEDED(aUnk->QueryInterface<IProvideClassInfo>(&ppci)))
	{
		found = SUCCEEDED(ppci->GetClassInfo(&ptinfo));
		ppci->Release();
		if (found)
			return ptinfo;
	}

	//
	// Find class typeinfo via IDispatch.
	//
	IDispatch *pdsp;
	ITypeLib *ptlib;
	IID iid;
	if (SUCCEEDED(aUnk->QueryInterface<IDispatch>(&pdsp)))
	{
		// Get IID and typelib of this object's IDispatch implementation.
		if (SUCCEEDED(pdsp->GetTypeInfo(0, LOCALE_USER_DEFAULT, &ptinfo)))
		{
			UINT index;
			if (SUCCEEDED(ptinfo->GetTypeAttr(&typeattr)))
			{
				iid = typeattr->guid;
				ptinfo->ReleaseTypeAttr(typeattr);
				found = SUCCEEDED(ptinfo->GetContainingTypeLib(&ptlib, &index));
			}
			ptinfo->Release();
		}
		pdsp->Release();
	}

	if (!found)
		// No typelib to search through, so give up.
		return NULL;
	found = false;

	UINT ctinfo = ptlib->GetTypeInfoCount();
	for (UINT i = 0; i < ctinfo; i++)
	{
		TYPEKIND typekind;
		// Consider only classes:
		if (FAILED(ptlib->GetTypeInfoType(i, &typekind)) || typekind != TKIND_COCLASS
			|| FAILED(ptlib->GetTypeInfo(i, &ptinfo)))
			continue;

		WORD cImplTypes = 0;
		if (SUCCEEDED(ptinfo->GetTypeAttr(&typeattr)))
		{
			cImplTypes = typeattr->cImplTypes;
			ptinfo->ReleaseTypeAttr(typeattr);
		}

		for (UINT j = 0; j < cImplTypes; j++)
		{
			INT flags;
			if (SUCCEEDED(ptinfo->GetImplTypeFlags(j, &flags)) && flags == IMPLTYPEFLAG_FDEFAULT)
			{
				// This is the default interface of the class.
				HREFTYPE reftype;
				ITypeInfo *prinfo;
				if (SUCCEEDED(ptinfo->GetRefTypeOfImplType(j, &reftype))
					&& SUCCEEDED(ptinfo->GetRefTypeInfo(reftype, &prinfo)))
				{
					if (SUCCEEDED(prinfo->GetTypeAttr(&typeattr)))
					{
						// If the IID matches, this is probably the right class.
						found = iid == typeattr->guid;
						prinfo->ReleaseTypeAttr(typeattr);
					}
					prinfo->Release();
				}
				break;
			}
		}
		if (found)
			break;
		ptinfo->Release();
	}
	ptlib->Release();
	return found ? ptinfo : NULL;
}


BIF_DECL(BIF_ComObjConnect)
{
	_f_set_retval_p(_T(""), 0);

	if (ComObject *obj = dynamic_cast<ComObject *>(TokenToObject(*aParam[0])))
	{
		if ((obj->mVarType != VT_DISPATCH && obj->mVarType != VT_UNKNOWN) || !obj->mUnknown)
		{
			ComError(-1, aResultToken); // Previously E_NOINTERFACE.
			return;
		}
		
		ITypeInfo *ptinfo;
		if (  !obj->mEventSink && (ptinfo = GetClassTypeInfo(obj->mUnknown))  )
		{
			TYPEATTR *typeattr;
			WORD cImplTypes = 0;
			if (SUCCEEDED(ptinfo->GetTypeAttr(&typeattr)))
			{
				cImplTypes = typeattr->cImplTypes;
				ptinfo->ReleaseTypeAttr(typeattr);
			}

			for(UINT j = 0; j < cImplTypes; j++)
			{
				INT flags;
				HREFTYPE reftype;
				ITypeInfo *prinfo;
				if (SUCCEEDED(ptinfo->GetImplTypeFlags(j, &flags)) && flags == (IMPLTYPEFLAG_FDEFAULT | IMPLTYPEFLAG_FSOURCE)
					&& SUCCEEDED(ptinfo->GetRefTypeOfImplType(j, &reftype))
					&& SUCCEEDED(ptinfo->GetRefTypeInfo(reftype, &prinfo)))
				{
					if (SUCCEEDED(prinfo->GetTypeAttr(&typeattr)))
					{
						if (typeattr->typekind == TKIND_DISPATCH)
						{
							obj->mEventSink = new ComEvent(obj, prinfo, typeattr->guid);
							prinfo->ReleaseTypeAttr(typeattr);
							break;
						}
						prinfo->ReleaseTypeAttr(typeattr);
					}
					prinfo->Release();
				}
			}
			ptinfo->Release();
		}

		if (obj->mEventSink)
		{
			HRESULT hr;
			if (aParamCount < 2)
				hr = obj->mEventSink->Connect(); // Disconnect.
			else
				hr = obj->mEventSink->Connect(TokenToString(*aParam[1]), TokenToObject(*aParam[1]));
			if (FAILED(hr))
				ComError(hr, aResultToken);
			return;
		}

		ComError(E_NOINTERFACE, aResultToken);
	}
	else
		ComError(-1, aResultToken); // "No COM object"
}


BIF_DECL(BIF_ComObjValue)
{
	ComObject *obj = dynamic_cast<ComObject *>(TokenToObject(*aParam[0]));
	if (!obj)
	{
		ComError(-1, aResultToken); // No COM object
		_f_return_empty;
	}
	aResultToken.value_int64 = obj->mVal64;
}


BIF_DECL(BIF_ComObjType)
{
	ComObject *obj = dynamic_cast<ComObject *>(TokenToObject(*aParam[0]));
	if (!obj)
		_f_return_empty; // So ComObjType(x) != "" can be used for "x is ComObject".
	if (aParamCount < 2)
	{
		aResultToken.value_int64 = obj->mVarType;
	}
	else
	{
		aResultToken.symbol = SYM_STRING; // for all code paths below
		aResultToken.marker = _T(""); // in case of error
		aResultToken.marker_length = 0;

		LPTSTR requested_info = TokenToString(*aParam[1]);

		ITypeInfo *ptinfo = NULL;
		if (tolower(*requested_info) == 'c')
		{
			// Get class information.
			if ((VT_DISPATCH == obj->mVarType || VT_UNKNOWN == obj->mVarType) && obj->mUnknown)
			{
				ptinfo = GetClassTypeInfo(obj->mUnknown);
				if (ptinfo)
				{
					if (!_tcsicmp(requested_info, _T("Class")))
						requested_info = _T("Name");
					else if (!_tcsicmp(requested_info, _T("CLSID")))
						requested_info = _T("IID");
				}
			}
		}
		else
		{
			// Get IDispatch information.
			if (VT_DISPATCH == obj->mVarType && obj->mDispatch)
				if (FAILED(obj->mDispatch->GetTypeInfo(0, LOCALE_USER_DEFAULT, &ptinfo)))
					ptinfo = NULL;
		}
		if (ptinfo)
		{
			if (!_tcsicmp(requested_info, _T("Name")))
			{
				BSTR name;
				if (SUCCEEDED(ptinfo->GetDocumentation(MEMBERID_NIL, &name, NULL, NULL, NULL)))
				{
					TokenSetResult(aResultToken, CStringTCharFromWCharIfNeeded(name), SysStringLen(name));
					SysFreeString(name);
				}
			}
			else if (!_tcsicmp(requested_info, _T("IID")))
			{
				TYPEATTR *typeattr;
				if (SUCCEEDED(ptinfo->GetTypeAttr(&typeattr)))
				{
					aResultToken.marker = aResultToken.buf;
#ifdef UNICODE
					aResultToken.marker_length = StringFromGUID2(typeattr->guid, aResultToken.marker, MAX_NUMBER_SIZE) - 1; // returns length including the null terminator
#else
					WCHAR cnvbuf[MAX_NUMBER_SIZE];
					StringFromGUID2(typeattr->guid, cnvbuf, MAX_NUMBER_SIZE);
					CStringCharFromWChar cnvstring(cnvbuf);
					memcpy(aResultToken.marker, cnvstring.GetBuffer(), aResultToken.marker_length = cnvstring.GetLength());
#endif
					ptinfo->ReleaseTypeAttr(typeattr);
				}
			}
			ptinfo->Release();
		}
	}
}


BIF_DECL(BIF_ComObjFlags)
{
	ComObject *obj = dynamic_cast<ComObject *>(TokenToObject(*aParam[0]));
	if (!obj)
	{
		ComError(-1, aResultToken); // No COM object
		_f_return_empty;
	}
	if (aParamCount > 1)
	{
		USHORT flags, mask;
		if (aParamCount > 2)
		{
			flags = (USHORT)TokenToInt64(*aParam[1]);
			mask = (USHORT)TokenToInt64(*aParam[2]);
		}
		else
		{
			__int64 bigflags = TokenToInt64(*aParam[1]);
			if (bigflags < 0)
			{
				// Remove specified -flags.
				flags = 0;
				mask = (USHORT)-bigflags;
			}
			else
			{
				// Add only specified flags.
				flags = (USHORT)bigflags;
				mask = flags;
			}
		}
		obj->mFlags = (obj->mFlags & ~mask) | (flags & mask);
	}
	aResultToken.value_int64 = obj->mFlags;
}


BIF_DECL(ComObjArray_Call)
{
	++aParam; // Exclude `this`
	--aParamCount;

	VARTYPE vt = (VARTYPE)TokenToInt64(*aParam[0]);
	SAFEARRAYBOUND bound[8]; // Same limit as ComObject::SafeArrayInvoke().
	int dims = aParamCount - 1;
	ASSERT(dims <= _countof(bound)); // Enforced by MaxParams.
	//if (dims > _countof(bound))
	//	_f_throw(ERR_TOO_MANY_PARAMS);
	for (int i = 0; i < dims; ++i)
	{
		bound[i].cElements = (ULONG)TokenToInt64(*aParam[i + 1]);
		bound[i].lLbound = 0;
	}
	if (SAFEARRAY *psa = SafeArrayCreate(vt, dims, bound))
	{
		if (ComObject *obj = new ComObject((__int64)psa, VT_ARRAY | vt, ComObject::F_OWNVALUE))
		{
			_f_return(obj);
		}
		SafeArrayDestroy(psa);
	}
	if (vt > 1 && vt < 0x18 && vt != 0xF)
		_f_throw_oom;
	_f_throw_param(0);
}


BIF_DECL(BIF_ComObjQuery)
{
	IUnknown *punk = NULL;
	__int64 pint = 0;
	IObject *iobj;
	HRESULT hr;

	if (auto *obj = dynamic_cast<ComObject *>(iobj = TokenToObject(*aParam[0])))
	{
		// We were passed a ComObject, but does it contain an interface pointer?
		if (obj->mVarType == VT_UNKNOWN || obj->mVarType == VT_DISPATCH)
			punk = obj->mUnknown;
	}
	if (!punk)
	{
		// Since it wasn't a valid ComObject, it should be a raw interface pointer.
		if (iobj)
		{
			// ComObject isn't handled this way since it could be VT_DISPATCH.
			UINT_PTR ptr;
			if (GetObjectPtrProperty(iobj, _T("Ptr"), ptr, aResultToken) != OK)
				return;
			punk = (IUnknown *)ptr;
		}
		else
			punk = (IUnknown *)TokenToInt64(*aParam[0]);
	}
	if (punk < (IUnknown *)65536) // Error-detection: the first 64KB of address space is always invalid.
		_f_throw_param(0);

	if (aParamCount > 2) // QueryService(obj, SID, IID)
	{
		GUID sid, iid;
		if (   SUCCEEDED(hr = CLSIDFromString(CStringWCharFromTCharIfNeeded(TokenToString(*aParam[1])), &sid))
			&& SUCCEEDED(hr = CLSIDFromString(CStringWCharFromTCharIfNeeded(TokenToString(*aParam[2])), &iid))   )
		{
			IServiceProvider *pprov;
			if (SUCCEEDED(hr = punk->QueryInterface<IServiceProvider>(&pprov)))
			{
				hr = pprov->QueryService(sid, iid, (void **)&pint);
			}
		}
	}
	else // QueryInterface(obj, IID)
	{
		GUID iid;
		if (SUCCEEDED(hr = CLSIDFromString(CStringWCharFromTCharIfNeeded(TokenToString(*aParam[1])), &iid)))
		{
			hr = punk->QueryInterface(iid, (void **)&pint);
		}
	}

	g->LastError = hr;
	if (pint)
		_f_return(new ComObject(pint, VT_UNKNOWN));
	ComError(hr, aResultToken);
}


bool SafeSetTokenObject(ExprTokenType &aToken, IObject *aObject)
{
	if (aObject)
	{
		aToken.symbol = SYM_OBJECT;
		aToken.object = aObject;
		return true;
	}
	else
	{
		aToken.symbol = SYM_STRING;
		aToken.marker = _T("");
		aToken.marker_length = 0;
		return false;
	}
}


void VariantToToken(VARIANT &aVar, ResultToken &aToken, bool aRetainVar = true)
{
	aToken.mem_to_free = NULL; // Set default.
	switch (aVar.vt)
	{
	case VT_BSTR:
		aToken.symbol = SYM_STRING;
		aToken.marker = _T(""); // Set default.
		aToken.marker_length = 0;
		size_t len;
		if (len = SysStringLen(aVar.bstrVal))
		{
#ifdef UNICODE
			if (aRetainVar)
			{
				// It's safe to pass back the actual BSTR from aVar in this case.
				aToken.marker = aVar.bstrVal;
				aToken.marker_length = len;
			}
			else
			{
				// Allocate some memory to pass back to caller:
				aToken.Malloc(aVar.bstrVal, len);
			}
#else
			CStringCharFromWChar buf(aVar.bstrVal, len);
			len = buf.GetLength(); // Get ANSI length.
			if (aToken.mem_to_free = buf.DetachBuffer())
			{
				aToken.marker = aToken.mem_to_free;
				aToken.marker_length = len;
			}
#endif
		}
		if (!aRetainVar)
			VariantClear(&aVar);
		break;
	case VT_I4:
	case VT_ERROR:
		aToken.symbol = SYM_INTEGER;
		aToken.value_int64 = aVar.lVal;
		break;
	case VT_I2:
	case VT_BOOL:
		aToken.symbol = SYM_INTEGER;
		aToken.value_int64 = aVar.iVal;
		break;
	case VT_R8:
		aToken.symbol = SYM_FLOAT;
		aToken.value_double = aVar.dblVal;
		break;
	case VT_R4:
		aToken.symbol = SYM_FLOAT;
		aToken.value_double = (double)aVar.fltVal;
		break;
	case VT_UNKNOWN:
		if (aVar.punkVal)
		{
			IEnumVARIANT *penum;
			if (SUCCEEDED(aVar.punkVal->QueryInterface(IID_IEnumVARIANT, (void**) &penum)))
			{
				if (!aRetainVar)
					aVar.punkVal->Release();
				if (!SafeSetTokenObject(aToken, new ComEnum(penum)))
					penum->Release();
				break;
			}
			IDispatch *pdisp;
			// .NET objects implement the COM IDispatch interface, but unfortunately
			// they don't always respond to QueryInterface() for said interface. So
			// instead we have to query for the _Object interface, which itself
			// inherits from IDispatch.
			if (SUCCEEDED(aVar.punkVal->QueryInterface(IID__Object, (void**) &pdisp)))
			{
				if (!aRetainVar)
					aVar.punkVal->Release();
				if (!SafeSetTokenObject(aToken, new ComObject(pdisp)))
					pdisp->Release();
				break;
			}
		}
		// FALL THROUGH to the next case:
	case VT_DISPATCH:
		if (aVar.punkVal)
		{
			IObjectComCompatible *pobj;
			if (SUCCEEDED(aVar.punkVal->QueryInterface(IID_IObjectComCompatible, (void**)&pobj)))
			{
				aToken.object = pobj;
				aToken.symbol = SYM_OBJECT;
				if (!aRetainVar)
					// QI called AddRef, so Release this reference.
					aVar.punkVal->Release();
				break;
			}
			if (aToken.object = new ComObject((__int64)aVar.punkVal, aVar.vt))
			{
				aToken.symbol = SYM_OBJECT;
				if (aRetainVar)
					// Caller is keeping their ref, so we AddRef.
					aVar.punkVal->AddRef();
				break;
			}
			if (!aRetainVar)
				// Above failed, but caller doesn't want their ref, so release it.
				aVar.punkVal->Release();
		}
		// FALL THROUGH to the next case:
	case VT_EMPTY:
	case VT_NULL:
		aToken.symbol = SYM_STRING;
		aToken.marker = _T("");
		aToken.marker_length = 0;
		break;
	default:
		{
			VARIANT var = {0};
			if (aVar.vt < VT_ARRAY // i.e. not byref or an array.
				&& SUCCEEDED(VariantChangeType(&var, &aVar, 0, VT_BSTR))) // Convert it to a BSTR.
			{
				// Recursive call to handle conversions and memory management correctly:
				VariantToToken(var, aToken, false);
			}
			else
			{
				if (!SafeSetTokenObject(aToken,
						new ComObject((__int64)aVar.parray, aVar.vt, aRetainVar ? 0 : ComObject::F_OWNVALUE)))
				{
					// Out of memory; value cannot be returned, so must be freed here.
					if (!aRetainVar)
						VariantClear(&aVar);
				}
			}
		}
	}
}

void AssignVariant(Var &aArg, VARIANT &aVar, bool aRetainVar = true)
{
	if (aVar.vt == VT_BSTR)
	{
		// Avoid an unnecessary mem alloc and copy in some cases.
		aArg.AssignStringW(aVar.bstrVal, SysStringLen(aVar.bstrVal));
		if (!aRetainVar)
			VariantClear(&aVar);
		return;
	}
	ResultToken token;
	VariantToToken(aVar, token, aRetainVar);
	switch (token.symbol)
	{
	case SYM_STRING:  // VT_BSTR was handled above, but VariantToToken coerces unhandled types to strings.
		if (token.mem_to_free)
			aArg.AcceptNewMem(token.mem_to_free, token.marker_length);
		else // Value was VT_EMPTY or VT_NULL, was coerced to an empty string, or malloc failed.
			aArg.Assign();
		break;
	case SYM_OBJECT:
		aArg.AssignSkipAddRef(token.object); // Let aArg take responsibility for it.
		break;
	default:
		aArg.Assign(token);
		break;
	}
	//token.Free(); // Above already took care of mem_to_free and object.
}


enum TTVArgType
{
	VariantIsValue,
	VariantIsAllocatedString,
	VariantIsVarRef
};

void TokenToVariant(ExprTokenType &aToken, VARIANT &aVar, TTVArgType *aVarIsArg = FALSE)
{
	if (aVarIsArg)
		*aVarIsArg = VariantIsValue;

	if (aToken.symbol == SYM_VAR)
		aToken.var->ToTokenSkipAddRef(aToken);

	switch(aToken.symbol)
	{
	case SYM_STRING:
		aVar.vt = VT_BSTR;
		aVar.bstrVal = SysAllocString(CStringWCharFromTCharIfNeeded(aToken.marker));
		if (aVarIsArg)
			*aVarIsArg = VariantIsAllocatedString;
		break;
	case SYM_INTEGER:
		{
			__int64 val = aToken.value_int64;
			if (val == (int)val)
			{
				aVar.vt = VT_I4;
				aVar.lVal = (int)val;
			}
			else
			{
				aVar.vt = VT_R8;
				aVar.dblVal = (double)val;
			}
		}
		break;
	case SYM_FLOAT:
		aVar.vt = VT_R8;
		aVar.dblVal = aToken.value_double;
		break;
	case SYM_OBJECT:
		if (ComObject *obj = dynamic_cast<ComObject *>(aToken.object))
		{
			obj->ToVariant(aVar);
			if (!aVarIsArg)
			{
				if (aVar.vt == VT_DISPATCH || aVar.vt == VT_UNKNOWN)
				{
					if (aVar.punkVal)
						aVar.punkVal->AddRef();
				}
				else if ((aVar.vt & ~VT_TYPEMASK) == VT_ARRAY && (obj->mFlags & ComObject::F_OWNVALUE))
				{
					// Copy array since both sides will call Destroy().
					if (FAILED(SafeArrayCopy(aVar.parray, &aVar.parray)))
						aVar.vt = VT_EMPTY;
				}
			}
			break;
		}
		if (aVarIsArg)
		{
			// Caller is equipped to marshal VT_BYREF|VT_VARIANT back to VarRef, so check for that.
			if (VarRef *ref = dynamic_cast<VarRef *>(aToken.object))
			{
				*aVarIsArg = VariantIsVarRef;
				aVar.vt = VT_BYREF | VT_VARIANT;
				aVar.pvarVal = (VARIANT *)malloc(sizeof(VARIANT));
				ExprTokenType token;
				ref->ToTokenSkipAddRef(token);
				TokenToVariant(token, *aVar.pvarVal, nullptr);
				break;
			}
		}
		aVar.vt = VT_DISPATCH;
		aVar.pdispVal = aToken.object;
		if (!aVarIsArg)
			aToken.object->AddRef();
		break;
	case SYM_MISSING:
		aVar.vt = VT_ERROR;
		aVar.scode = DISP_E_PARAMNOTFOUND;
		break;
	}
}


HRESULT TokenToVarType(ExprTokenType &aToken, VARTYPE aVarType, void *apValue)
// Copy the value of a given token into a variable of the given VARTYPE.
{
	if (aVarType == VT_VARIANT)
	{
		VariantClear((VARIANT *)apValue);
		TokenToVariant(aToken, *(VARIANT *)apValue, FALSE);
		return S_OK;
	}

#define U 0 // Unsupported in SAFEARRAY/VARIANT.
#define P sizeof(void *)
	// Unfortunately there appears to be no function to get the size of a given VARTYPE,
	// and VariantCopyInd() copies in the wrong direction.  An alternative approach to
	// the following would be to switch(aVarType) and copy using the appropriate pointer
	// type, but disassembly shows that approach produces larger code and internally uses
	// an array of sizes like this anyway:
	static char vt_size[] = {U,U,2,4,4,8,8,8,P,P,4,2,0,P,0,U,1,1,2,4,8,8,4,4,U,U,U,U,U,U,U,U,U,U,U,U,U,P,P};
	size_t vsize = (aVarType < _countof(vt_size)) ? vt_size[aVarType] : 0;
	if (!vsize)
		return DISP_E_BADVARTYPE;
#undef P
#undef U

	VARIANT src;
	TokenToVariant(aToken, src, FALSE);
	// Above may have set var.vt to VT_BSTR (a newly allocated string or one passed via ComObject),
	// VT_DISPATCH or VT_UNKNOWN (in which case it called AddRef()).  The value is either freed by
	// VariantChangeType() or moved into *apValue, so we don't free it except on failure.
	if (src.vt != aVarType)
	{
		// Attempt to coerce var to the correct type:
		HRESULT hr = VariantChangeType(&src, &src, 0, aVarType);
		if (FAILED(hr))
		{
			VariantClear(&src);
			return hr;
		}
	}
	// Free existing value.
	if (aVarType == VT_UNKNOWN || aVarType == VT_DISPATCH)
	{
		IUnknown *punk = *(IUnknown **)apValue;
		if (punk)
			punk->Release();
	}
	else if (aVarType == VT_BSTR)
	{
		SysFreeString(*(BSTR *)apValue);
	}
	// Write new value (shallow copy).
	memcpy(apValue, &src.lVal, vsize);
	return S_OK;
}


void VarTypeToToken(VARTYPE aVarType, void *apValue, ResultToken &aToken)
// Copy a value of the given VARTYPE into a token.
{
	VARIANT src, dst;
	src.vt = VT_BYREF | aVarType;
	src.pbVal = (BYTE *)apValue;
	dst.vt = VT_EMPTY;
	if (FAILED(VariantCopyInd(&dst, &src)))
		dst.vt = VT_EMPTY;
	VariantToToken(dst, aToken, false);
}


ResultType ComError(HRESULT hr)
{
	ResultToken errorToken;
	errorToken.SetResult(OK);
	ComError(hr, errorToken);
	return errorToken.Result();
}

void ComError(HRESULT hr, ResultToken &aResultToken, LPTSTR name, EXCEPINFO* pei)
{
	if (hr != DISP_E_EXCEPTION)
		pei = NULL;

	if (pei)
	{
		if (pei->pfnDeferredFillIn)
			(*pei->pfnDeferredFillIn)(pei);
		hr = pei->wCode ? 0x80040200 + pei->wCode : pei->scode;
	}

	TCHAR buf[4096], *error_text;
	if (hr == -1)
		error_text = _T("No valid COM object!");
	else
	{
		int size = _stprintf(buf, _T("(0x%X) "), hr);
		auto msg_buf = buf + size;
		int msg_size = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS, NULL, hr, 0, msg_buf, _countof(buf) - size - 1, NULL);
		if (msg_size)
		{
			// Remove any possible trailing \r\n.
			if (msg_buf[msg_size-1] == '\n')
				msg_buf[--msg_size] = '\0';
			if (msg_buf[msg_size-1] == '\r')
				msg_buf[--msg_size] = '\0';
			// Add a trailing \n if we'll be adding more lines below.
			if (pei)
				msg_buf[msg_size++] = '\n';
		}
		size += msg_size;
		if (pei)
			_sntprintf(buf + size, _countof(buf) - size, _T("%ws\nSource:\t%ws"), pei->bstrDescription, pei->bstrSource);
		error_text = buf;
	}

	if (pei)
	{
		SysFreeString(pei->bstrSource);
		SysFreeString(pei->bstrDescription);
		SysFreeString(pei->bstrHelpFile);
	}

	aResultToken.Error(error_text, name);
}



STDMETHODIMP ComEvent::QueryInterface(REFIID riid, void **ppv)
{
	if (riid == mIID || riid == IID_IDispatch || riid == IID_IUnknown)
	{
		AddRef();
		*ppv = this;
		return S_OK;
	}
	else
	{
		*ppv = NULL;
		return E_NOINTERFACE;
	}
}

STDMETHODIMP ComEvent::GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId)
{
	return mTypeInfo->GetIDsOfNames(rgszNames, cNames, rgDispId);
}

STDMETHODIMP ComEvent::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	if (!mObject) // mObject == NULL should be next to impossible since it is only set NULL after calling Unadvise(), in which case there shouldn't be anyone left to call this->Invoke().  Check it anyway since it might be difficult to debug, depending on what we're connected to.
		return DISP_E_MEMBERNOTFOUND;

	// Resolve method name.
	BSTR memberName;
	UINT nNames;
	if (FAILED(mTypeInfo->GetNames(dispIdMember, &memberName, 1, &nNames)))
		return DISP_E_MEMBERNOTFOUND;

	UINT cArgs = pDispParams->cArgs;
	const UINT ADDITIONAL_PARAMS = 2; // mObject (passed by us) and mAhkObject (passed by Object::Invoke).
	const UINT MAX_COM_PARAMS = MAX_FUNCTION_PARAMS - ADDITIONAL_PARAMS;
	if (cArgs > MAX_COM_PARAMS) // Probably won't happen in any real-world script.
		cArgs = MAX_COM_PARAMS; // Just omit the rest of the params.

	// Make a temporary copy of pDispParams to allow us to insert one:
	DISPPARAMS dispParams;
	VARIANTARG *vargs = (VARIANTARG *)_alloca((cArgs + 1) * sizeof(VARIANTARG));
	memcpy(&dispParams, pDispParams, sizeof(dispParams));
	memcpy(vargs + 1, pDispParams->rgvarg, cArgs * sizeof(VARIANTARG));
	dispParams.rgvarg = vargs;
	
	// Pass our object last (right-to-left) for either of the following cases:
	//	a) Our caller doesn't include its IDispatch interface pointer in the parameter list.
	//	b) The script needs a reference to the original wrapper object; i.e. mObject.
	vargs[0].vt = VT_DISPATCH;
	vargs[0].pdispVal = mObject;
	dispParams.cArgs = ++cArgs;

	HRESULT hr;
	IDispatch *func;
	DISPID dispid;

	if (mAhkObject)
	{
		func = mAhkObject;
		hr = func->GetIDsOfNames(IID_NULL, &memberName, 1, lcid, &dispid);
	}
	else
	{
		// Copy method name into our buffer, applying prefix and converting if necessary.
		TCHAR funcName[256];
		sntprintf(funcName, _countof(funcName), _T("%s%ws"), mPrefix, memberName);
		// Find the script function:
		func = g_script.FindFunc(funcName);
		dispid = DISPID_VALUE;
		hr = func ? S_OK : DISP_E_MEMBERNOTFOUND;
	}
	SysFreeString(memberName);

	if (SUCCEEDED(hr))
		hr = func->Invoke(dispid, riid, lcid, wFlags, &dispParams, pVarResult, pExcepInfo, puArgErr);

	return hr;
}

HRESULT ComEvent::Connect(LPTSTR pfx, IObject *ahkObject)
{
	HRESULT hr;

	if ((pfx != NULL) != (mCookie != 0)) // want_connection != have_connection
	{
		IConnectionPointContainer *pcpc;
		hr = mObject->mDispatch->QueryInterface(IID_IConnectionPointContainer, (void **)&pcpc);
		if (SUCCEEDED(hr))
		{
			IConnectionPoint *pconn;
			hr = pcpc->FindConnectionPoint(mIID, &pconn);
			if (SUCCEEDED(hr))
			{
				if (pfx)
				{
					hr = pconn->Advise(this, &mCookie);
				}
				else
				{
					hr = pconn->Unadvise(mCookie);
					if (SUCCEEDED(hr))
						mCookie = 0;
					if (mAhkObject) // Even if above failed:
					{
						mAhkObject->Release();
						mAhkObject = NULL;
					}
				}
				pconn->Release();
			}
			pcpc->Release();
		}
	}
	else
		hr = S_OK; // No change required.

	if (SUCCEEDED(hr))
	{
		if (mAhkObject)
			// Release this object before storing the new one below.
			mAhkObject->Release();
		// Update prefix/object.
		if (mAhkObject = ahkObject)
			mAhkObject->AddRef();
		if (pfx)
			_tcscpy(mPrefix, pfx);
		else
			*mPrefix = '\0'; // For maintainability.
		return OK;
	}
	return hr;
}

ResultType ComObject::Invoke(IObject_Invoke_PARAMS_DECL)
{
	if (mVarType != VT_DISPATCH || !mDispatch)
	{
		if (!IS_INVOKE_CALL && aName && !_tcsicmp(aName, _T("Ptr")))
		{
			if (IS_INVOKE_SET && (mVarType == VT_UNKNOWN || mVarType == VT_DISPATCH) && !mUnknown)
			{
				// Allow this assignment only in the specific cases indicated above, to avoid ambiguity
				// about what to do with the old value.  This operation is specifically intended for use
				// with DllCall; i.e. DllCall(..., "ptr*", o := ComObject(13,0)) to wrap a returned ptr.
				// Assigning zero is permitted and there is no AddRef because the caller wants us to
				// Release the interface pointer automatically.
				mUnknown = (IUnknown *)ParamIndexToInt64(0);
				return OK;
			}
			if (aParamCount)
				_o_throw(ERR_INVALID_USAGE);
			// Support passing VT_ARRAY, VT_BYREF or IUnknown to DllCall.
			if ((mVarType & (VT_ARRAY | VT_BYREF)) || mVarType == VT_UNKNOWN || mVarType == VT_DISPATCH)
				_o_return(mVal64);
		}
		if (mVarType & VT_ARRAY)
			return SafeArrayInvoke(IObject_Invoke_PARAMS);
		if (mVarType & VT_BYREF)
			return ByRefInvoke(IObject_Invoke_PARAMS);
		// Otherwise: this object can't be invoked.
		ComError(-1, aResultToken);
		return aResultToken.Result();
	}

	DISPID dispid;
	HRESULT	hr;
	if (aFlags & IF_NEWENUM)
	{
		hr = S_OK;
		dispid = DISPID_NEWENUM;
		aName = _T("_NewEnum"); // Init for ComError().
		aParamCount = 0;
	}
	else if (!aName)
	{
		hr = S_OK;
		dispid = DISPID_VALUE;
	}
	else
	{
#ifdef UNICODE
		LPOLESTR wname = aName;
#else
		CStringWCharFromChar cnvbuf(aName);
		LPOLESTR wname = (LPOLESTR)(LPCWSTR)cnvbuf;
#endif
		hr = mDispatch->GetIDsOfNames(IID_NULL, &wname, 1, LOCALE_USER_DEFAULT, &dispid);
		if (hr == DISP_E_UNKNOWNNAME) // v1.1.18: Retry with IDispatchEx if supported, to allow creating new properties.
		{
			if (IS_INVOKE_SET)
			{
				IDispatchEx *dispEx;
				if (SUCCEEDED(mDispatch->QueryInterface<IDispatchEx>(&dispEx)))
				{
					BSTR bname = SysAllocString(wname);
					// fdexNameEnsure gives us a new ID if needed, though GetIDsOfNames() will
					// still fail for some objects until after the assignment is performed below.
					hr = dispEx->GetDispID(bname, fdexNameEnsure, &dispid);
					SysFreeString(bname);
					dispEx->Release();
				}
			}
		}
		if (FAILED(hr))
			aParamCount = 0; // Skip parameter conversion and cleanup.
	}
	
	static DISPID dispidParam = DISPID_PROPERTYPUT;
	DISPPARAMS dispparams = {NULL, NULL, 0, 0};
	VARIANTARG *rgvarg;
	TTVArgType *argtype;
	EXCEPINFO excepinfo = {0};
	VARIANT varResult = {0};
	
	if (aParamCount)
	{
		rgvarg = (VARIANTARG *)_alloca(sizeof(VARIANTARG) * aParamCount);
		argtype = (TTVArgType *)_alloca(sizeof(TTVArgType) * aParamCount);

		for (int i = 1; i <= aParamCount; i++)
		{
			TokenToVariant(*aParam[aParamCount-i], rgvarg[i-1], &argtype[i-1]);
		}

		dispparams.rgvarg = rgvarg;
		dispparams.cArgs = aParamCount;
		if (IS_INVOKE_SET)
		{
			dispparams.rgdispidNamedArgs = &dispidParam;
			dispparams.cNamedArgs = 1;
		}
	}

	if (SUCCEEDED(hr)
		// For obj.x:=y where y is a ComObject, invoke PROPERTYPUTREF first:
		&& !(IS_INVOKE_SET && rgvarg[0].vt == VT_DISPATCH && SUCCEEDED(mDispatch->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUTREF, &dispparams, NULL, NULL, NULL))
		// For obj.x(), invoke METHOD first since PROPERTYGET|METHOD is ambiguous and gets undesirable results in some known cases; but re-invoke with PROPERTYGET only if DISP_E_MEMBERNOTFOUND is returned:
		  || IS_INVOKE_CALL && !aParamCount && DISP_E_MEMBERNOTFOUND != (hr = mDispatch->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dispparams, &varResult, &excepinfo, NULL))))
		// Invoke PROPERTYPUT or PROPERTYGET|METHOD as appropriate:
		hr = mDispatch->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, IS_INVOKE_SET ? DISPATCH_PROPERTYPUT : DISPATCH_PROPERTYGET | DISPATCH_METHOD, &dispparams, &varResult, &excepinfo, NULL);

	for (int i = 0; i < aParamCount; i++)
	{
		// TokenToVariant() in "arg" mode never calls AddRef() or SafeArrayCopy(), so the arg needs to be freed
		// only if it is a BSTR and not one which came from a ComObject wrapper (such as ComObject(9, pbstr)).
		switch (argtype[i])
		{
		case VariantIsAllocatedString:
			SysFreeString(rgvarg[i].bstrVal);
			break;
		case VariantIsVarRef:
			AssignVariant(*(VarRef *)aParam[aParamCount - (i + 1)]->object, *rgvarg[i].pvarVal, false);
			free(rgvarg[i].pvarVal);
			break;
		}
	}

	if (FAILED(hr))
	{
		if (hr == DISP_E_UNKNOWNNAME || hr == DISP_E_MEMBERNOTFOUND)
			return INVOKE_NOT_HANDLED;

		ComError(hr, aResultToken, aName, &excepinfo);
	}
	else if	(IS_INVOKE_SET)
	{
		// aResultToken's value is ignored by ExpandExpression() for IT_SET.
		VariantClear(&varResult);
	}
	else
	{
		VariantToToken(varResult, aResultToken, false);
	}

	return	aResultToken.Result();
}


ResultType ComObject::SafeArrayInvoke(IObject_Invoke_PARAMS_DECL)
{
	HRESULT hr = S_OK;
	SAFEARRAY *psa = (SAFEARRAY*)mVal64;
	VARTYPE item_type = (mVarType & VT_TYPEMASK);

	if (IS_INVOKE_CALL)
	{
		if (!aName)
			return INVOKE_NOT_HANDLED;

		LONG retval;
		if (!_tcsicmp(aName, _T("__Enum")))
		{
			ComArrayEnum *enm;
			if (SUCCEEDED(hr = ComArrayEnum::Begin(this, enm)))
			{
				aResultToken.symbol = SYM_OBJECT;
				aResultToken.object = enm;
			}
		}
		else if (!_tcsicmp(aName, _T("Clone")))
		{
			SAFEARRAY *clone;
			if (SUCCEEDED(hr = SafeArrayCopy(psa, &clone)))
				if (!SafeSetTokenObject(aResultToken, new ComObject((__int64)clone, mVarType, F_OWNVALUE)))
					SafeArrayDestroy(clone);
		}
		else
		{
			if (!_tcsicmp(aName, _T("MaxIndex")))
				hr = SafeArrayGetUBound(psa, aParamCount ? (UINT)TokenToInt64(*aParam[0]) : 1, &retval);
			else if (!_tcsicmp(aName, _T("MinIndex")))
				hr = SafeArrayGetLBound(psa, aParamCount ? (UINT)TokenToInt64(*aParam[0]) : 1, &retval);
			else
				hr = DISP_E_UNKNOWNNAME; // Seems slightly better than ignoring the call.
			if (SUCCEEDED(hr))
			{
				aResultToken.symbol = SYM_INTEGER;
				aResultToken.value_int64 = retval;
			}
		}
		if (FAILED(hr))
			ComError(hr, aResultToken);
		return aResultToken.Result();
	}
	if (aName && _tcsicmp(aName, _T("__Item")))
		return INVOKE_NOT_HANDLED;

	UINT dims = SafeArrayGetDim(psa);
	LONG index[8];
	// Verify correct number of parameters/dimensions (maximum 8).
	if (dims > _countof(index) || dims != (IS_INVOKE_SET ? aParamCount - 1 : aParamCount))
	{
		ComError(DISP_E_BADPARAMCOUNT, aResultToken);
		return OK;
	}
	// Build array of indices from parameters.
	for (UINT i = 0; i < dims; ++i)
	{
		if (!TokenIsNumeric(*aParam[i]))
		{
			ComError(DISP_E_BADINDEX, aResultToken);
			return OK;
		}
		index[i] = (LONG)TokenToInt64(*aParam[i]);
	}

	void *item;

	SafeArrayLock(psa);

	hr = SafeArrayPtrOfIndex(psa, index, &item);
	if (SUCCEEDED(hr))
	{
		if (IS_INVOKE_GET)
		{
			VarTypeToToken(item_type, item, aResultToken);
		}
		else // SET
		{
			ExprTokenType &rvalue = *aParam[dims];
			hr = TokenToVarType(rvalue, item_type, item);
		}
	}

	SafeArrayUnlock(psa);

	if (FAILED(hr))
		ComError(hr, aResultToken);
	return aResultToken.Result();
}


ResultType ComObject::ByRefInvoke(IObject_Invoke_PARAMS_DECL)
{
	if (IS_INVOKE_CALL || aName && _tcsicmp(aName, _T("__Item"))) // It's not [] or .__Item.
		return INVOKE_NOT_HANDLED;

	if (aParamCount > (IS_INVOKE_SET ? 1 : 0))
		_o_throw(ERR_INVALID_USAGE);
	//else: Something like x[] or x[]:=y.
	
	VARTYPE item_type = mVarType & VT_TYPEMASK;
	if (IS_INVOKE_SET)
	{
		auto hr = TokenToVarType(*aParam[0], item_type, mValPtr);
		if (FAILED(hr))
		{
			ComError(hr, aResultToken);
			return aResultToken.Result();
		}
	}
	else
	{
		VarTypeToToken(item_type, mValPtr, aResultToken);
	}
	return OK;
}


LPTSTR ComObject::Type()
{
	if ((mVarType == VT_DISPATCH || mVarType == VT_UNKNOWN) && mUnknown)
	{
		BSTR name;
		ITypeInfo *ptinfo;
		// Use COM class name if available.
		if (  (ptinfo = GetClassTypeInfo(mUnknown))
			&& SUCCEEDED(ptinfo->GetDocumentation(MEMBERID_NIL, &name, NULL, NULL, NULL))  )
		{
			static TCHAR sBuf[64]; // Seems generous enough.
			tcslcpy(sBuf, CStringTCharFromWCharIfNeeded(name), _countof(sBuf));
			SysFreeString(name);
			return sBuf;
		}
	}
	ExprTokenType value;
	if (Base()->GetOwnProp(value, _T("__Class")))
		return TokenToString(value);
	return _T("ComValue"); // Provide a safe default in case __Class was removed.
}


Object *ComObject::Base()
{
	if (mVarType & VT_ARRAY)
		return Object::sComArrayPrototype;
	if (mVarType & VT_BYREF)
		return Object::sComRefPrototype;
	if (mVarType == VT_DISPATCH && mUnknown)
		return Object::sComObjectPrototype;
	return Object::sComValuePrototype;
}


ResultType ComEnum::Next(Var *aOutput, Var *aOutputType)
{
	VARIANT varResult = {0};
	if (penum->Next(1, &varResult, NULL) == S_OK)
	{
		if (aOutputType)
			aOutputType->Assign((__int64)varResult.vt);
		if (aOutput)
			AssignVariant(*aOutput, varResult, false);
		return CONDITION_TRUE;
	}
	return CONDITION_FALSE;
}


HRESULT ComArrayEnum::Begin(ComObject *aArrayObject, ComArrayEnum *&aEnum)
{
	HRESULT hr;
	SAFEARRAY *psa = aArrayObject->mArray;
	char *arrayData, *arrayEnd;
	long lbound, ubound;

	if (SafeArrayGetDim(psa) != 1)
		return E_NOTIMPL;
	
	if (   SUCCEEDED(hr = SafeArrayGetLBound(psa, 1, &lbound))
		&& SUCCEEDED(hr = SafeArrayGetUBound(psa, 1, &ubound))
		&& SUCCEEDED(hr = SafeArrayAccessData(psa, (void **)&arrayData))   )
	{
		VARTYPE arrayType = aArrayObject->mVarType & VT_TYPEMASK;
		UINT elemSize = SafeArrayGetElemsize(psa);
		arrayEnd = arrayData + (ubound - lbound) * (long)elemSize; // Must cast to signed long for correct result when array is empty (ubound - lbound == -1).
		if (aEnum = new ComArrayEnum(aArrayObject, arrayData, arrayEnd, elemSize, arrayType))
		{
			aArrayObject->AddRef(); // Keep obj alive until enumeration completes.
		}
		else
		{
			SafeArrayUnaccessData(psa);
			hr = E_OUTOFMEMORY;
		}
	}
	return hr;
}

ComArrayEnum::~ComArrayEnum()
{
	SafeArrayUnaccessData(mArrayObject->mArray); // Counter the lock done by ComArrayEnum::Begin.
	mArrayObject->Release();
}

ResultType ComArrayEnum::Next(Var *aOutput, Var *aOutputType)
{
	if ((mPointer += mElemSize) <= mEnd)
	{
		VARIANT var = {0};
		if (mType == VT_VARIANT)
		{
			// Make shallow copy of the VARIANT item.
			memcpy(&var, mPointer, sizeof(VARIANT));
		}
		else
		{
			// Build VARIANT with shallow copy of the item.
			var.vt = mType;
			memcpy(&var.lVal, mPointer, mElemSize);
		}
		if (aOutput)
			AssignVariant(*aOutput, var);
		if (aOutputType)
			aOutputType->Assign(var.vt);
		return CONDITION_TRUE;
	}
	return CONDITION_FALSE;
}


IObject *GuiType::ControlGetActiveX(HWND aWnd)
{
	typedef HRESULT (WINAPI *MyAtlAxGetControl)(HWND h, IUnknown **p);
	static MyAtlAxGetControl fnAtlAxGetControl = NULL;
	if (!fnAtlAxGetControl)
		if (HMODULE hmodAtl = GetModuleHandle(_T("atl"))) // GuiType::AddControl should have already permanently loaded it.
			fnAtlAxGetControl = (MyAtlAxGetControl)GetProcAddress(hmodAtl, "AtlAxGetControl");
	if (fnAtlAxGetControl) // Should always be non-NULL if aControl is actually an ActiveX control.
	{
		IUnknown *punk;
		if (SUCCEEDED(fnAtlAxGetControl(aWnd, &punk)))
		{
			IObject *pobj;
			IDispatch *pdisp;
			if (SUCCEEDED(punk->QueryInterface(IID_IDispatch, (void **)&pdisp)))
			{
				punk->Release();
				if (  !(pobj = new ComObject(pdisp))  )
					pdisp->Release();
			}
			else
			{
				if (  !(pobj = new ComObject((__int64)punk, VT_UNKNOWN))  )
					punk->Release();
			}
			return pobj;
		}
	}
	return NULL;
}


STDMETHODIMP IObjectComCompatible::QueryInterface(REFIID riid, void **ppv)
{
	if (riid == IID_IDispatch || riid == IID_IUnknown || riid == IID_IObjectComCompatible)
	{
		AddRef();
		//*ppv = static_cast<IDispatch *>(this);
		*ppv = this;
		return S_OK;
	}
	else
	{
		*ppv = NULL;
		return E_NOINTERFACE;
	}
}

STDMETHODIMP IObjectComCompatible::GetTypeInfoCount(UINT *pctinfo)
{
	*pctinfo = 0;
	return S_OK;
}

STDMETHODIMP IObjectComCompatible::GetTypeInfo(UINT itinfo, LCID lcid, ITypeInfo **pptinfo)
{
	*pptinfo = NULL;
	return E_NOTIMPL;
}

static Array *g_IdToName;
static Object *g_NameToId;

STDMETHODIMP IObjectComCompatible::GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId)
{
#ifdef UNICODE
	LPTSTR name = *rgszNames;
#else
	CStringCharFromWChar name_buf(*rgszNames);
	LPTSTR name = const_cast<LPTSTR>(name_buf.GetString());
#endif
	if ( !(g_IdToName || (g_IdToName = Array::Create())) ||
		 !(g_NameToId || (g_NameToId = Object::Create())) )
		return E_OUTOFMEMORY;
	ExprTokenType id;
	if (!g_NameToId->GetOwnProp(id, name))
	{
		if (!g_IdToName->Append(name))
			return E_OUTOFMEMORY;
		id.symbol = SYM_INTEGER;
		id.value_int64 = g_IdToName->Length();
		if (!g_NameToId->SetOwnProp(name, id))
			return E_OUTOFMEMORY;
	}
	*rgDispId = (DISPID)id.value_int64;
	if (cNames == 1)
		return S_OK;
	for (UINT i = 1; i < cNames; ++i)
		rgDispId[i] = DISPID_UNKNOWN;
	return DISP_E_UNKNOWNNAME;
}

STDMETHODIMP IObjectComCompatible::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	ResultToken param_token[MAX_FUNCTION_PARAMS];
	ExprTokenType *param[MAX_FUNCTION_PARAMS];
	
	UINT cArgs = pDispParams->cArgs;
	if (cArgs >= MAX_FUNCTION_PARAMS) // Probably won't happen in any real-world script.
		cArgs = MAX_FUNCTION_PARAMS - 1; // Just omit the rest of the params.

	// PUT is probably never combined with GET/METHOD, so is handled first.  Some common problem
	// cases which combine GET and METHOD include:
	//  - Foo.Bar in VBScript.
	//  - foo.bar() and foo.bar[] in C#.  There's no way to differentiate, so we just use METHOD
	//    when there are parameters to cover the most common cases.  Both will work with user-
	//    defined properties (v1.1.16+) since they can be "called", but won't work with __Get.
	//  - foo[bar] and foo(bar) in C# both use METHOD|GET with DISPID_VALUE.
	//  - foo(bar) (but not foo[bar]) in JScript uses METHOD|GET with DISPID_VALUE.
	int flags = (wFlags & (DISPATCH_PROPERTYPUT | DISPATCH_PROPERTYPUTREF)) ? IT_SET
		: (wFlags & DISPATCH_METHOD) ? IT_CALL : IT_GET;
	
	ExprTokenType **first_param = param + 1;
	int param_count = cArgs;
	LPTSTR name;
	
	if (dispIdMember > 0)
	{
		if (!g_IdToName->ItemToToken(dispIdMember - 1, param_token[0]))
			return DISP_E_MEMBERNOTFOUND;
		name = param_token[0].marker;
	}
	else
	{
		if (dispIdMember != DISPID_VALUE)
			return DISP_E_MEMBERNOTFOUND;
		name = nullptr;
	}
	
	for (UINT i = 1; i <= cArgs; ++i)
	{
		VARIANTARG &varg = pDispParams->rgvarg[cArgs-i];
		if (varg.vt & VT_BYREF)
		{
			// Allocate and pass a temporary Var to transparently support ByRef.
			auto varref = new VarRef();
			param_token[i].SetValue(varref);
			param_token[i].mem_to_free = nullptr;
			if (varg.vt == (VT_BYREF | VT_BSTR))
			{
				// Avoid an unnecessary string-copy that would be made by VariantCopyInd().
				varref->AssignStringW(*varg.pbstrVal, SysStringLen(*varg.pbstrVal));
			}
			else if (varg.vt == (VT_BYREF | VT_VARIANT))
			{
				AssignVariant(*varref, *varg.pvarVal);
			}
			else
			{
				VARIANT value;
				value.vt = VT_EMPTY;
				VariantCopyInd(&value, &varg);
				AssignVariant(*varref, value, false); // false causes value's memory/ref to be transferred to var or freed.
			}
		}
		else
		{
			VariantToToken(varg, param_token[i]);
		}
		param[i] = &param_token[i];
	}
	
	ExprTokenType this_token;
	this_token.symbol = SYM_OBJECT;
	this_token.object = this;

	HRESULT result_to_return;
	FuncResult result_token;
	int outer_excptmode = g->ExcptMode;
	g->ExcptMode |= EXCPTMODE_CATCH; // Indicate exceptions will be handled (by our caller, the COM client).

	for (;;)
	{
		switch (static_cast<IObject *>(this)->Invoke(result_token, flags, name, this_token, first_param, param_count))
		{
		case FAIL:
			result_to_return = E_FAIL;
			if (g->ThrownToken)
			{
				Object *obj;
				if (pExcepInfo && (obj = dynamic_cast<Object*>(TokenToObject(*g->ThrownToken)))) // MSDN: pExcepInfo "Can be NULL"
				{
					ZeroMemory(pExcepInfo, sizeof(EXCEPINFO));
					pExcepInfo->scode = result_to_return = DISP_E_EXCEPTION;
					
					#define SysStringFromToken(...) \
						SysAllocString(CStringWCharFromTCharIfNeeded(TokenToString(__VA_ARGS__)))

					ExprTokenType token;
					if (obj->GetOwnProp(token, _T("Message")))
						pExcepInfo->bstrDescription = SysStringFromToken(token, result_token.buf);
					if (obj->GetOwnProp(token, _T("What")))
						pExcepInfo->bstrSource = SysStringFromToken(token, result_token.buf);
					if (obj->GetOwnProp(token, _T("File")))
						pExcepInfo->bstrHelpFile = SysStringFromToken(token, result_token.buf);
					if (obj->GetOwnProp(token, _T("Line")))
						pExcepInfo->dwHelpContext = (DWORD)TokenToInt64(token);
				}
				else
				{
					// Showing an error message seems more helpful than silently returning E_FAIL.
					// Although there's risk of our caller showing a second message or handling
					// the error in some other way, it couldn't report anything specific since
					// all it will have is the generic failure code.
					g_script.UnhandledException(NULL);
				}
				g_script.FreeExceptionToken(g->ThrownToken);
			}
			break;
		case INVOKE_NOT_HANDLED:
			if ((flags & IT_BITMASK) == IT_CALL && (wFlags & DISPATCH_PROPERTYGET))
			{
				// Request was to call a method OR retrieve a property, but we now know there's no such
				// method.  For DISPID_VALUE, it could be Obj(X) in VBScript and C#, or Obj[X] in C#.
				// For other IDs, Obj.X(Y) in VBScript is known to do this.  VBScript uses () for both
				// functions and array indexing.
				flags = IT_GET;
				continue;
			}
			result_to_return = DISP_E_MEMBERNOTFOUND;
			break;
		default:
			result_to_return = S_OK;
			if (pVarResult)
				TokenToVariant(result_token, *pVarResult, FALSE);
		}
		break;
	}

	g->ExcptMode = outer_excptmode;

	// Clean up:
	result_token.Free();
	for (UINT i = 1; i <= cArgs; ++i)
	{
		VARIANTARG &varg = pDispParams->rgvarg[cArgs-i];
		if (varg.vt & VT_BYREF)
		{
			ASSERT(param_token[i].symbol == SYM_OBJECT && dynamic_cast<VarRef *>(param_token[i].object));
			auto varref = (VarRef *)param_token[i].object;
			ExprTokenType value;
			varref->ToTokenSkipAddRef(value);
			TokenToVarType(value, varg.vt & ~VT_BYREF, varg.pvRecord);
		}
		param_token[i].Free();
	}

	return result_to_return;
}


#ifdef CONFIG_DEBUGGER

void WriteComObjType(IDebugProperties *aDebugger, ComObject *aObject, LPCSTR aName, LPTSTR aWhichType)
{
	TCHAR buf[_f_retval_buf_size];
	ResultToken resultToken;
	resultToken.symbol = SYM_INTEGER;
	resultToken.marker_length = -1;
	resultToken.mem_to_free = NULL;
	resultToken.buf = buf;
	ExprTokenType paramToken[] = { aObject, aWhichType };
	ExprTokenType *param[] = { &paramToken[0], &paramToken[1] };
	BIF_ComObjType(resultToken, param, 2);
	aDebugger->WriteProperty(aName, resultToken);
	if (resultToken.mem_to_free)
		free(resultToken.mem_to_free);
}

void ComObject::DebugWriteProperty(IDebugProperties *aDebugger, int aPage, int aPageSize, int aDepth)
{
	DebugCookie rootCookie;
	aDebugger->BeginProperty(NULL, "object", 2 + (mVarType == VT_DISPATCH)*2 + (mEventSink != NULL), rootCookie);
	if (aPage == 0)
	{
		// For simplicity, assume they all fit within aPageSize.
		
		aDebugger->WriteProperty("Value", ExprTokenType(mVal64));
		aDebugger->WriteProperty("VarType", ExprTokenType((__int64)mVarType));

		if (mVarType == VT_DISPATCH)
		{
			WriteComObjType(aDebugger, this, "DispatchType", _T("Name"));
			WriteComObjType(aDebugger, this, "DispatchIID", _T("IID"));
		}
		
		if (mEventSink)
		{
			DebugCookie sinkCookie;
			aDebugger->BeginProperty("EventSink", "object", 2, sinkCookie);
			
			if (mEventSink->mAhkObject)
				aDebugger->WriteProperty("Object", ExprTokenType(mEventSink->mAhkObject));
			else
				aDebugger->WriteProperty("Prefix", ExprTokenType(mEventSink->mPrefix));
			
			OLECHAR buf[40];
			if (!StringFromGUID2(mEventSink->mIID, buf, _countof(buf)))
				*buf = 0;
			aDebugger->WriteProperty("IID", ExprTokenType((LPTSTR)(LPCTSTR)CStringTCharFromWCharIfNeeded(buf)));
			
			aDebugger->EndProperty(sinkCookie);
		}
	}
	aDebugger->EndProperty(rootCookie);
}

#endif
