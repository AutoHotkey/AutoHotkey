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
		_f_return(new ComObject(pdisp));
	}
	_f_set_retval_p(_T(""), 0);
	ComError(hr, aResultToken);
}


BIF_DECL(BIF_ComObj) // Handles both ComObjFromPtr and ComValue.Call.
{
	if (!TokenIsNumeric(*aParam[0]))
		_f_throw_param(0, _T("Number"));
		
	VARTYPE vt;
	union
	{
		__int64 llVal = 0; // Must be zero-initialized for TokenToVarType().
		double dblVal;
		float fltVal;
	};
	HRESULT hr;
	USHORT flags = 0;

	if (aParamCount > 1)
	{
		// ComValue(vt, value [, flags])
		if (aParamCount > 2)
		{
			Throw_if_Param_NaN(2);
			flags = (USHORT)TokenToInt64(*aParam[2]);
		}
		vt = (VARTYPE)TokenToInt64(*aParam[0]);
		if (FAILED(hr = TokenToVarType(*aParam[1], vt, &llVal, true)))
			return ComError(hr, aResultToken);
		if (vt == VT_BSTR && TypeOfToken(*aParam[1]) != SYM_INTEGER)
			flags |= ComObject::F_OWNVALUE; // BSTR was allocated above, so make sure it will be freed later.
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

	_f_return(new ComObject(llVal, vt, flags));
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
				_f_return(new ComObject(pdisp));
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
	
	LPCTSTR prefix = nullptr;    // Set default: disconnect.
	IObject *handlers = nullptr; //
	if (!ParamIndexIsOmitted(1))
	{
		prefix = ParamIndexToString(1);
		handlers = ParamIndexToObject(1);
		if (!*prefix && !handlers) // Blank or pure numeric (which is not valid as a prefix).
			_f_throw_param(1);
	}

	if (ComObject *obj = dynamic_cast<ComObject *>(TokenToObject(*aParam[0])))
	{
		if ((obj->mVarType != VT_DISPATCH && obj->mVarType != VT_UNKNOWN) || !obj->mUnknown)
		{
			_f_throw_param(0);
		}

		bool already_connected = obj->mEventSink;
		if (already_connected)
		{
			if (!prefix)
			{
				HRESULT hr = obj->mEventSink->Connect(); // This should result in mEventSink being deleted.
				if (FAILED(hr))
					ComError(hr, aResultToken);
				return;
			}
		}
		else
			obj->mEventSink = new ComEvent(obj);
		
		// Set or update prefix/event sink prior to calling Advise().
		obj->mEventSink->SetPrefixOrSink(prefix, handlers);
		if (already_connected)
			return;
		
		HRESULT hr = E_NOINTERFACE;
		
		if (ITypeInfo *ptinfo = GetClassTypeInfo(obj->mUnknown))
		{
			TYPEATTR *typeattr;
			WORD cImplTypes = 0;
			if (SUCCEEDED(ptinfo->GetTypeAttr(&typeattr)))
			{
				cImplTypes = typeattr->cImplTypes;
				ptinfo->ReleaseTypeAttr(typeattr);
			}

			for (UINT j = 0; j < cImplTypes; j++)
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
							hr = obj->mEventSink->Connect(prinfo, &typeattr->guid);
							// Let the connection point's reference be the only one, so we can detect when it releases.
							// If Connect() failed, this will delete the event sink, which will set mEventSink = nullptr.
							obj->mEventSink->Release();
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

		if (FAILED(hr))
			ComError(hr, aResultToken);
	}
	else
		_f_throw_param(0, _T("ComValue"));
}


ComEvent::~ComEvent()
{
	// At this point, all references have been released, including the
	// reference held by the connection point (if a connection was made),
	// so Unadvise() isn't necessary and couldn't be successful anyway.
	if (mObject)
		mObject->mEventSink = nullptr;
	if (mTypeInfo)
		mTypeInfo->Release();
	if (mAhkObject)
		mAhkObject->Release();
}


BIF_DECL(BIF_ComObjValue)
{
	ComObject *obj = dynamic_cast<ComObject *>(TokenToObject(*aParam[0]));
	if (!obj)
		_f_throw_param(0, _T("ComValue"));
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
			else
				aResultToken.ParamError(1, aParam[1]);
			ptinfo->Release();
		}
	}
}


BIF_DECL(BIF_ComObjFlags)
{
	ComObject *obj = dynamic_cast<ComObject *>(TokenToObject(*aParam[0]));
	if (!obj)
		_f_throw_param(0, _T("ComValue"));
	if (aParamCount > 1)
	{
		Throw_if_Param_NaN(1);
		USHORT flags, mask;
		if (aParamCount > 2)
		{
			Throw_if_Param_NaN(2);
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

	Throw_if_Param_NaN(0);
	VARTYPE vt = (VARTYPE)TokenToInt64(*aParam[0]);
	SAFEARRAYBOUND bound[8]; // Same limit as ComObject::SafeArrayInvoke().
	int dims = aParamCount - 1;
	ASSERT(dims <= _countof(bound)); // Enforced by MaxParams.
	//if (dims > _countof(bound))
	//	_f_throw(ERR_TOO_MANY_PARAMS);
	for (int i = 0; i < dims; ++i)
	{
		Throw_if_Param_NaN(i + 1);
		bound[i].cElements = (ULONG)TokenToInt64(*aParam[i + 1]);
		bound[i].lLbound = 0;
	}
	if (SAFEARRAY *psa = SafeArrayCreate(vt, dims, bound))
	{
		_f_return(new ComObject((__int64)psa, VT_ARRAY | vt, ComObject::F_OWNVALUE));
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

	GUID iid;
	if (aParamCount > 2) // QueryService(obj, SID, IID)
	{
		GUID sid;
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
		if (SUCCEEDED(hr = CLSIDFromString(CStringWCharFromTCharIfNeeded(TokenToString(*aParam[1])), &iid)))
		{
			hr = punk->QueryInterface(iid, (void **)&pint);
		}
	}

	if (pint)
		_f_return(new ComObject(pint, iid == IID_IDispatch ? VT_DISPATCH : VT_UNKNOWN));
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
		aToken.symbol = SYM_INTEGER;
		aToken.value_int64 = aVar.lVal;
		break;
	case VT_R8:
		aToken.symbol = SYM_FLOAT;
		aToken.value_double = aVar.dblVal;
		break;
	case VT_R4:
		aToken.symbol = SYM_FLOAT;
		aToken.value_double = (double)aVar.fltVal;
		break;
	case VT_BOOL:
		aToken.symbol = SYM_INTEGER;
		aToken.value_int64 = aVar.iVal != VARIANT_FALSE;
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
			if (aRetainVar)
				// Caller is keeping their ref, so we AddRef.
				aVar.punkVal->AddRef();
			aToken.object = new ComObject((__int64)aVar.punkVal, aVar.vt);
			aToken.symbol = SYM_OBJECT;
			break;
		}
		// FALL THROUGH to the next case:
	case VT_EMPTY:
	case VT_NULL:
		aToken.symbol = SYM_STRING;
		aToken.marker = _T("");
		aToken.marker_length = 0;
		break;
	case VT_I1:
	case VT_I2:
	case VT_I8:
	case VT_UI1:
	case VT_UI2:
	case VT_UI4:
	case VT_UI8:
	{
		VARIANT var {};
		VariantChangeType(&var, &aVar, 0, VT_I8);
		aToken.symbol = SYM_INTEGER;
		aToken.value_int64 = var.llVal;
		break;
	}
	case VT_ERROR:
		if (aVar.scode == DISP_E_PARAMNOTFOUND)
		{
			aToken.symbol = SYM_MISSING;
			break;
		}
		// FALL THROUGH to the next case:
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


void TokenToVariant(ExprTokenType &aToken, VARIANT &aVar, TTVArgType *aVarIsArg)
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
		aVar.llVal = aToken.value_int64;
		aVar.vt = (aVar.llVal == (int)aVar.llVal) ? VT_I4 : VT_I8;
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
				else if (obj->mFlags & ComObject::F_OWNVALUE)
				{
					if ((aVar.vt & ~VT_TYPEMASK) == VT_ARRAY)
					{
						// Copy array since both sides will call Destroy().
						if (FAILED(SafeArrayCopy(aVar.parray, &aVar.parray)))
							aVar.vt = VT_EMPTY;
					}
					else if (aVar.vt == VT_BSTR)
					{
						// Copy the string.
						aVar.bstrVal = SysAllocStringLen(aVar.bstrVal, SysStringLen(aVar.bstrVal));
					}
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
				if (token.symbol == SYM_MISSING)
					aVar.pvarVal->vt = VT_EMPTY; // Seems better than BSTR("") or ERROR(DISP_E_PARAMNOTFOUND).
				else
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


HRESULT TokenToVarType(ExprTokenType &aToken, VARTYPE aVarType, void *apValue, bool aCallerIsComValue)
// Copy the value of a given token into a variable of the given VARTYPE.
{
	if (aVarType == VT_VARIANT)
	{
		VariantClear((VARIANT *)apValue);
		TokenToVariant(aToken, *(VARIANT *)apValue, FALSE);
		return S_OK;
	}

#define U 0 // Unsupported in SAFEARRAY/VARIANT (but vt 15 has no assigned meaning).
#define P sizeof(void *)
	// Unfortunately there appears to be no function to get the size of a given VARTYPE,
	// and VariantCopyInd() copies in the wrong direction.  An alternative approach to
	// the following would be to switch(aVarType) and copy using the appropriate pointer
	// type, but disassembly shows that approach produces larger code and internally uses
	// an array of sizes like this anyway:
	static const char vt_size[] = {U,U,2,4,4,8,8,8,P,P,4,2,U,P,U,U,1,1,2,4,8,8,4,4,U,U,U,U,U,U,U,U,U,U,U,U,U,P,P};
	// Checking aCallerIsComValue has these purposes:
	//  - Let ComValue() accept a raw pointer/integer value for VT_BYREF, VT_ARRAY and
	//    any types not explicitly supported.
	//  - When an integer value is passed by ComValue(), copy all 64 bits.
	size_t vsize = aCallerIsComValue ? 8 : (aVarType < _countof(vt_size)) ? vt_size[aVarType] : 0;
	if (!vsize)
		return DISP_E_BADVARTYPE;
#undef P
#undef U

	VARIANT src;

	if (aVarType == VT_BOOL)
	{
		// Use AutoHotkey's boolean evaluation rules, but VARIANT_TRUE == -1.
		*((VARIANT_BOOL*)apValue) = -TokenToBOOL(aToken);
		return S_OK;
	}

	if (TypeOfToken(aToken) == SYM_INTEGER)
	{
		// This has a few uses:
		//  - Allows pointer types to be initialized by pointer value for ComValue().
		//  - Avoids truncation or loss of precision for large integer values,
		//    since TokenToVariant() only uses the common VT_I4 or VT_R8 types.
		//  - Probably faster.
		src.vt = VT_I8;
		src.llVal = TokenToInt64(aToken);
		switch (aVarType)
		{
		case VT_R4:
		case VT_R8:
		case VT_DATE: // Date is "double" based.
			// Doesn't make sense to reinterpret the bits of an integer value as float.
			break;
		case VT_CY:
			// This ensures 42 and 42.0 produce the same value.
			src.llVal *= 10000;
			*((CY*)apValue) = src.cyVal;
			return S_OK;
		case VT_BSTR:
		case VT_DISPATCH:
		case VT_UNKNOWN:
			if (!aCallerIsComValue)
				break;
			// Otherwise, fall through:
		default:
			memcpy(apValue, &src.llVal, vsize);
			return S_OK;
		}
	}
	else
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

	TCHAR buf[4096];
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
	}
	size += msg_size;
	if (pei)
	{
		if (pei->bstrDescription)
			size += _sntprintf(buf + size, _countof(buf) - size, _T("\n%ws"), pei->bstrDescription);
		if (pei->bstrSource)
			size += _sntprintf(buf + size, _countof(buf) - size, _T("\nSource:\t%ws"), pei->bstrSource);
	}

	if (pei)
	{
		SysFreeString(pei->bstrSource);
		SysFreeString(pei->bstrDescription);
		SysFreeString(pei->bstrHelpFile);
	}

	aResultToken.Error(buf, name);
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
	if (!mObject) // mObject == NULL should be impossible unless Unadvise() somehow fails while leaving behind a valid connection.  Check it anyway since it might be difficult to debug, depending on what we're connected to.
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
		func = g_script.FindGlobalFunc(funcName);
		dispid = DISPID_VALUE;
		hr = func ? S_OK : DISP_E_MEMBERNOTFOUND;
	}
	SysFreeString(memberName);

	if (SUCCEEDED(hr))
		hr = func->Invoke(dispid, riid, lcid, wFlags, &dispParams, pVarResult, pExcepInfo, puArgErr);

	return hr;
}

HRESULT ComEvent::Connect(ITypeInfo *tinfo, IID *iid)
{
	bool want_to_connect = tinfo;
	if (want_to_connect)
	{
		// Set these unconditionally to ensure they are released on failure or disconnection.
		ASSERT(!mTypeInfo && iid);
		mTypeInfo = tinfo;
		mIID = *iid;
	}
	
	IConnectionPointContainer *pcpc;
	HRESULT hr = mObject->mDispatch->QueryInterface(IID_IConnectionPointContainer, (void **)&pcpc);
	if (SUCCEEDED(hr))
	{
		IConnectionPoint *pconn;
		hr = pcpc->FindConnectionPoint(mIID, &pconn);
		if (SUCCEEDED(hr))
		{
			if (want_to_connect)
				hr = pconn->Advise(this, &mCookie);
			else if (mCookie != 0) // This check preserves legacy behaviour of ComObjConnect(obj) without a prior connection.
				hr = pconn->Unadvise(mCookie); // This should result in this ComEvent being deleted.
			pconn->Release();
		}
		pcpc->Release();
	}
	return hr;
}

void ComEvent::SetPrefixOrSink(LPCTSTR pfx, IObject *ahkObject)
{
	if (mAhkObject)
	{
		mAhkObject->Release();
		mAhkObject = NULL;
	}
	if (ahkObject)
	{
		ahkObject->AddRef();
		mAhkObject = ahkObject;
	}
	if (pfx)
		tcslcpy(mPrefix, pfx, _countof(mPrefix));
}

ResultType ComObject::Invoke(IObject_Invoke_PARAMS_DECL)
{
	if (mVarType != VT_DISPATCH || !mDispatch)
	{
		IObject *proto;
		if (mVarType & VT_ARRAY)
			proto = Object::sComArrayPrototype;
		else if (mVarType & VT_BYREF)
			proto = Object::sComRefPrototype;
		else
			proto = Object::sComValuePrototype;
		return proto->Invoke(aResultToken, aFlags | IF_SUBSTITUTE_THIS, aName, aThisToken, aParam, aParamCount);
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



ObjectMemberMd ComObject::sArrayMembers[]
{
	md_member_x(ComObject, __Item, SafeArray_Item, GET, (Ret, Variant, RetVal), (In, Params, Index)),
	md_member_x(ComObject, __Item, SafeArray_Item, SET, (In, Variant, Value), (In, Params, Index)),
	md_member_x(ComObject, __Enum, SafeArray_Enum, CALL, (In_Opt, Int32, N), (Ret, Object, RetVal)),
	md_member_x(ComObject, Clone, SafeArray_Clone, CALL, (Ret, Object, RetVal)),
	md_member_x(ComObject, MaxIndex, SafeArray_MaxIndex, CALL, (In_Opt, UInt32, Dims), (Ret, Int32, RetVal)),
	md_member_x(ComObject, MinIndex, SafeArray_MinIndex, CALL, (In_Opt, UInt32, Dims), (Ret, Int32, RetVal))
};

ObjectMember ComObject::sRefMembers[]
{
	Object_Property_get_set(__Item),
};

ObjectMember ComObject::sValueMembers[]
{
	Object_Property_get_set(Ptr),
};


void DefineComPrototypeMembers()
{
	Object::DefineMembers(Object::sComValuePrototype, _T("ComValue"), ComObject::sValueMembers, _countof(ComObject::sValueMembers));
	Object::DefineMembers(Object::sComRefPrototype, _T("ComValueRef"), ComObject::sRefMembers, _countof(ComObject::sRefMembers));
	Object::DefineMetadataMembers(Object::sComArrayPrototype, _T("ComObjArray"), ComObject::sArrayMembers, _countof(ComObject::sArrayMembers));
}


void ComObject::Invoke(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	if (aID == P___Item) // ComValueRef
	{
		VARTYPE item_type = mVarType & VT_TYPEMASK;
		if (IS_INVOKE_SET)
		{
			auto hr = TokenToVarType(*aParam[0], item_type, mValPtr);
			if (FAILED(hr))
			{
				ComError(hr, aResultToken);
				return;
			}
		}
		else
		{
			VarTypeToToken(item_type, mValPtr, aResultToken);
		}
	}
	else // P_Ptr
	{
		if (IS_INVOKE_SET)
		{
			if ((mVarType == VT_UNKNOWN || mVarType == VT_DISPATCH) && !mUnknown)
			{
				// Allow this assignment only in the specific cases indicated above, to avoid ambiguity
				// about what to do with the old value.  This operation is specifically intended for use
				// with DllCall; i.e. DllCall(..., "ptr*", o := ComValue(13,0)) to wrap a returned ptr.
				// Assigning zero is permitted and there is no AddRef because the caller wants us to
				// Release the interface pointer automatically.
				if (!ParamIndexIsNumeric(0))
					_f_throw_type(_T("Number"), *aParam[0]);
				mUnknown = (IUnknown *)ParamIndexToInt64(0);
				return;
			}
		}
		else // GET
		{
			// Support passing VT_ARRAY, VT_BYREF or IUnknown to DllCall.
			if ((mVarType & (VT_ARRAY | VT_BYREF)) || mVarType == VT_UNKNOWN || mVarType == VT_DISPATCH)
				_o_return(mVal64);
		}
		_o_throw(ERR_INVALID_USAGE);
	}
}


FResult ComObject::SafeArray_Enum(optl<int> n, IObject *&aRetVal)
{
	HRESULT hr;
	ComArrayEnum *enm;
	if (SUCCEEDED(hr = ComArrayEnum::Begin(this, enm, n.value_or(2))))
		aRetVal = enm;
	return OK;
}


FResult ComObject::SafeArray_Clone(IObject *&aRetVal)
{
	HRESULT hr;
	SAFEARRAY *clone;
	if (SUCCEEDED(hr = SafeArrayCopy((SAFEARRAY*)mVal64, &clone)))
		aRetVal = new ComObject((__int64)clone, mVarType, F_OWNVALUE);
	return hr;
}


FResult ComObject::SafeArray_MaxIndex(optl<UINT> aDims, int &aRetVal)
{
	return SafeArrayGetUBound((SAFEARRAY*)mVal64, aDims.value_or(1), (LONG*)&aRetVal);
}


FResult ComObject::SafeArray_MinIndex(optl<UINT> aDims, int &aRetVal)
{
	return SafeArrayGetLBound((SAFEARRAY*)mVal64, aDims.value_or(1), (LONG*)&aRetVal);
}


FResult ComObject::SafeArray_Item(VariantParams &aParam, ExprTokenType *aNewValue, ResultToken *aResultToken)
{
	ASSERT(aNewValue || aResultToken);
	SAFEARRAY *psa = (SAFEARRAY*)mVal64;
	VARTYPE item_type = (mVarType & VT_TYPEMASK);

	UINT dims = SafeArrayGetDim(psa);
	LONG index[8];
	// Verify correct number of parameters/dimensions (maximum 8).
	if (dims > _countof(index) || dims != aParam.count)
		return FR_E_ARGS;
	// Build array of indices from parameters.
	for (UINT i = 0; i < dims; ++i)
	{
		if (!TokenIsNumeric(*aParam.value[i]))
			return FParamError(i, aParam.value[i], _T("Number"));
		index[i] = (LONG)TokenToInt64(*aParam.value[i]);
	}

	void *item;

	SafeArrayLock(psa);

	auto hr = SafeArrayPtrOfIndex(psa, index, &item);
	if (SUCCEEDED(hr))
	{
		if (aNewValue)
			hr = TokenToVarType(*aNewValue, item_type, item);
		else
			VarTypeToToken(item_type, item, *aResultToken);
	}

	SafeArrayUnlock(psa);

	return hr;
}


LPTSTR ComObject::Type()
{
	if ((mVarType == VT_DISPATCH || mVarType == VT_UNKNOWN) && mUnknown)
	{
		// Use COM class name if available.
		BSTR name = nullptr;
		if (ITypeInfo *ptinfo = GetClassTypeInfo(mUnknown))
		{
			ptinfo->GetDocumentation(MEMBERID_NIL, &name, NULL, NULL, NULL);
			ptinfo->Release();
		}
		if (name)
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


ComEnum::ComEnum(IEnumVARIANT *enm)
	: penum(enm)
{
	IServiceProvider *sp;
	if (SUCCEEDED(enm->QueryInterface<IServiceProvider>(&sp)))
	{
		IUnknown *unk;
		if (SUCCEEDED(sp->QueryService<IUnknown>(IID_IObjectComCompatible, &unk)))
		{
			cheat = true;
			unk->Release();
		}
		sp->Release();
	}
}


ResultType ComEnum::Next(Var *aVar0, Var *aVar1)
{
	VARIANT var[2] = {0};
	if (penum->Next(1 + (cheat && aVar1), var, NULL) == S_OK)
	{
		AssignVariant(*aVar0, var[0], false);
		if (aVar1)
		{
			if (cheat && aVar1)
				AssignVariant(*aVar1, var[1], false);
			else
				aVar1->Assign((__int64)var[0].vt);
		}
		return CONDITION_TRUE;
	}
	return CONDITION_FALSE;
}


HRESULT ComArrayEnum::Begin(ComObject *aArrayObject, ComArrayEnum *&aEnum, int aMode)
{
	HRESULT hr;
	SAFEARRAY *psa = aArrayObject->mArray;
	void *arrayData;
	long lbound, ubound;

	if (SafeArrayGetDim(psa) != 1)
		return E_NOTIMPL;
	
	if (   SUCCEEDED(hr = SafeArrayGetLBound(psa, 1, &lbound))
		&& SUCCEEDED(hr = SafeArrayGetUBound(psa, 1, &ubound))
		&& SUCCEEDED(hr = SafeArrayAccessData(psa, &arrayData))   )
	{
		VARTYPE arrayType = aArrayObject->mVarType & VT_TYPEMASK;
		UINT elemSize = SafeArrayGetElemsize(psa);
		aEnum = new ComArrayEnum(aArrayObject, arrayData, lbound, ubound, elemSize, arrayType, aMode >= 2);
		aArrayObject->AddRef(); // Keep obj alive until enumeration completes.
	}
	return hr;
}

ComArrayEnum::~ComArrayEnum()
{
	SafeArrayUnaccessData(mArrayObject->mArray); // Counter the lock done by ComArrayEnum::Begin.
	mArrayObject->Release();
}

ResultType ComArrayEnum::Next(Var *aVar1, Var *aVar2)
{
	++mOffset;
	if (mLBound + mOffset <= mUBound)
	{
		void *p = ((char *)mData) + mOffset * mElemSize;
		VARIANT var = {0};
		if (mType == VT_VARIANT)
		{
			// Make shallow copy of the VARIANT item.
			memcpy(&var, p, sizeof(VARIANT));
		}
		else
		{
			// Build VARIANT with shallow copy of the item.
			var.vt = mType;
			memcpy(&var.lVal, p, mElemSize);
		}
		if (mIndexMode)
		{
			if (aVar1)
				aVar1->Assign(mLBound + mOffset);
			aVar1 = aVar2;
		}
		if (aVar1)
			AssignVariant(*aVar1, var);
		return CONDITION_TRUE;
	}
	return CONDITION_FALSE;
}


STDMETHODIMP EnumComCompat::QueryInterface(REFIID riid, void **ppvObject)
{
	if (riid == IID_IUnknown || riid == IID_IEnumVARIANT)
		*ppvObject = static_cast<IEnumVARIANT*>(this);
	else if (riid == IID_IServiceProvider)
		*ppvObject = static_cast<IServiceProvider*>(this);
	else
		return E_NOTIMPL;
	AddRef();
	return S_OK;
}

STDMETHODIMP EnumComCompat::QueryService(REFGUID guidService, REFIID riid, void **ppvObject)
{
	// This is our secret handshake for enabling AutoHotkey enumeration behaviour.
	// Unlike calls to QueryInterface for this IID (due to the lack of registration etc.),
	// calls to this method should pass through the process/thread apartment boundary.
	if (guidService == IID_IObjectComCompatible && riid == IID_IUnknown)
	{
		*ppvObject = static_cast<IEnumVARIANT*>(this);
		AddRef();
		mCheat = true;
		return S_OK;
	}
	*ppvObject = nullptr;
	return E_NOTIMPL;
}

STDMETHODIMP_(ULONG) EnumComCompat::AddRef()
{
	return ++mRefCount;
}

STDMETHODIMP_(ULONG) EnumComCompat::Release()
{
	if (mRefCount)
		return --mRefCount;
	delete this;
	return 0;
}

STDMETHODIMP EnumComCompat::Next(ULONG celt, /*out*/ VARIANT *rgVar, /*out*/ ULONG *pCeltFetched)
{
	if (!celt)
		return E_INVALIDARG;

	int pc = min(celt, 1U + mCheat);
	VarRef *var[2] = { new VarRef, pc > 1 ? new VarRef : nullptr };
	ExprTokenType tparam[2], *param[] = { tparam, tparam + 1 };
	tparam[0].SetValue(var[0]);
	if (var[1])
		tparam[1].SetValue(var[1]);
	switch (CallEnumerator(mEnum, param, pc, false))
	{
	case CONDITION_TRUE:
		{
			ExprTokenType value;
			var[0]->ToTokenSkipAddRef(value);
			TokenToVariant(value, rgVar[0], FALSE);
			if (var[1])
			{
				var[1]->ToTokenSkipAddRef(value);
				TokenToVariant(value, rgVar[1], FALSE);
			}
			if (pCeltFetched)
				*pCeltFetched = pc - 1;
			break;
		}
		// else fall through.
	case INVOKE_NOT_HANDLED:
	case EARLY_EXIT:
	case FAIL:
		if (pCeltFetched)
			*pCeltFetched = 0;
		celt = -1;
		break;
	}

	delete var[0];
	if (var[1])
		delete var[1];

	return celt == pc ? S_OK : S_FALSE;
}

STDMETHODIMP EnumComCompat::Skip(ULONG celt)
{
	return E_NOTIMPL;
}

STDMETHODIMP EnumComCompat::Reset()
{
	return E_NOTIMPL;
}

STDMETHODIMP EnumComCompat::Clone(/*out*/ IEnumVARIANT **ppEnum)
{
	return E_NOTIMPL;
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
				pobj = new ComObject(pdisp);
			}
			else
			{
				pobj = new ComObject((__int64)punk, VT_UNKNOWN);
			}
			return pobj;
		}
	}
	return NULL;
}


STDMETHODIMP IObjectComCompatible::QueryInterface(REFIID riid, void **ppv)
{
	if (riid == IID_IDispatch || riid == IID_IUnknown || &riid == &IID_IObjectComCompatible)
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

static LPTSTR *sDispNameByIdMinus1;
static DISPID *sDispIdSortByName;
static DISPID sDispNameCount, sDispNameMax;

STDMETHODIMP IObjectComCompatible::GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId)
{
	HRESULT result_on_success = cNames == 1 ? S_OK : DISP_E_UNKNOWNNAME;
	for (UINT i = 0; i < cNames; ++i)
		rgDispId[i] = DISPID_UNKNOWN;

#ifdef UNICODE
	LPTSTR name = *rgszNames;
#else
	CStringCharFromWChar name_buf(*rgszNames);
	LPTSTR name = const_cast<LPTSTR>(name_buf.GetString());
#endif

	int left, right, mid, result;
	for (left = 0, right = sDispNameCount - 1; left <= right;)
	{
		mid = (left + right) / 2;
		// Comparison is case-sensitive so that the proper case of the name comes through for
		// meta-functions or new assignments.  Using different case will produce a different ID,
		// but the ID is ultimately mapped back to the name when the member is invoked anyway.
		result = _tcscmp(name, sDispNameByIdMinus1[sDispIdSortByName[mid] - 1]);
		if (result > 0)
			left = mid + 1;
		else if (result < 0)
			right = mid - 1;
		else // Match found.
		{
			*rgDispId = sDispIdSortByName[mid];
			return result_on_success;
		}
	}

	if (sDispNameMax == sDispNameCount)
	{
		int new_max = sDispNameMax ? sDispNameMax * 2 : 16;
		LPTSTR *new_names = (LPTSTR *)realloc(sDispNameByIdMinus1, new_max * sizeof(LPTSTR *));
		if (!new_names)
			return E_OUTOFMEMORY;
		DISPID *new_ids = (DISPID *)realloc(sDispIdSortByName, new_max * sizeof(DISPID));
		if (!new_ids)
		{
			free(new_names);
			return E_OUTOFMEMORY;
		}
		sDispNameByIdMinus1 = new_names;
		sDispIdSortByName = new_ids;
		sDispNameMax = new_max;
	}

	LPTSTR name_copy = SimpleHeap::Malloc(name);
	if (!name_copy)
		return E_OUTOFMEMORY;

	sDispNameByIdMinus1[sDispNameCount] = name_copy; // Put names in ID order; index = ID - 1.
	if (left < sDispNameCount)
		memmove(sDispIdSortByName + left + 1, sDispIdSortByName + left, (sDispNameCount - left) * sizeof(DISPID));
	sDispIdSortByName[left] = ++sDispNameCount; // Insert ID in order sorted by name, for binary search.  ID = index + 1, to avoid DISPID_VALUE.

	*rgDispId = sDispNameCount;
	return result_on_success;
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
	
	if (dispIdMember > 0 && dispIdMember <= sDispNameCount)
		name = sDispNameByIdMinus1[dispIdMember - 1];
	else if (dispIdMember == DISPID_VALUE)
		name = nullptr;
	else if (dispIdMember == DISPID_NEWENUM && (wFlags & (DISPATCH_METHOD | DISPATCH_PROPERTYGET)))
	{
		name = _T("__Enum");
		flags |= IF_NEWENUM;
		wFlags = (wFlags & ~DISPATCH_PROPERTYGET) | DISPATCH_METHOD;
	}
	else
		return DISP_E_MEMBERNOTFOUND;
	
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
			if (!pVarResult)
				break;
			if (dispIdMember == DISPID_NEWENUM && result_token.symbol == SYM_OBJECT)
			{
				pVarResult->vt = VT_UNKNOWN;
				pVarResult->punkVal = static_cast<IEnumVARIANT*>(new EnumComCompat(result_token.object));
				result_token.symbol = SYM_INTEGER; // Skip Release().
			}
			else
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
	if (aPage == 0 && aDepth > 0)
	{
		// For simplicity, assume aPageSize >= 2.
		aDebugger->WriteProperty("Value", ExprTokenType(mVal64));
		aDebugger->WriteProperty("VarType", ExprTokenType((__int64)mVarType));
	}
	aDebugger->EndProperty(rootCookie);
}

#endif
