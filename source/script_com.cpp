#include "stdafx.h"
#include "globaldata.h"
#include "script.h"
#include "script_object.h"
#include "script_com.h"


// IID__IObject -- .NET's System.Object:
const IID IID__Object = {0x65074F7F, 0x63C0, 0x304E, 0xAF, 0x0A, 0xD5, 0x17, 0x41, 0xCB, 0x4A, 0x8D};


BIF_DECL(BIF_ComObjCreate)
{
	HRESULT hr;
	CLSID clsid, iid;
	for (;;)
	{
		hr = CLSIDFromString(CStringWCharFromTCharIfNeeded(TokenToString(*aParam[0])), &clsid);
		if (FAILED(hr)) break;

		if (aParamCount > 1)
		{
			hr = CLSIDFromString(CStringWCharFromTCharIfNeeded(TokenToString(*aParam[1])), &iid);
			if (FAILED(hr)) break;
			
			IUnknown *punk;
			hr = CoCreateInstance(clsid, NULL, CLSCTX_SERVER, iid, (void **)&punk);
			if (FAILED(hr)) break;

			// Return interface pointer, as requested.
			aResultToken.symbol = SYM_INTEGER;
			aResultToken.value_int64 = (__int64)punk;
		}
		else
		{
			IDispatch *pdisp;
			hr = CoCreateInstance(clsid, NULL, CLSCTX_SERVER, IID_IDispatch, (void **)&pdisp);
			if (FAILED(hr)) break;
			
			// Return dispatchable object.
			if ( !(aResultToken.object = new ComObject(pdisp)) )
				break;
			aResultToken.symbol = SYM_OBJECT;
		}
		return;
	}
	_f_set_retval_p(_T(""), 0);
	ComError(hr);
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
		pdisp->Release();
	}
	_f_set_retval_p(_T(""), 0);
	ComError(hr);
}


BIF_DECL(BIF_ComObject)
{
	if (!TokenIsNumeric(*aParam[0]))
	{
		aResultToken.Error(ERR_PARAM1_INVALID);
		return;
	}
		
	VARTYPE vt;
	__int64 llVal;
	USHORT flags = 0;

	if (aParamCount > 1)
	{
		if (!TokenIsNumeric(*aParam[1]))
		{
			aResultToken.Error(ERR_PARAM2_INVALID);
			return;
		}
		// ComObject(vt, value [, flags])
		vt = (VARTYPE)TokenToInt64(*aParam[0]);
		llVal = TokenToInt64(*aParam[1]);
		if (aParamCount > 2)
			flags = (USHORT)TokenToInt64(*aParam[2]);
	}
	else
	{
		// ComObject(pdisp)
		vt = VT_DISPATCH;
		llVal = TokenToInt64(*aParam[0]);
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
		aResultToken.Error(ERR_OUTOFMEM);
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
			if (SUCCEEDED(punk->QueryInterface(IID_IDispatch, (void **)&pdisp)))
			{
				if (ComObject *obj = new ComObject(pdisp))
				{
					aResultToken.symbol = SYM_OBJECT;
					aResultToken.object = obj;
				}
				else
				{
					hr = E_OUTOFMEMORY;
					pdisp->Release();
				}
			}
			punk->Release();
		}
	}
	if (FAILED(hr))
		ComError(hr);
}


BIF_DECL(BIF_ComObjConnect)
{
	_f_set_retval_p(_T(""), 0);

	if (ComObject *obj = dynamic_cast<ComObject *>(TokenToObject(*aParam[0])))
	{
		if (!obj->mEventSink && VT_DISPATCH == obj->mVarType && obj->mDispatch)
		{
			ITypeLib *ptlib;
			ITypeInfo *ptinfo;
			TYPEATTR *typeattr;
			IID iid;
			BOOL found = false;

			if (SUCCEEDED(obj->mDispatch->GetTypeInfo(0, LOCALE_USER_DEFAULT, &ptinfo)))
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

			if (found)
			{
				found = false;
				UINT ctinfo = ptlib->GetTypeInfoCount();
				for (UINT i = 0; i < ctinfo; i++)
				{
					TYPEKIND typekind;
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
							HREFTYPE reftype;
							ITypeInfo *prinfo;
							if (SUCCEEDED(ptinfo->GetRefTypeOfImplType(j, &reftype))
								&& SUCCEEDED(ptinfo->GetRefTypeInfo(reftype, &prinfo)))
							{
								if (SUCCEEDED(prinfo->GetTypeAttr(&typeattr)))
								{
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
			}

			if (found)
			{
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
		}

		if (obj->mEventSink)
		{
			if (aParamCount < 2)
				obj->mEventSink->Connect(); // Disconnect.
			else
				obj->mEventSink->Connect(TokenToString(*aParam[1]), TokenToObject(*aParam[1]));
			return;
		}

		ComError(E_NOINTERFACE);
	}
	else
		ComError(-1); // "No COM object"
}


BIF_DECL(BIF_ComObjError)
{
	aResultToken.value_int64 = g_ComErrorNotify;
	if (aParamCount && TokenIsNumeric(*aParam[0]))
		g_ComErrorNotify = (TokenToInt64(*aParam[0]) != 0);
}


BIF_DECL(BIF_ComObjTypeOrValue)
{
	ComObject *obj = dynamic_cast<ComObject *>(TokenToObject(*aParam[0]));
	if (!obj)
		_f_return_empty;
	if (_f_callee_id == FID_ComObjValue)
	{
		aResultToken.value_int64 = obj->mVal64;
	}
	else
	{
		if (aParamCount < 2)
		{
			aResultToken.value_int64 = obj->mVarType;
		}
		else
		{
			aResultToken.symbol = SYM_STRING; // for all code paths below
			aResultToken.marker = _T(""); // in case of error
			aResultToken.marker_length = 0;

			ITypeInfo *ptinfo;
			if (VT_DISPATCH == obj->mVarType && obj->mDispatch
				&& SUCCEEDED(obj->mDispatch->GetTypeInfo(0, LOCALE_USER_DEFAULT, &ptinfo)))
			{
				LPTSTR requested_info = TokenToString(*aParam[1]);
				if (!_tcsicmp(requested_info, _T("name")))
				{
					BSTR name;
					if (SUCCEEDED(ptinfo->GetDocumentation(MEMBERID_NIL, &name, NULL, NULL, NULL)))
					{
						TokenSetResult(aResultToken, CStringTCharFromWCharIfNeeded(name), SysStringLen(name));
						SysFreeString(name);
					}
				}
				else if (!_tcsicmp(requested_info, _T("iid")))
				{
					TYPEATTR *typeattr;
					if (SUCCEEDED(ptinfo->GetTypeAttr(&typeattr)))
					{
						aResultToken.marker = aResultToken.buf;
#ifdef UNICODE
						aResultToken.marker_length = StringFromGUID2(typeattr->guid, aResultToken.marker, MAX_NUMBER_SIZE);
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
}


BIF_DECL(BIF_ComObjFlags)
{
	ComObject *obj = dynamic_cast<ComObject *>(TokenToObject(*aParam[0]));
	if (!obj)
		_f_return_empty;
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


BIF_DECL(BIF_ComObjArray)
{
	VARTYPE vt = (VARTYPE)TokenToInt64(*aParam[0]);
	SAFEARRAYBOUND bound[8]; // Same limit as ComObject::SafeArrayInvoke().
	int dims = aParamCount - 1;
	ASSERT(dims <= _countof(bound)); // Prior validation should ensure aParamCount-1 never exceeds 8.
	for (int i = 0; i < dims; ++i)
	{
		bound[i].cElements = (ULONG)TokenToInt64(*aParam[i + 1]);
		bound[i].lLbound = 0;
	}
	SAFEARRAY *psa = SafeArrayCreate(vt, dims, bound);
	if (!SafeSetTokenObject(aResultToken, psa ? new ComObject((__int64)psa, VT_ARRAY | vt, ComObject::F_OWNVALUE) : NULL) && psa)
		SafeArrayDestroy(psa);
}


BIF_DECL(BIF_ComObjQuery)
{
	IUnknown *punk = NULL;
	ComObject *obj;
	HRESULT hr;
	
	aResultToken.value_int64 = 0; // Set default; on 32-bit builds, only the low 32 bits may be set below.

	if (obj = dynamic_cast<ComObject *>(TokenToObject(*aParam[0])))
	{
		// We were passed a ComObject, but does it contain an interface pointer?
		if (obj->mVarType == VT_UNKNOWN || obj->mVarType == VT_DISPATCH)
			punk = obj->mUnknown;
	}
	if (!punk)
	{
		// Since it wasn't a valid ComObject, it should be a raw interface pointer.
		punk = (IUnknown *)TokenToInt64(*aParam[0]);
		if (punk < (IUnknown *)65536) // Error-detection: the first 64KB of address space is always invalid.
		{
			g->LastError = E_INVALIDARG; // For consistency.
			ComError(-1);
			return;
		}
	}

	if (aParamCount > 2) // QueryService(obj, SID, IID)
	{
		GUID sid, iid;
		if (   SUCCEEDED(hr = CLSIDFromString(CStringWCharFromTCharIfNeeded(TokenToString(*aParam[1])), &sid))
			&& SUCCEEDED(hr = CLSIDFromString(CStringWCharFromTCharIfNeeded(TokenToString(*aParam[2])), &iid))   )
		{
			IServiceProvider *pprov;
			if (SUCCEEDED(hr = punk->QueryInterface<IServiceProvider>(&pprov)))
			{
				hr = pprov->QueryService(sid, iid, (void **)&aResultToken.value_int64);
			}
		}
	}
	else // QueryInterface(obj, IID)
	{
		GUID iid;
		if (SUCCEEDED(hr = CLSIDFromString(CStringWCharFromTCharIfNeeded(TokenToString(*aParam[1])), &iid)))
		{
			hr = punk->QueryInterface(iid, (void **)&aResultToken.value_int64);
		}
	}

	g->LastError = hr;
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


void TokenToVariant(ExprTokenType &aToken, VARIANT &aVar)
{
	if (aToken.symbol == SYM_VAR)
		aToken.var->ToTokenSkipAddRef(aToken);

	switch(aToken.symbol)
	{
	case SYM_STRING:
		aVar.vt = VT_BSTR;
		aVar.bstrVal = SysAllocString(CStringWCharFromTCharIfNeeded(aToken.marker));
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
			obj->ToVariant(aVar);
		else
			aVar.vt = VT_EMPTY;
		break;
	case SYM_MISSING:
		aVar.vt = VT_ERROR;
		aVar.scode = DISP_E_PARAMNOTFOUND;
		break;
	}
}


bool g_ComErrorNotify = true;

void ComError(HRESULT hr, LPTSTR name, EXCEPINFO* pei)
{
	if (hr != DISP_E_EXCEPTION)
		pei = NULL;

	if (g_ComErrorNotify)
	{
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
			int size = _stprintf(buf, _T("0x%08X - "), hr);
			size += FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS, NULL, hr, 0, buf + size, _countof(buf) - size, NULL);
			if (buf[size-1] == '\n')
				buf[--size] = '\0';
				// probably also has '\r', but doesn't seem necessary to remove it.
			if (pei)
				_vsntprintf(buf + size, _countof(buf) - size, _T("\nSource:\t\t%ws\nDescription:\t%ws\nHelpFile:\t\t%ws\nHelpContext:\t%d"), (va_list) &pei->bstrSource);
			error_text = buf;
		}

		g_script.mCurrLine->LineError(error_text, EARLY_EXIT, name);
	}

	if (pei)
	{
		SysFreeString(pei->bstrSource);
		SysFreeString(pei->bstrDescription);
		SysFreeString(pei->bstrHelpFile);
	}
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

STDMETHODIMP_(ULONG) ComEvent::AddRef()
{
	return ++mRefCount;
}

STDMETHODIMP_(ULONG) ComEvent::Release()
{
	if (--mRefCount)
		return mRefCount;
	delete this;
	return 0;
}

STDMETHODIMP ComEvent::GetTypeInfoCount(UINT *pctinfo)
{
	*pctinfo = 0;
	return S_OK;
}

STDMETHODIMP ComEvent::GetTypeInfo(UINT itinfo, LCID lcid, ITypeInfo **pptinfo)
{
	*pptinfo = NULL;
	return E_NOTIMPL;
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

	// Copy method name into our buffer, applying prefix and converting if necessary.
	TCHAR funcName[256];
	int funcNameLen = sntprintf(funcName, _countof(funcName), _T("%s%ws"), mPrefix, memberName);
	SysFreeString(memberName);

	UINT cArgs = pDispParams->cArgs;
	const UINT ADDITIONAL_PARAMS = 2; // mObject (passed by us) and mAhkObject (passed by Object::Invoke).
	const UINT MAX_COM_PARAMS = MAX_FUNCTION_PARAMS - ADDITIONAL_PARAMS;
	if (cArgs > MAX_COM_PARAMS) // Probably won't happen in any real-world script.
		cArgs = MAX_COM_PARAMS; // Just omit the rest of the params.

	ResultToken param_token[MAX_FUNCTION_PARAMS];
	ExprTokenType *param[MAX_FUNCTION_PARAMS];
	
	for (UINT i = 1; i <= cArgs; ++i)
	{
		VARIANTARG *pvar = &pDispParams->rgvarg[cArgs-i];
		while (pvar->vt == (VT_BYREF | VT_VARIANT))
			pvar = pvar->pvarVal;
		VariantToToken(*pvar, param_token[i]);
		param[i] = &param_token[i];
	}
	
	// Pass our object last for either of the following cases:
	//	a) Our caller doesn't include its IDispatch interface pointer in the parameter list.
	//	b) The script needs a reference to the original wrapper object; i.e. mObject.
	param_token[cArgs + 1].symbol = SYM_OBJECT;
	param_token[cArgs + 1].object = mObject;
	param[cArgs + 1] = &param_token[cArgs + 1];

	HRESULT result_to_return;
	FuncResult result_token;

	if (mAhkObject)
	{
		ExprTokenType this_token;
		this_token.symbol = SYM_OBJECT;
		this_token.object = mAhkObject;

		param_token[0].symbol = SYM_STRING;
		param_token[0].marker = funcName;
		param_token[0].marker_length = funcNameLen;
		param[0] = &param_token[0];

		// Call method of mAhkObject by name.
		if (mAhkObject->Invoke(result_token, this_token, IT_CALL, param, cArgs + 2) != INVOKE_NOT_HANDLED)
			result_to_return = S_OK;
		else
			result_to_return = DISP_E_MEMBERNOTFOUND;
	}
	else
	{
		// Call function by name (= prefix . method_name).
		Func *func = g_script.FindFunc(funcName, funcNameLen);

		// Call the function.
		if (func && func->Call(result_token, param + 1, cArgs + 1))
			result_to_return = S_OK;
		else
			result_to_return = DISP_E_MEMBERNOTFOUND; // Indicate result_token should not be used below.
	}

	if (pVarResult && result_to_return == S_OK)
		TokenToVariant(result_token, *pVarResult);

	// Clean up:
	result_token.Free();
	for (UINT i = 1; i <= cArgs; ++i)
		param_token[i].Free();

	return S_OK;
}

void ComEvent::Connect(LPTSTR pfx, IObject *ahkObject)
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
	}
	else
		ComError(hr);
}

ResultType STDMETHODCALLTYPE ComObject::Invoke(ResultToken &aResultToken, ExprTokenType &aThisToken, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	if (aParamCount < (IS_INVOKE_SET ? 2 : 1))
	{
		// Something like x[] or x[]:=y -- reserved for possible future use.  However, it could
		// be x[prms*] where prms is an empty array or not an array at all, so raise an error:
		ComError(g->LastError = DISP_E_BADPARAMCOUNT);
		return OK;
	}

	if (mVarType != VT_DISPATCH || !mDispatch)
	{
		if (mVarType & VT_ARRAY)
			return SafeArrayInvoke(aResultToken, aFlags, aParam, aParamCount);
		// Otherwise: this object can't be invoked.
		g->LastError = DISP_E_BADVARTYPE; // Seems more informative than -1.
		ComError(-1);
		return OK;
	}

	static DISPID dispidParam = DISPID_PROPERTYPUT;
	DISPPARAMS dispparams = {NULL, NULL, 0, 0};
	VARIANTARG *rgvarg;
	EXCEPINFO excepinfo = {0};
	VARIANT varResult = {0};
	DISPID dispid;

	LPTSTR aName = TokenToString(*aParam[0], aResultToken.buf);

	if (--aParamCount)
	{
		rgvarg = (VARIANTARG *)_alloca(sizeof(VARIANTARG) * aParamCount);

		for (int i = 0; i < aParamCount; i++)
		{
			TokenToVariant(*aParam[aParamCount-i], rgvarg[i]);
		}

		dispparams.rgvarg = rgvarg;
		dispparams.cArgs = aParamCount;
		if (IS_INVOKE_SET)
		{
			dispparams.rgdispidNamedArgs = &dispidParam;
			dispparams.cNamedArgs = 1;
		}
	}

	HRESULT	hr;
	if (aFlags & IF_NEWENUM)
	{
		hr = S_OK;
		dispid = DISPID_NEWENUM;
	}
	else
	{
#ifdef UNICODE
		hr = mDispatch->GetIDsOfNames(IID_NULL, &aName, 1, LOCALE_USER_DEFAULT, &dispid);
#else
		CStringWCharFromChar cnvbuf(aName);
		LPOLESTR cnvbuf_ptr = (LPOLESTR)(LPCWSTR)cnvbuf;
		hr = mDispatch->GetIDsOfNames(IID_NULL, &cnvbuf_ptr, 1, LOCALE_USER_DEFAULT, &dispid);
#endif
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
		// If this param is SYM_OBJECT, it is either an unsupported object (in which case rgvarg[i] is empty)
		// or a ComObject, in which case rgvarg[i] is a shallow copy and calling VariantClear would free the
		// caller's data prematurely. Even VT_DISPATCH should not be cleared since TokenToVariant didn't AddRef.
		if (aParam[aParamCount-i]->symbol != SYM_OBJECT)
			VariantClear(&rgvarg[i]);
	}

	if	(FAILED(hr))
	{
		ComError(hr, aName, &excepinfo);
	}
	else if	(IS_INVOKE_SET)
	{
		// Allow chaining, e.g. obj2.prop := obj1.prop := val.
		aResultToken.CopyValueFrom(*aParam[aParamCount]); // Can't be SYM_VAR at this stage because TokenToVariant() would have converted it to its contents.
		if (aResultToken.symbol == SYM_OBJECT)
			aResultToken.object->AddRef();
	}
	else
	{
		VariantToToken(varResult, aResultToken, false);
	}

	g->LastError = hr;
	return	OK;
}

ResultType ComObject::SafeArrayInvoke(ResultToken &aResultToken, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	HRESULT hr = S_OK;
	SAFEARRAY *psa = (SAFEARRAY*)mVal64;
	VARTYPE item_type = (mVarType & VT_TYPEMASK);

	if (IS_INVOKE_CALL)
	{
		LPTSTR name = TokenToString(*aParam[0]);
		LONG retval;
		if (!_tcsicmp(name, _T("_NewEnum")))
		{
			ComArrayEnum *enm;
			if (SUCCEEDED(hr = ComArrayEnum::Begin(this, enm)))
			{
				aResultToken.symbol = SYM_OBJECT;
				aResultToken.object = enm;
			}
		}
		else if (!_tcsicmp(name, _T("Clone")))
		{
			SAFEARRAY *clone;
			if (SUCCEEDED(hr = SafeArrayCopy(psa, &clone)))
				if (!SafeSetTokenObject(aResultToken, new ComObject((__int64)clone, mVarType, F_OWNVALUE)))
					SafeArrayDestroy(clone);
		}
		else
		{
			if (!_tcsicmp(name, _T("MaxIndex")))
				hr = SafeArrayGetUBound(psa, aParamCount > 1 ? (UINT)TokenToInt64(*aParam[1]) : 1, &retval);
			else if (!_tcsicmp(name, _T("MinIndex")))
				hr = SafeArrayGetLBound(psa, aParamCount > 1 ? (UINT)TokenToInt64(*aParam[1]) : 1, &retval);
			else
				hr = DISP_E_UNKNOWNNAME; // Seems slightly better than ignoring the call.
			if (SUCCEEDED(hr))
			{
				aResultToken.symbol = SYM_INTEGER;
				aResultToken.value_int64 = retval;
			}
		}
		g->LastError = hr;
		if (FAILED(hr))
			ComError(hr);
		return OK;
	}

	UINT dims = SafeArrayGetDim(psa);
	LONG index[8];
	// Verify correct number of parameters/dimensions (maximum 8).
	if (dims > _countof(index) || dims != (IS_INVOKE_SET ? aParamCount - 1 : aParamCount))
	{
		g->LastError = DISP_E_BADPARAMCOUNT;
		return OK;
	}
	// Build array of indices from parameters.
	for (UINT i = 0; i < dims; ++i)
	{
		if (!TokenIsNumeric(*aParam[i]))
		{
			g->LastError = E_INVALIDARG;
			return OK;
		}
		index[i] = (LONG)TokenToInt64(*aParam[i]);
	}

	VARIANT var = {0};
	void *item;

	SafeArrayLock(psa);

	hr = SafeArrayPtrOfIndex(psa, index, &item);
	if (SUCCEEDED(hr))
	{
		if (IS_INVOKE_GET)
		{
			if (item_type == VT_VARIANT)
			{
				// Make shallow copy of the VARIANT item.
				memcpy(&var, item, sizeof(VARIANT));
			}
			else
			{
				// Build VARIANT with shallow copy of the item.
				var.vt = item_type;
				memcpy(&var.lVal, item, SafeArrayGetElemsize(psa));
			}
			// Copy value into result token.
			VariantToToken(var, aResultToken);
		}
		else // SET
		{
			ExprTokenType &rvalue = *aParam[dims];
			TokenToVariant(rvalue, var);
			if ((var.vt == VT_DISPATCH || var.vt == VT_UNKNOWN) && var.punkVal)
				var.punkVal->AddRef();
			// Otherwise: it could be VT_BSTR, in which case TokenToVariant created a new BSTR which we want to
			// put directly it into the array rather than copying it, since it would only be freed later anyway.
			if (item_type == VT_VARIANT)
			{
				if ((var.vt & ~VT_TYPEMASK) == VT_ARRAY // Implies rvalue contains a ComObject.
					&& (((ComObject *)rvalue.object)->mFlags & F_OWNVALUE))
				{
					// Copy array since both sides will call Destroy().
					hr = VariantCopy((VARIANTARG *)item, &var);
				}
				else
				{
					// Free existing value.
					VariantClear((VARIANTARG *)item);
					// Write new value (shallow copy).
					memcpy(item, &var, sizeof(VARIANT));
				}
			}
			else
			{
				if (var.vt != item_type)
				{
					// Attempt to coerce var to the correct type:
					hr = VariantChangeType(&var, &var, 0, item_type);
					if (FAILED(hr))
					{
						VariantClear(&var);
						goto unlock_and_return;
					}
				}
				// Free existing value.
				if (item_type == VT_UNKNOWN || item_type == VT_DISPATCH)
				{
					IUnknown *punk = ((IUnknown **)item)[0];
					if (punk)
						punk->Release();
				}
				else if (item_type == VT_BSTR)
				{
					SysFreeString(*((BSTR *)item));
				}
				// Write new value (shallow copy).
				memcpy(item, &var.lVal, SafeArrayGetElemsize(psa));
			}
			// Allow chaining - using the following rather than VariantToToken allows any wrapper objects to
			// be passed along rather than being coerced to a simple value or a new wrapper being created.
			aResultToken.CopyValueFrom(rvalue);
			if (aResultToken.symbol == SYM_OBJECT)
				aResultToken.object->AddRef();
		}
	}

unlock_and_return:
	SafeArrayUnlock(psa);

	g->LastError = hr;
	if (FAILED(hr))
		ComError(hr);
	return OK;
}


int ComEnum::Next(Var *aOutput, Var *aOutputType)
{
	VARIANT varResult = {0};
	if (penum->Next(1, &varResult, NULL) == S_OK)
	{
		if (aOutputType)
			aOutputType->Assign((__int64)varResult.vt);
		if (aOutput)
			AssignVariant(*aOutput, varResult, false);
		return true;
	}
	return	false;
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
		arrayEnd = arrayData + (ubound - lbound) * elemSize;
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

int ComArrayEnum::Next(Var *aOutput, Var *aOutputType)
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
		// Copy value into var.
		AssignVariant(*aOutput, var);
		if (aOutputType)
			aOutputType->Assign(var.vt);
		return true;
	}
	return false;
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


#ifdef CONFIG_DEBUGGER

void WriteComObjType(IDebugProperties *aDebugger, ComObject *aObject, LPCSTR aName, LPTSTR aWhichType)
{
	TCHAR buf[_f_retval_buf_size];
	ResultToken resultToken;
	static Func *ComObjType = NULL;
	if (!ComObjType)
		ComObjType = g_script.FindFunc(_T("ComObjType")); // FIXME: Probably better to split ComObjType and ComObjValue.
	resultToken.func = ComObjType;
	resultToken.symbol = SYM_INTEGER;
	resultToken.marker_length = -1;
	resultToken.mem_to_free = NULL;
	resultToken.buf = buf;
	ExprTokenType paramToken[] = { aObject, aWhichType };
	ExprTokenType *param[] = { &paramToken[0], &paramToken[1] };
	BIF_ComObjTypeOrValue(resultToken, param, 2);
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
		
		aDebugger->WriteProperty("Value", mVal64);
		aDebugger->WriteProperty("VarType", mVarType);

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
				aDebugger->WriteProperty("Object", mEventSink->mAhkObject);
			else
				aDebugger->WriteProperty("Prefix", mEventSink->mPrefix);
			
			OLECHAR buf[40];
			if (!StringFromGUID2(mEventSink->mIID, buf, _countof(buf)))
				*buf = 0;
			aDebugger->WriteProperty("IID", (LPTSTR)(LPCTSTR)CStringTCharFromWCharIfNeeded(buf));
			
			aDebugger->EndProperty(sinkCookie);
		}
	}
	aDebugger->EndProperty(rootCookie);
}

#endif
