#include "stdafx.h"
#include "globaldata.h"
#include "script.h"
#include "script_object.h"
#include "script_com.h"

#ifdef CONFIG_EXPERIMENTAL


// IID__IObject -- .NET's System.Object:
const IID IID__Object = {0x65074F7F, 0x63C0, 0x304E, 0xAF, 0x0A, 0xD5, 0x17, 0x41, 0xCB, 0x4A, 0x8D};


void BIF_ComObjCreate(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
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
	aResultToken.symbol = SYM_STRING;
	aResultToken.marker = _T("");
	ComError(hr);
}


void BIF_ComObjGet(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
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
	aResultToken.symbol = SYM_STRING;
	aResultToken.marker = _T("");
	ComError(hr);
}


void BIF_ComObjActive(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	if (!aParamCount) // ComObjMissing()
	{
		SafeSetTokenObject(aResultToken, new ComObject(DISP_E_PARAMNOTFOUND, VT_ERROR));
		return;
	}

	aResultToken.symbol = SYM_STRING;
	aResultToken.marker = _T("");

	ComObject *obj;

	if (TokenIsPureNumeric(*aParam[0]))
	{
		VARTYPE vt;
		__int64 llVal;
		USHORT flags = 0;

		if (aParamCount > 1)
		{
			// ComObj(vt, value [, flags])
			vt = (VARTYPE)TokenToInt64(*aParam[0]);
			llVal = TokenToInt64(*aParam[1]);
			if (aParamCount > 2)
				flags = (USHORT)TokenToInt64(*aParam[2]);
		}
		else
		{
			// ComObj(pdisp)
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
						// Replace caller-specified interface pointer with pdisp.  If caller
						// has requested we take responsibility for freeing it, do that now:
						if (flags & ComObject::F_OWNVALUE)
							punk->Release();
						flags |= ComObject::F_OWNVALUE; // Don't AddRef() below since we own this reference.
						llVal = (__int64)pdisp;
					}
					// Otherwise interpret it as IDispatch anyway, since caller has requested it and
					// there are known cases where it works (such as some CLR COM callable wrappers).
				}
				if ( !(flags & ComObject::F_OWNVALUE) )
					punk->AddRef(); // "Copy" caller's reference.
				// Otherwise caller (or above) indicated the object now owns this reference.
			}
			// Otherwise, NULL may have some meaning, so allow it.  If the
			// script tries to invoke the object, it'll get a warning then.
		}

		if (obj = new ComObject(llVal, vt, flags))
		{
			aResultToken.symbol = SYM_OBJECT;
			aResultToken.object = obj;
		}
		else if (vt == VT_DISPATCH || vt == VT_UNKNOWN)
			((IUnknown *)llVal)->Release();
	}
	else if (obj = dynamic_cast<ComObject *>(TokenToObject(*aParam[0])))
	{
		if (aParamCount > 1)
		{
			// For backward-compatibility:
			aResultToken.symbol = SYM_INTEGER;
			BIF_ComObjTypeOrValue(aResultToken, aParam, aParamCount);
		}
		else if (VT_DISPATCH == obj->mVarType)
		{
			aResultToken.symbol = SYM_INTEGER;
			aResultToken.value_int64 = (__int64) obj->mDispatch; // mDispatch vs mVal64 ensures we zero the high 32 bits in 32-bit builds.
			if (obj->mDispatch)
				obj->mDispatch->AddRef();
		}
	}
	else
	{
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
					if (obj = new ComObject(pdisp))
					{
						aResultToken.symbol = SYM_OBJECT;
						aResultToken.object = obj;
					}
					else
						pdisp->Release();
				}
				punk->Release();
				return;
			}
		}
		ComError(hr);
	}
}


void BIF_ComObjConnect(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	aResultToken.symbol = SYM_STRING;
	aResultToken.marker = _T("");

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
					if (SUCCEEDED(ptinfo->GetImplTypeFlags(j, &flags)) && flags == (IMPLTYPEFLAG_FDEFAULT | IMPLTYPEFLAG_FSOURCE))
					{
						HREFTYPE reftype;
						ITypeInfo *prinfo;
						if (SUCCEEDED(ptinfo->GetRefTypeOfImplType(j, &reftype))
							&& SUCCEEDED(ptinfo->GetRefTypeInfo(reftype, &prinfo)))
						{
							if (SUCCEEDED(prinfo->GetTypeAttr(&typeattr)))
							{
								obj->mEventSink = new ComEvent(obj, prinfo, typeattr->guid);
								prinfo->ReleaseTypeAttr(typeattr);
							}
							else
								prinfo->Release();
						}
						break;
					}
				}
				ptinfo->Release();
			}
		}

		if (obj->mEventSink)
		{
			obj->mEventSink->Connect(aParamCount>1 ? TokenToString(*aParam[1]) : NULL);
			return;
		}
	}
	ComError(-1);
}


void BIF_ComObjError(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	aResultToken.value_int64 = g_ComErrorNotify;
	if (aParamCount && TokenIsPureNumeric(*aParam[0]))
		g_ComErrorNotify = (TokenToInt64(*aParam[0]) != 0);
}


void BIF_ComObjTypeOrValue(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	ComObject *obj = dynamic_cast<ComObject *>(TokenToObject(*aParam[0]));
	if (!obj)
	{
		aResultToken.symbol = SYM_STRING;
		aResultToken.marker = _T("");
		return;
	}
	if (ctoupper(aResultToken.marker[6]) == 'V')
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
						StringFromGUID2(typeattr->guid, aResultToken.marker, MAX_NUMBER_SIZE);
#else
						WCHAR cnvbuf[MAX_NUMBER_SIZE];
						StringFromGUID2(typeattr->guid, cnvbuf, MAX_NUMBER_SIZE);
						CStringCharFromWChar cnvstring(cnvbuf);
						strncpy(aResultToken.marker, cnvstring.GetBuffer(), MAX_NUMBER_SIZE);
#endif
						ptinfo->ReleaseTypeAttr(typeattr);
					}
				}
				ptinfo->Release();
			}
		}
	}
}


void BIF_ComObjFlags(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	ComObject *obj = dynamic_cast<ComObject *>(TokenToObject(*aParam[0]));
	if (!obj)
	{
		aResultToken.symbol = SYM_STRING;
		aResultToken.marker = _T("");
		return;
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


void BIF_ComObjArray(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
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
	if (psa && !SafeSetTokenObject(aResultToken, new ComObject((__int64)psa, VT_ARRAY | vt, ComObject::F_OWNVALUE)))
		SafeArrayDestroy(psa);
}


void BIF_ComObjQuery(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
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
		aToken.mem_to_free = NULL; // Only needed for some callers.
		return false;
	}
}


void VariantToToken(VARIANT &aVar, ExprTokenType &aToken, bool aRetainVar = true)
{
	switch (aVar.vt)
	{
	case VT_BSTR:
		aToken.symbol = SYM_STRING;
		aToken.marker = _T("");		// Set defaults.
		aToken.mem_to_free = NULL;	//
		size_t len;
		if (len = SysStringLen(aVar.bstrVal))
		{
#ifdef UNICODE
			if (aRetainVar)
			{
				// It's safe to pass back the actual BSTR from aVar in this case.
				aToken.marker = aVar.bstrVal;
			}
			// Allocate some memory to pass back to caller:
			else if (aToken.mem_to_free = tmalloc(len + 1))
			{
				aToken.marker = aToken.mem_to_free;
				aToken.marker_length = len;
				tmemcpy(aToken.marker, aVar.bstrVal, len + 1); // +1 for null-terminator
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
		aToken.mem_to_free = NULL;
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
	ExprTokenType token;
	VariantToToken(aVar, token, aRetainVar);
	switch (token.symbol)
	{
	case SYM_STRING:  // VT_BSTR was handled above, but VariantToToken coerces unhandled types to strings.
		if (token.mem_to_free)
			aArg.AcceptNewMem(token.mem_to_free, token.marker_length);
		else
			aArg.Assign();
		break;
	case SYM_OBJECT:
		aArg.AssignSkipAddRef(token.object); // Let aArg take responsibility for it.
		break;
	default:
		aArg.Assign(token);
		break;
	}
}


void TokenToVariant(ExprTokenType &aToken, VARIANT &aVar)
{
	if (aToken.symbol == SYM_VAR)
		aToken.var->TokenToContents(aToken);

	switch(aToken.symbol)
	{
	case SYM_OPERAND:
		if (aToken.buf)
		{
			__int64 val = *(__int64 *)aToken.buf;
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
			break;
		}
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
	BSTR memberName;
	UINT nNames;
	if (FAILED(mTypeInfo->GetNames(dispIdMember, &memberName, 1, &nNames)))
		return DISP_E_MEMBERNOTFOUND;

	TCHAR funcName[256];
	sntprintf(funcName, _countof(funcName), _T("%s%ws"), mPrefix, memberName);
	SysFreeString(memberName);

	Func *func = g_script.FindFunc(funcName);
	if (!func)
		return DISP_E_MEMBERNOTFOUND;

	UINT cArgs = pDispParams->cArgs;

	if (cArgs >= MAX_FUNCTION_PARAMS // >= vs > to allow for 'this' object param.
		|| !mObject) // mObject == NULL should be next to impossible since it is only set NULL after calling Unadvise(), in which case there shouldn't be anyone left to call this->Invoke().  Check it anyway since it might be difficult to debug, depending on what we're connected to.
		return DISP_E_MEMBERNOTFOUND;

	ExprTokenType param_token[MAX_FUNCTION_PARAMS];
	ExprTokenType *param[MAX_FUNCTION_PARAMS];

	for (UINT i = 0; i < cArgs; ++i)
	{
		VARIANTARG *pvar = &pDispParams->rgvarg[cArgs-1-i];
		while (pvar->vt == (VT_BYREF | VT_VARIANT))
			pvar = pvar->pvarVal;
		VariantToToken(*pvar, param_token[i]);
		param[i] = &param_token[i];
	}

	// Pass our object last for either of the following cases:
	//	a) Our caller doesn't include its IDispatch interface pointer in the parameter list.
	//	b) The script needs a reference to the original wrapper object; i.e. mObject.
	param_token[cArgs].symbol = SYM_OBJECT;
	param_token[cArgs].object = mObject;
	param[cArgs] = &param_token[cArgs];

	FuncCallData func_call;
	ExprTokenType result_token;
	ResultType result;
	HRESULT result_to_return;

	// Call the function.
	if (func->Call(func_call, result, result_token, param, cArgs + 1))
	{
		if (pVarResult)
			TokenToVariant(result_token, *pVarResult);
		if (result_token.symbol == SYM_OBJECT)
			result_token.object->Release();
		result_to_return = S_OK;
	}
	else // above failed or exited, so result_token should be ignored.
		result_to_return = DISP_E_MEMBERNOTFOUND; // For consistency.  Probably doesn't matter whether we return this or S_OK.

	// Clean up:
	for (UINT i = 0; i < cArgs; ++i)
	{
		// Release COM wrapper objects:
		if (param_token[i].symbol == SYM_OBJECT)
			param_token[i].object->Release();
		// Free any temporary memory used to hold strings; see VariantToToken().
		else if (param_token[i].symbol == SYM_STRING && param_token[i].mem_to_free)
			free(param_token[i].mem_to_free);
	}

	return S_OK;
}

void ComEvent::Connect(LPTSTR pfx)
{
	HRESULT hr;
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
				if (!mCookie)
				{
					_tcscpy(mPrefix, pfx);
					hr = pconn->Advise(this, &mCookie);
				}
			}
			else
			{
				if (mCookie)
				{
					hr = pconn->Unadvise(mCookie);
					if (SUCCEEDED(hr))
						mCookie = 0;
				}
			}
			pconn->Release();
		}
		pcpc->Release();
	}
	if (FAILED(hr))
		ComError(hr);
}

ResultType STDMETHODCALLTYPE ComObject::Invoke(ExprTokenType &aResultToken, ExprTokenType &aThisToken, int aFlags, ExprTokenType *aParam[], int aParamCount)
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

#ifdef UNICODE
	HRESULT	hr = mDispatch->GetIDsOfNames(IID_NULL, &aName, 1, LOCALE_USER_DEFAULT, &dispid);
#else
	CStringWCharFromChar cnvbuf(aName);
	LPOLESTR cnvbuf_ptr = (LPOLESTR)(LPCWSTR) cnvbuf;
	HRESULT	hr = mDispatch->GetIDsOfNames(IID_NULL, &cnvbuf_ptr, 1, LOCALE_USER_DEFAULT, &dispid);
#endif
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
	{	// Allow chaining, e.g. obj2.prop := obj1.prop := val.
		ExprTokenType &rvalue = *aParam[aParamCount];
		aResultToken.symbol = (rvalue.symbol == SYM_OPERAND) ? SYM_STRING : rvalue.symbol;
		aResultToken.value_int64 = rvalue.value_int64;
		if (rvalue.symbol == SYM_OBJECT)
			rvalue.object->AddRef();
	}
	else
	{
		VariantToToken(varResult, aResultToken, false);
	}

	g->LastError = hr;
	return	OK;
}

ResultType ComObject::SafeArrayInvoke(ExprTokenType &aResultToken, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	HRESULT hr = S_OK;
	SAFEARRAY *psa = (SAFEARRAY*)mVal64;
	VARTYPE item_type = (mVarType & VT_TYPEMASK);

	if (IS_INVOKE_CALL)
	{
		LPTSTR name = TokenToString(*aParam[0]);
		if (*name == '_')
			++name;
		LONG retval;
		if (!_tcsicmp(name, _T("NewEnum")))
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
		if (!TokenIsPureNumeric(*aParam[i]))
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
			switch (rvalue.symbol)
			{
			case SYM_OPERAND:
				if (rvalue.buf)
				{
					aResultToken.symbol = SYM_INTEGER;
					aResultToken.value_int64 = *(__int64 *)rvalue.buf;
					break;
				}
				// FALL THROUGH to next case:
			case SYM_STRING:
				aResultToken.symbol = SYM_STRING;
				aResultToken.marker = rvalue.marker;
				break;
			case SYM_OBJECT:
				aResultToken.symbol = SYM_OBJECT;
				aResultToken.object = rvalue.object;
				aResultToken.object->AddRef();
				break;
			case SYM_INTEGER:
			case SYM_FLOAT:
				aResultToken.symbol = rvalue.symbol;
				aResultToken.value_int64 = rvalue.value_int64;
				break;
			}
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


#endif
