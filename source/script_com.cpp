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
	aResultToken.symbol = SYM_STRING;
	aResultToken.marker = _T("");
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
	aResultToken.symbol = SYM_STRING;
	aResultToken.marker = _T("");
	ComError(hr);
}


BIF_DECL(BIF_ComObjActive)
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
			aResultToken.marker = _T("ComObjType");
			BIF_ComObjTypeOrValue(aResult, aResultToken, aParam, aParamCount);
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


BIF_DECL(BIF_ComObjConnect)
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
	if (aParamCount && TokenIsPureNumeric(*aParam[0]))
		g_ComErrorNotify = (TokenToInt64(*aParam[0]) != 0);
}


BIF_DECL(BIF_ComObjTypeOrValue)
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


BIF_DECL(BIF_ComObjFlags)
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


void TokenToVariant(ExprTokenType &aToken, VARIANT &aVar, BOOL aVarIsArg)
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
		}
		else
		{
			aVar.vt = VT_DISPATCH;
			aVar.pdispVal = aToken.object;
			if (!aVarIsArg)
				aToken.object->AddRef();
		}
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


void VarTypeToToken(VARTYPE aVarType, void *apValue, ExprTokenType &aToken)
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


void RValueToResultToken(ExprTokenType &aRValue, ExprTokenType &aResultToken)
// Support function for chaining assignments.  Provides consistent results when aRValue has been
// converted to a VARIANT.  Passes wrapper objects as is, whereas VariantToToken would return
// either a simple value or a new wrapper object.
{
	switch (aRValue.symbol)
	{
	case SYM_OPERAND:
		if (aRValue.buf)
		{
			aResultToken.symbol = SYM_INTEGER;
			aResultToken.value_int64 = *(__int64 *)aRValue.buf;
			break;
		}
		// FALL THROUGH to next case:
	case SYM_STRING:
		aResultToken.symbol = SYM_STRING;
		aResultToken.marker = aRValue.marker;
		break;
	case SYM_OBJECT:
		aResultToken.symbol = SYM_OBJECT;
		aResultToken.object = aRValue.object;
		aResultToken.object->AddRef();
		break;
	case SYM_INTEGER:
	case SYM_FLOAT:
		aResultToken.symbol = aRValue.symbol;
		aResultToken.value_int64 = aRValue.value_int64;
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

	// It's hard to say what our caller will do if DISP_E_MEMBERNOTFOUND is returned,
	// so just return S_OK as in previous versions:
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

ResultType STDMETHODCALLTYPE ComObject::Invoke(ExprTokenType &aResultToken, ExprTokenType &aThisToken, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	if (aParamCount < (IS_INVOKE_SET ? 2 : 1))
	{
		HRESULT hr = DISP_E_BADPARAMCOUNT; // Set default.
		// Something like x[] or x[]:=y.
		if (mVarType & VT_BYREF)
		{
			VARTYPE item_type = mVarType & VT_TYPEMASK;
			if (aParamCount) // Implies SET.
			{
				hr = TokenToVarType(*aParam[0], item_type, mValPtr);
				if (SUCCEEDED(hr))
				{
					RValueToResultToken(*aParam[0], aResultToken);
					return OK;
				}
			}
			else // !aParamCount: Implies GET.
			{
				VarTypeToToken(item_type, mValPtr, aResultToken);
				return OK;
			}
		}
		if ((mVarType & VT_ARRAY) // Not meaningful for SafeArrays.
			|| IS_INVOKE_SET) // Wouldn't be handled correctly below and probably has no real-world use.
		{
			ComError(g->LastError = hr);
			return OK;
		}
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

	DISPID dispid;
	LPTSTR aName;
	HRESULT	hr;
	if (aFlags & IF_NEWENUM)
	{
		hr = S_OK;
		dispid = DISPID_NEWENUM;
		aName = _T("_NewEnum"); // Init for ComError().
	}
	else if (!aParamCount || aParam[0]->symbol == SYM_MISSING)
	{
		// v1.1.20: a[,b] is now permitted.  Rather than treating it the same as 
		// an empty string, invoke the object's default/"value" property/method.
		hr = S_OK;
		dispid = DISPID_VALUE;
		aName = _T("");
	}
	else
	{
		aName = TokenToString(*aParam[0], aResultToken.buf);
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
			else if (IS_INVOKE_CALL && TokenIsEmptyString(*aParam[0]))
			{
				// Fn.() and %Fn%() both produce this condition.  aParam[0] is checked instead of aName
				// because aName is also empty if an object was passed, as for Fn[Obj]() or {X: Fn}.X().
				// Although allowing JScript functions to act as methods of AutoHotkey objects could be
				// useful, a different approach would be needed to pass 'this', such as:
				//  - IDispatchEx::InvokeEx with the named arg DISPID_THIS.
				//  - Fn.call(this) -- this would also work with AutoHotkey functions, but would break
				//    existing meta-function scripts.
				dispid = DISPID_VALUE;
				hr = S_OK;
			}
		}
		if (FAILED(hr))
			aParamCount = 0; // Skip parameter conversion and cleanup.
	}
	
	static DISPID dispidParam = DISPID_PROPERTYPUT;
	DISPPARAMS dispparams = {NULL, NULL, 0, 0};
	VARIANTARG *rgvarg;
	EXCEPINFO excepinfo = {0};
	VARIANT varResult = {0};
	
	if (aParamCount)
		--aParamCount; // Exclude the member name, if any.

	if (aParamCount)
	{
		rgvarg = (VARIANTARG *)_alloca(sizeof(VARIANTARG) * aParamCount);

		for (int i = 0; i < aParamCount; i++)
		{
			TokenToVariant(*aParam[aParamCount-i], rgvarg[i], TRUE);
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
		if (rgvarg[i].vt == VT_BSTR && aParam[aParamCount-i]->symbol != SYM_OBJECT)
			SysFreeString(rgvarg[i].bstrVal);
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
			if (SUCCEEDED(hr))
				RValueToResultToken(rvalue, aResultToken);
			// Otherwise, leave aResultToken blank.
		}
	}

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

static Object *g_IdToName;
static Object *g_NameToId;

STDMETHODIMP IObjectComCompatible::GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId)
{
#ifdef UNICODE
	LPTSTR name = *rgszNames;
#else
	CStringCharFromWChar name_buf(*rgszNames);
	LPTSTR name = const_cast<LPTSTR>(name_buf.GetString());
#endif
	if ( !(g_IdToName || (g_IdToName = Object::Create())) ||
		 !(g_NameToId || (g_NameToId = Object::Create())) )
		return E_OUTOFMEMORY;
	ExprTokenType id;
	if (!g_NameToId->GetItem(id, name))
	{
		if (!g_IdToName->Append(name))
			return E_OUTOFMEMORY;
		id.symbol = SYM_INTEGER;
		id.value_int64 = g_IdToName->GetNumericItemCount();
		if (!g_NameToId->SetItem(name, id))
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
	ExprTokenType param_token[MAX_FUNCTION_PARAMS];
	ExprTokenType *param[MAX_FUNCTION_PARAMS];
	ExprTokenType result_token;
	TCHAR result_token_buf[MAX_NUMBER_SIZE];
	
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
	
	ExprTokenType **first_param = param;
	int param_count = cArgs;
	
	if (dispIdMember > 0)
	{
		if (!g_IdToName->GetItemOffset(param_token[0], dispIdMember - 1))
			return DISP_E_MEMBERNOTFOUND;
		if (IsPureNumeric(param_token[0].marker, FALSE, FALSE)) // o[1] in JScript produces a numeric name.
		{
			param_token[0].symbol = SYM_INTEGER;
			param_token[0].value_int64 = ATOI(param_token[0].marker);
		}
		param[0] = &param_token[0];
		++param_count;
		if (flags == IT_CALL && (wFlags & DISPATCH_PROPERTYGET))
			flags |= IF_CALL_FUNC_ONLY;
	}
	else
	{
		if (dispIdMember != DISPID_VALUE)
			return DISP_E_MEMBERNOTFOUND;
		if (flags == IT_CALL && !(wFlags & DISPATCH_PROPERTYGET))
		{
			// This approach works well for Func, but not for an Object implementing __Call,
			// which always expects a method name or the object whose method is being called:
			//flags = (IT_CALL|IF_FUNCOBJ);
			//++first_param;
			// This is consistent with %func%():
			param_token[0].symbol = SYM_STRING;
			param_token[0].marker = _T("");
			param[0] = &param_token[0];
			++param_count;
		}
		else
		{
			if (flags == IT_CALL) // Obj(X) in VBScript and C#, or Obj[X] in C#
				flags = IT_GET|IF_FUNCOBJ;
			++first_param;
		}
	}
	
	for (UINT i = 1; i <= cArgs; ++i)
	{
		VARIANTARG *pvar = &pDispParams->rgvarg[cArgs-i];
		while (pvar->vt == (VT_BYREF | VT_VARIANT))
			pvar = pvar->pvarVal;
		VariantToToken(*pvar, param_token[i]);
		param[i] = &param_token[i];
	}
	
	result_token.buf = result_token_buf; // May be used below for short return values and misc purposes.
	result_token.marker = _T("");
	result_token.symbol = SYM_STRING;	// These must be initialized for the cleanup code below.
	result_token.mem_to_free = NULL;	//
	
	ExprTokenType this_token;
	this_token.symbol = SYM_OBJECT;
	this_token.object = this;

	HRESULT result_to_return;

	for (;;)
	{
		switch (static_cast<IObject *>(this)->Invoke(result_token, this_token, flags, first_param, param_count))
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
					if (obj->GetItem(token, _T("Message")))
						pExcepInfo->bstrDescription = SysStringFromToken(token, result_token_buf);
					if (obj->GetItem(token, _T("What")))
						pExcepInfo->bstrSource = SysStringFromToken(token, result_token_buf);
					if (obj->GetItem(token, _T("File")))
						pExcepInfo->bstrHelpFile = SysStringFromToken(token, result_token_buf);
					if (obj->GetItem(token, _T("Line")))
						pExcepInfo->dwHelpContext = (DWORD)TokenToInt64(token);
				}
				g_script.FreeExceptionToken(g->ThrownToken);
			}
			break;
		case INVOKE_NOT_HANDLED:
			if ((flags & IT_BITMASK) != IT_GET)
			{
				if (wFlags & DISPATCH_PROPERTYGET)
				{
					flags = IT_GET;
					continue;
				}
				result_to_return = DISP_E_MEMBERNOTFOUND;
				break;
			}
		default:
			result_to_return = S_OK;
			if (pVarResult)
				TokenToVariant(result_token, *pVarResult, FALSE);
		}
		break;
	}

	if (result_token.symbol == SYM_OBJECT)
		result_token.object->Release();
	if (result_token.mem_to_free)
		free(result_token.mem_to_free);

	for (UINT i = 1; i <= cArgs; ++i)
	{
		// Release objects (some or all of which may have been created by VariantToToken()):
		if (param_token[i].symbol == SYM_OBJECT)
			param_token[i].object->Release();
		// Free any temporary memory used to hold strings; see VariantToToken().
		else if (param_token[i].symbol == SYM_STRING && param_token[i].mem_to_free)
			free(param_token[i].mem_to_free);
	}

	return result_to_return;
}


#ifdef CONFIG_DEBUGGER

void WriteComObjType(IDebugProperties *aDebugger, ComObject *aObject, LPCSTR aName, LPTSTR aWhichType)
{
	TCHAR buf[MAX_NUMBER_SIZE];
	ExprTokenType resultToken, paramToken[2], *param[2] = { &paramToken[0], &paramToken[1] };
	paramToken[0].symbol = SYM_OBJECT;
	paramToken[0].object = aObject;
	paramToken[1].symbol = SYM_STRING;
	paramToken[1].marker = aWhichType;
	resultToken.symbol = SYM_INTEGER;
	resultToken.marker = _T("ComObjType");
	resultToken.mem_to_free = NULL;
	resultToken.buf = buf;
	ResultType result = OK;
	BIF_ComObjTypeOrValue(result, resultToken, param, 2);
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
