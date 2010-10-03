#include "stdafx.h"
#include "globaldata.h"
#include "script.h"
#include "script_object.h"
#include "script_com.h"

#ifdef CONFIG_EXPERIMENTAL


void BIF_ComObjCreate(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	HRESULT hr;
	CLSID clsid;
	IDispatch *pdisp;
	hr = CLSIDFromString(CStringWCharFromTCharIfNeeded(TokenToString(*aParam[0])), &clsid);
	if (SUCCEEDED(hr))
		hr = CoCreateInstance(clsid, NULL, CLSCTX_SERVER, IID_IDispatch, (void **)&pdisp);
	if (SUCCEEDED(hr))
	{
		aResultToken.symbol = SYM_OBJECT;
		aResultToken.object = new ComObject(pdisp);
	}
	else
	{
		aResultToken.symbol = SYM_STRING;
		aResultToken.marker = _T("");
		ComError(hr);
	}
}

void BIF_ComObjGet(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	HRESULT hr;
	IDispatch *pdisp;
	hr = CoGetObject(CStringWCharFromTCharIfNeeded(TokenToString(*aParam[0])), NULL, IID_IDispatch, (void **)&pdisp);
	if (SUCCEEDED(hr))
	{
		aResultToken.symbol = SYM_OBJECT;
		aResultToken.object = new ComObject(pdisp);
	}
	else
	{
		aResultToken.symbol = SYM_STRING;
		aResultToken.marker = _T("");
		ComError(hr);
	}
}

void BIF_ComObjActive(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	aResultToken.symbol = SYM_STRING;
	aResultToken.marker = _T("");

	if (TokenIsPureNumeric(*aParam[0]))
	{
		if (aParamCount > 1)
		{
			VARTYPE	vt = (VARTYPE)TokenToInt64(*aParam[0]);
			__int64 llVal = TokenToInt64(*aParam[1]);
			aResultToken.symbol = SYM_OBJECT;
			aResultToken.object = new ComObject(llVal, vt);
			if (vt == VT_DISPATCH && (IDispatch *)llVal) // Type-cast so upper 32-bits are ignored in 32-bit builds.
				((IDispatch *)llVal)->AddRef();
		}
		else if (IUnknown *punk = (IUnknown *)TokenToInt64(*aParam[0]))
		{
			IDispatch *pdisp;
			if (FAILED(punk->QueryInterface(IID_IDispatch, (void **)&pdisp)))
			{
				pdisp = (IDispatch *)punk;
				pdisp->AddRef();
			}
			aResultToken.symbol = SYM_OBJECT;
			aResultToken.object = new ComObject(pdisp);
		}
		else
			ComError(-1);
	}
	else if (ComObject *obj = dynamic_cast<ComObject *>(TokenToObject(*aParam[0])))
	{
		if (aParamCount > 1)
		{
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
					aResultToken.symbol = SYM_OBJECT;
					aResultToken.object = new ComObject(pdisp);
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
			obj->mEventSink->Connect(aParamCount>1 ? TokenToString(*aParam[1]) : NULL);
		else
			ComError(-1);
	}
	else if	(aParamCount == 1 && TokenIsPureNumeric(*aParam[0]))
		g_ComErrorNotify = (TokenToInt64(*aParam[0]) != 0);
}


void VariantToToken(VARIANT &aVar, ExprTokenType &aToken, bool aRetainVar = true)
{
	switch (aVar.vt)
	{
	case VT_BSTR:
		aToken.symbol = SYM_STRING;
#ifdef UNICODE
		aToken.marker = aVar.bstrVal;
#else
		{
			UINT len = SysStringLen(aVar.bstrVal);
			if (len)
			{
				CStringCharFromWChar buf(aVar.bstrVal, len);
				// Caller knows to free this afterward:
				if ( !(aToken.marker = buf.DetachBuffer()) )
					aToken.marker = Var::sEmptyString;
			}
			else
				aToken.marker = Var::sEmptyString; // Return an empty value which caller knows not to free().
		}
#endif
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
	case VT_UNKNOWN:
		aToken.symbol = SYM_INTEGER;
		aToken.value_int64 = (__int64)aVar.punkVal;
		break;
	case VT_DISPATCH:
		if (aToken.object = new ComObject(aVar.pdispVal))
		{
			aToken.symbol = SYM_OBJECT;
			if (aRetainVar && aVar.pdispVal)
				aVar.pdispVal->AddRef();
			//else we're taking ownership of the reference.
			break;
		}
		// Above failed, so check if we need to release the pointer:
		if (!aRetainVar && aVar.pdispVal)
			aVar.pdispVal->Release();
		// FALL THROUGH to the next case:
	case VT_EMPTY:
	case VT_NULL:
		aToken.symbol = SYM_STRING;
		aToken.marker = Var::sEmptyString; // For ANSI builds: return an empty value which caller knows not to free().
		break;
	default:
		{
			VARIANT var = {0};
			if (aVar.vt < VT_ARRAY // i.e. not byref or an array.
				&& SUCCEEDED(VariantChangeType(&var, &aVar, 0, VT_BSTR)))
			{
				// Above put a string representation of aVar into var.
				// Recursively call self for simplicitly (esp. for ANSI builds):
				VariantToToken(var, aToken, false);
			}
			else
			{
				aToken.symbol = SYM_INTEGER;
				aToken.value_int64 = (__int64)aVar.parray;
			}
		}
	}
}

void AssignVariant(Var &aArg, VARIANT &aVar, bool aRetainVar = true)
{
	if (aVar.vt == VT_BSTR)
	{
		// Avoid an unnecessary mem alloc and copy in ANSI builds.
		aArg.AssignStringW(aVar.bstrVal, SysStringLen(aVar.bstrVal));
		if (!aRetainVar)
			VariantClear(&aVar);
		return;
	}
	ExprTokenType token;
	VariantToToken(aVar, token, aRetainVar);
	if (token.symbol == SYM_OBJECT)
		aArg.AssignSkipAddRef(token.object);
	else
		aArg.Assign(token);
}


void TokenToVariant(ExprTokenType &aToken, VARIANT &aVar)
{
	if (aToken.symbol == SYM_VAR)
		aToken.var->TokenToContents(aToken);

	switch(aToken.symbol)
	{
	case SYM_OPERAND:
		if(aToken.buf)
		{
			aVar.vt = VT_I4;
			aVar.lVal = *(int *)aToken.buf;
			break;
		}
	case SYM_STRING:
		aVar.vt = VT_BSTR;
		aVar.bstrVal = SysAllocString(CStringWCharFromTCharIfNeeded(aToken.marker));
		break;
	case SYM_INTEGER:
		aVar.vt = VT_I4;
		aVar.lVal = (int)aToken.value_int64;
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
#ifndef UNICODE
		// Free any temporary memory used to hold ANSI strings; see VariantToToken().
		else if (param_token[i].symbol == SYM_STRING && param_token[i].marker != Var::sEmptyString)
			free(param_token[i].marker);
#endif
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
		return OK;

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
	if (SUCCEEDED(hr) && !(IS_INVOKE_SET && rgvarg[0].vt == VT_DISPATCH
		&& SUCCEEDED(mDispatch->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUTREF, &dispparams, NULL, NULL, NULL))))
		hr = mDispatch->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, IS_INVOKE_SET ? DISPATCH_PROPERTYPUT : IS_INVOKE_CALL && !aParamCount ? DISPATCH_METHOD : DISPATCH_PROPERTYGET | DISPATCH_METHOD, &dispparams, &varResult, &excepinfo, NULL);

	for (int i = 0; i < aParamCount; i++)
		VariantClear(&rgvarg[i]);

	if	(FAILED(hr))
		ComError(hr, aName, &excepinfo);
	else if	(IS_INVOKE_SET)
	{	// Allow chaining, e.g. obj2.prop := obj1.prop := val.
		ExprTokenType &rvalue = *aParam[aParamCount];
		aResultToken.symbol = rvalue.symbol;
		aResultToken.value_int64 = rvalue.value_int64;
		if (rvalue.symbol == SYM_OBJECT)
			rvalue.object->AddRef();
	}
	else
	{
		switch(varResult.vt)
		{
		case VT_BSTR:
			TokenSetResult(aResultToken, CStringTCharFromWCharIfNeeded(varResult.bstrVal), SysStringLen(varResult.bstrVal));
			VariantClear(&varResult);
			break;
		case VT_I4:
		case VT_ERROR:
			aResultToken.symbol = SYM_INTEGER;
			aResultToken.value_int64 = varResult.lVal;
			break;
		case VT_I2:
		case VT_BOOL:
			aResultToken.symbol = SYM_INTEGER;
			aResultToken.value_int64 = varResult.iVal;
			break;
		case VT_UNKNOWN:
			if (varResult.punkVal)
			{
				IEnumVARIANT *penum;
				if (SUCCEEDED(varResult.punkVal->QueryInterface(IID_IEnumVARIANT, (void**) &penum)))
				{
					varResult.punkVal->Release();
					aResultToken.symbol = SYM_OBJECT;
					aResultToken.object = new ComEnum(penum);
					break;
				}
			}
			aResultToken.symbol = SYM_INTEGER;
			aResultToken.value_int64 = (__int64) varResult.punkVal;
			break;
		case VT_DISPATCH:
			if (varResult.pdispVal)
			{
				aResultToken.symbol = SYM_OBJECT;
				aResultToken.object = new ComObject(varResult.pdispVal);
				break;
			}
		case VT_EMPTY:
		case VT_NULL:
			break;
		default:
			if (varResult.vt < VT_ARRAY
				&& SUCCEEDED(VariantChangeType(&varResult, &varResult, 0, VT_BSTR)))
			{
				TokenSetResult(aResultToken, CStringTCharFromWCharIfNeeded(varResult.bstrVal), SysStringLen(varResult.bstrVal));
				VariantClear(&varResult);
			}
			else
			{
				aResultToken.symbol = SYM_INTEGER;
				aResultToken.value_int64 = (__int64) varResult.parray;
			}
		}
	}

	g->LastError = hr;
	return	OK;
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


#endif
