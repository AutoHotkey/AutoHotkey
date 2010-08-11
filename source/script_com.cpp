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



void AssignVariant(Var &arg, VARIANT *pvar, bool bsink = true)
{
	switch(pvar->vt)
	{
	case VT_BSTR:
		arg.AssignStringW(pvar->bstrVal, SysStringLen(pvar->bstrVal));
		if (!bsink)
			VariantClear(pvar);
		break;
	case VT_I4:
	case VT_ERROR:
		arg.Assign((__int64)pvar->lVal);
		break;
	case VT_I2:
	case VT_BOOL:
		arg.Assign((__int64)pvar->iVal);
		break;
	case VT_UNKNOWN:
		arg.Assign((__int64)pvar->punkVal);
		break;
	case VT_DISPATCH:
		arg.AssignSkipAddRef(new ComObject(pvar->pdispVal));
		if (bsink && pvar->pdispVal)
			pvar->pdispVal->AddRef();
		break;
	case VT_EMPTY:
	case VT_NULL:
		arg.Assign();
		break;
	default:
		{
			VARIANT var = {0};
			if (pvar->vt < VT_ARRAY
				&& SUCCEEDED(VariantChangeType(&var, pvar, 0, VT_BSTR)))
			{
				arg.AssignStringW(var.bstrVal, SysStringLen(var.bstrVal));
				VariantClear(&var);
			}
			else
				arg.Assign((__int64)pvar->parray);
		}
	}
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
	sntprintf(funcName, _countof(funcName), _T("%s%s"), mPrefix, memberName);
	SysFreeString(memberName);

	Func *func = g_script.FindFunc(funcName);
	if (!func)
		return DISP_E_MEMBERNOTFOUND;

	// TODO: ComEvent::Invoke should follow proper procedure for calling script functions:
	//	Abort if mMinParams is too high; allow any number of optional parameters.
	//	If any instances are already running, back up their local variables.
	//	Set unused optional parameters to their default values.
	//	Abort if ByRef params are encountered; or convert to non-alias.
	//	Pass a result token to Func::Call and use this as the result.
	//	Free/restore local variables; if result is a string, make a persistent copy first.

	UINT cArgs = pDispParams->cArgs;
	UINT mArgs = func->mParamCount;
	//   nArgs = func->mMinParams;
	if (mArgs > cArgs)
	{
		if (mArgs > cArgs + 1 || !mObject) // mObject == NULL should be next to impossible since it is only set NULL after calling Unadvise(), in which case there shouldn't be anyone left to call this->Invoke().  Check it anyway since it might be difficult to debug, depending on what we're connected to.
			return DISP_E_MEMBERNOTFOUND;
		func->mParam[--mArgs].var->Assign(mObject);
	}

	for (UINT i = 0; i < mArgs; i++)
	{
		VARIANTARG *pvar = &pDispParams->rgvarg[cArgs-1-i];
		while (pvar->vt == (VT_BYREF | VT_VARIANT))
			pvar = pvar->pvarVal;
		AssignVariant(*func->mParam[i].var, pvar);
	}

	if (pVarResult)
	{
		ExprTokenType result_token;
		func->Call(&result_token);
		TokenToVariant(result_token, *pVarResult);
		if (result_token.symbol == SYM_OBJECT)
			result_token.object->Release();
	}
	else
		func->Call(NULL);
	
	VarBkp *var_backup = NULL; int var_backup_count;
	Var::FreeAndRestoreFunctionVars(*func, var_backup, var_backup_count);

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
		hr = mDispatch->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, IS_INVOKE_SET ? DISPATCH_PROPERTYPUT : DISPATCH_PROPERTYGET | DISPATCH_METHOD, &dispparams, &varResult, &excepinfo, NULL);

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
			AssignVariant(*aOutput, &varResult, false);
		return true;
	}
	return	false;
}


#endif
