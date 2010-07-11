#pragma once

#ifdef CONFIG_EXPERIMENTAL


extern bool g_ComErrorNotify;


class ComObject;
class ComEvent : public IDispatch
{
	DWORD mRefCount;
	DWORD mCookie;
	ComObject *mObject;
	ITypeInfo *mTypeInfo;
	IID mIID;
	TCHAR mPrefix[64];

public:
	STDMETHODIMP QueryInterface(REFIID riid, void **ppv);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();
	STDMETHODIMP GetTypeInfoCount(UINT *pctinfo);
	STDMETHODIMP GetTypeInfo(UINT itinfo, LCID lcid, ITypeInfo **pptinfo);
	STDMETHODIMP GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId);
	STDMETHODIMP Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr);

	void Connect(LPTSTR pfx = NULL);

	ComEvent(ComObject *obj, ITypeInfo *tinfo, IID iid)
		: mRefCount(1), mCookie(0), mObject(obj), mTypeInfo(tinfo), mIID(iid)
	{
	}
	~ComEvent()
	{
		mTypeInfo->Release();
	}

	friend class ComObject;
};


class ComObject : public ObjectBase
{
	union
	{
		IDispatch *mDispatch;
		__int64 mVal64; // Allow 64-bit values when ComObject is used as a VARIANT in 32-bit builds.
	};
	ComEvent *mEventSink;
	VARTYPE mVarType;

public:
	ResultType STDMETHODCALLTYPE Invoke(ExprTokenType &aResultToken, ExprTokenType &aThisToken, int aFlags, ExprTokenType *aParam[], int aParamCount);

	void ToVariant(VARIANT &aVar)
	{
		aVar.vt = mVarType;
		aVar.llVal = mVal64;
		if (VT_DISPATCH == mVarType && mDispatch)
			mDispatch->AddRef();
	}

	ComObject(IDispatch *disp)
		: mDispatch(disp), mVarType(VT_DISPATCH), mEventSink(NULL) { }
	ComObject(__int64 llVal, VARTYPE vt)
		: mVal64(llVal), mVarType(vt), mEventSink(NULL) { }
	~ComObject()
	{
		if (VT_DISPATCH == mVarType && mDispatch)
		{
			if (mEventSink)
			{
				mEventSink->Connect();
				mEventSink->mObject = NULL;
				mEventSink->Release();
			}
			mDispatch->Release();
		}
	}

	friend void BIF_ComObjActive(ExprTokenType&, ExprTokenType*[], int);
	friend void BIF_ComObjConnect(ExprTokenType&, ExprTokenType*[], int);
	friend class ComEvent;
};


class ComEnum : public EnumBase
{
	IEnumVARIANT *penum;

public:
	int Next(Var *aOutput, Var *aOutputType);

	ComEnum(IEnumVARIANT *enm)
		: penum(enm)
	{
	}
	~ComEnum()
	{
		penum->Release();
	}
};


void ComError(HRESULT, LPTSTR = _T(""), EXCEPINFO* = NULL);


#endif
