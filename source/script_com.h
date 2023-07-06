#pragma once


extern bool g_ComErrorNotify;


class ComObject;
class ComEvent : public ObjectBase
{
	DWORD mCookie = 0;
	ComObject *mObject;
	ITypeInfo *mTypeInfo = nullptr;
	IID mIID;
	IObject *mAhkObject = nullptr;
	TCHAR mPrefix[64];

public:
	STDMETHODIMP QueryInterface(REFIID riid, void **ppv);
	STDMETHODIMP GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId);
	STDMETHODIMP Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr);
	
	// IObject::Invoke() and Type() are unlikely to be called, since that would mean
	// the script has a reference to the object, which means either that the script
	// itself has implemented IConnectionPoint (and why would it?), or has used the
	// IEnumConnections interface to retrieve its own object (unlikely).
	//ResultType Invoke(IObject_Invoke_PARAMS_DECL); // ObjectBase::Invoke is sufficient.
	IObject_Type_Impl("ComEvent") // Unlikely to be called; see above.
	Object *Base() { return nullptr; }

	HRESULT Connect(ITypeInfo *tinfo = nullptr, IID *iid = nullptr);
	void SetPrefixOrSink(LPCTSTR pfx, IObject *ahkObject);

	ComEvent(ComObject *obj) : mObject(obj) { }
	~ComEvent();

	friend class ComObject;
};


class ComObject : public ObjectBase
{
public:
	union
	{
		IDispatch *mDispatch;
		IUnknown *mUnknown;
		SAFEARRAY *mArray;
		void *mValPtr;
		__int64 mVal64; // Allow 64-bit values when ComObject is used as a VARIANT in 32-bit builds.
	};
	ComEvent *mEventSink;
	VARTYPE mVarType;
	enum { F_OWNVALUE = 1 };
	USHORT mFlags;

	enum
	{
		P_Ptr,
		// ComValueRef
		P___Item,
	};
	static ObjectMember sRefMembers[], sValueMembers[];
	static ObjectMemberMd sArrayMembers[];
	
	FResult SafeArray_Item(VariantParams &aParam, ExprTokenType *aNewValue, ResultToken *aResultToken);
	FResult set_SafeArray_Item(ExprTokenType &aNewValue, VariantParams &aParam) { return SafeArray_Item(aParam, &aNewValue, nullptr); }
	FResult get_SafeArray_Item(ResultToken &aResultToken, VariantParams &aParam) { return SafeArray_Item(aParam, nullptr, &aResultToken); }
	
	FResult SafeArray_Clone(IObject *&aRetVal);
	FResult SafeArray_Enum(optl<int>, IObject *&aRetVal);
	FResult SafeArray_MaxIndex(optl<UINT> aDims, int &aRetVal);
	FResult SafeArray_MinIndex(optl<UINT> aDims, int &aRetVal);

	ResultType Invoke(IObject_Invoke_PARAMS_DECL);
	void Invoke(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount);
	LPTSTR Type();
	Object *Base();
	IObject_DebugWriteProperty_Def;

	void ToVariant(VARIANT &aVar)
	{
		aVar.vt = mVarType;
		aVar.llVal = mVal64;
		// Caller handles this if needed:
		//if (VT_DISPATCH == mVarType && mDispatch)
		//	mDispatch->AddRef();
	}

	ComObject(IDispatch *pdisp)
		: mVal64((__int64)pdisp), mVarType(VT_DISPATCH), mEventSink(NULL), mFlags(0) { }
	ComObject(__int64 llVal, VARTYPE vt, USHORT flags = 0)
		: mVal64(llVal), mVarType(vt), mEventSink(NULL), mFlags(flags) { }
	~ComObject()
	{
		if ((VT_DISPATCH == mVarType || VT_UNKNOWN == mVarType) && mUnknown)
		{
			if (mEventSink)
			{
				mEventSink->Connect(FALSE);
				if (mEventSink) // i.e. it wasn't fully released as a result of calling Unadvise().
					mEventSink->mObject = nullptr;
			}
			mUnknown->Release();
		}
		else if ((mVarType & (VT_BYREF|VT_ARRAY)) == VT_ARRAY && (mFlags & F_OWNVALUE))
		{
			SafeArrayDestroy(mArray);
		}
		else if (mVarType == VT_BSTR && (mFlags & F_OWNVALUE))
		{
			SysFreeString((BSTR)mValPtr);
		}
	}
};


class ComEnum : public EnumBase
{
	IEnumVARIANT *penum;
	bool cheat;

public:
	ResultType Next(Var *aOutput, Var *aOutputType);

	ComEnum(IEnumVARIANT *enm);
	~ComEnum()
	{
		penum->Release();
	}
};


class ComArrayEnum : public EnumBase
{
	ComObject *mArrayObject;
	void *mData;
	long mLBound, mUBound;
	UINT mElemSize;
	VARTYPE mType;
	bool mIndexMode;
	long mOffset = -1;

	ComArrayEnum(ComObject *aObj, void *aData, long aLBound, long aUBound, UINT aElemSize, VARTYPE aType, bool aIndexMode)
		: mArrayObject(aObj), mData(aData), mLBound(aLBound), mUBound(aUBound), mElemSize(aElemSize), mType(aType), mIndexMode(aIndexMode)
	{
	}

public:
	static HRESULT Begin(ComObject *aArrayObject, ComArrayEnum *&aOutput, int aMode);
	ResultType Next(Var *aVar1, Var *aVar2);
	~ComArrayEnum();
};


// Adapts an AutoHotkey enumerator object to the IEnumVARIANT COM interface.
class EnumComCompat : public IEnumVARIANT, public IServiceProvider
{
	IObject *mEnum;
	int mRefCount;
	bool mCheat;

public:
	EnumComCompat(IObject *enumObj) : mEnum(enumObj), mRefCount(1), mCheat(false) {}
	~EnumComCompat() { mEnum->Release(); }

	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();

	STDMETHODIMP Next(ULONG celt, /*out*/ VARIANT *rgVar, /*out*/ ULONG *pCeltFetched);
	STDMETHODIMP Skip(ULONG celt);
	STDMETHODIMP Reset();
	STDMETHODIMP Clone(/*out*/ IEnumVARIANT **ppEnum);

	STDMETHODIMP QueryService(REFGUID guidService, REFIID riid, void **ppvObject);
};


enum TTVArgType
{
	VariantIsValue,
	VariantIsAllocatedString,
	VariantIsVarRef
};
void TokenToVariant(ExprTokenType &aToken, VARIANT &aVar, TTVArgType *aVarIsArg = FALSE);
HRESULT TokenToVarType(ExprTokenType &aToken, VARTYPE aVarType, void *apValue, bool aCallerIsComValue = false);

void ComError(HRESULT, ResultToken &, LPTSTR = _T(""), EXCEPINFO* = NULL);

bool SafeSetTokenObject(ExprTokenType &aToken, IObject *aObject);

