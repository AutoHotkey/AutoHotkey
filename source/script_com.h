#pragma once


extern bool g_ComErrorNotify;


class ComObject;
class ComEvent : public ObjectBase
{
	DWORD mCookie;
	ComObject *mObject;
	ITypeInfo *mTypeInfo;
	IID mIID;
	IObject *mAhkObject;
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

	HRESULT Connect(LPTSTR pfx = NULL, IObject *ahkObject = NULL);

	ComEvent(ComObject *obj, ITypeInfo *tinfo, IID iid)
		: mCookie(0), mObject(obj), mTypeInfo(tinfo), mIID(iid), mAhkObject(NULL)
	{
	}
	~ComEvent()
	{
		mTypeInfo->Release();
		if (mAhkObject)
			mAhkObject->Release();
	}

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
		// ComObjArray
		M___Enum,
		M_Clone,
		M_MaxIndex,
		M_MinIndex,
	};
	static ObjectMember sArrayMembers[], sRefMembers[], sValueMembers[];

	ResultType Invoke(IObject_Invoke_PARAMS_DECL);
	void Invoke(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount);
	void SafeArrayInvoke(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount);
	void SafeArrayItem(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount);
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
				mEventSink->Connect();
				mEventSink->mObject = NULL;
				mEventSink->Release();
			}
			mUnknown->Release();
		}
		else if ((mVarType & (VT_BYREF|VT_ARRAY)) == VT_ARRAY && (mFlags & F_OWNVALUE))
		{
			SafeArrayDestroy(mArray);
		}
	}
};


class ComEnum : public EnumBase
{
	IEnumVARIANT *penum;

public:
	ResultType Next(Var *aOutput, Var *aOutputType);

	ComEnum(IEnumVARIANT *enm)
		: penum(enm)
	{
	}
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

