#pragma once

#define INVOKE_TYPE			(aFlags & IT_BITMASK)
#define IS_INVOKE_SET		(aFlags & IT_SET)
#define IS_INVOKE_GET		(INVOKE_TYPE == IT_GET)
#define IS_INVOKE_CALL		(aFlags & IT_CALL)
#define IS_INVOKE_META		(aFlags & IF_BYPASS_METAFUNC)

#define INVOKE_NOT_HANDLED	CONDITION_FALSE

//
// ObjectBase - Common base class, implements reference counting.
//

class DECLSPEC_NOVTABLE ObjectBase : public IObjectComCompatible
{
protected:
	ULONG mRefCount;
#ifdef _WIN64
	// Used by Object, but defined here on (x64 builds only) to utilize the space
	// that would otherwise just be padding, due to alignment requirements.
	UINT mFlags; 
#endif
	
	virtual bool Delete()
	{
		delete this; // Derived classes MUST be instantiated with 'new' or override this function.
		return true; // See Release() for comments.
	}

public:
	ULONG STDMETHODCALLTYPE AddRef()
	{
		return ++mRefCount;
	}

	ULONG STDMETHODCALLTYPE Release()
	{
		if (mRefCount == 1)
		{
			// If an object is implemented by script, it may need to run cleanup code before the object
			// is deleted.  This introduces the possibility that before it is deleted, the object ref
			// is copied to another variable (AddRef() is called).  To gracefully handle this, let
			// implementors decide when to delete and just decrement mRefCount if it doesn't happen.
			if (Delete())
				return 0; // Object was deleted, so cannot check or decrement --mRefCount.
			// If the object is implemented correctly, false was returned because:
			//  a) mRefCount > 1; e.g. __delete has "resurrected" the object by storing a reference.
			//  b) There are no more counted references to this object, but its lifetime is tied to
			//     another object, which can "resurrect" this one.  The other object must have some
			//     way to delete this one when appropriate, such as by calling AddRef() & Release().
			// Implementor has ensured Delete() returns false only if delete wasn't called (due to
			// remaining references to this), so we must assume mRefCount > 1.  If Delete() really
			// deletes the object and (erroneously) returns false, checking if mRefCount is still
			// 1 may be just as unsafe as decrementing mRefCount as per usual.
		}
		return --mRefCount;
	}

	ULONG RefCount() { return mRefCount; }

	ObjectBase() : mRefCount(1) {}

	// Declare a virtual destructor for correct 'delete this' behaviour in Delete(),
	// and because it is likely to be more convenient and reliable than overriding
	// Delete(), especially with a chain of derived types.
	virtual ~ObjectBase() {}

	ResultType Invoke(IObject_Invoke_PARAMS_DECL);
};


//
// Helpers for IObject::Invoke implementations.
//

typedef BuiltInFunctionType ObjectCtor;
typedef ResultType (IObject::*ObjectMethod)(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount);
struct ObjectMember
{
	LPTSTR name;
	ObjectMethod method;
	UCHAR id, invokeType, minParams, maxParams;
};

#define Object_Member(name, impl, id, invokeType, ...) \
	{ _T(#name), static_cast<ObjectMethod>(&impl), id, invokeType, __VA_ARGS__ }
#define Object_Method_(name, minP, maxP, impl, id) Object_Member(name, impl,   id,       IT_CALL, minP, maxP)
#define Object_Method(name, minP, maxP)            Object_Member(name, Invoke, M_##name, IT_CALL, minP, maxP)
#define Object_Method1(name, minP, maxP)           Object_Member(name, name,   0,        IT_CALL, minP, maxP)
#define Object_Property_get(name, ...)             Object_Member(name, Invoke, P_##name, IT_GET, __VA_ARGS__)
#define Object_Property_get_set(name, ...)         Object_Member(name, Invoke, P_##name, IT_SET, __VA_ARGS__)
#define MAXP_VARIADIC 255


//
// FlatVector - utility class.
//

template <typename T, typename index_t = ::size_t>
class FlatVector
{
	struct Data
	{
		index_t size;
		index_t length;
	};
	Data *data;
	
	struct OneT : public Data { char zero_buf[sizeof(T)]; }; // zero_buf guarantees zero-termination when used for strings (fixes an issue observed in debug mode).
	static OneT Empty;

	void FreeRange(index_t i, index_t count)
	{
		auto v = Value();
		for (; i < count; ++i)
			v[i].~T();
	}

public:
	FlatVector<T, index_t>() { data = &Empty; }
	~FlatVector<T, index_t>() { Free(); }
	
	void Free()
	{
		if (data->size)
		{
			FreeRange(0, data->length);
			free(data);
			data = &Empty;
		}
	}

	bool SetCapacity(index_t new_size)
	{
		index_t length = data->length;
		ASSERT(new_size > 0 && new_size >= length);
		Data *d = data->size ? data : nullptr;
		if (  !(d = (Data *)realloc(d, new_size * sizeof(T) + sizeof(Data)))  )
			return false;
		data = d;
		data->size = new_size;
		data->length = length; // Only strictly necessary if NULL was passed to realloc.
		return true;
	}

	T *InsertUninitialized(index_t i, index_t count)
	{
		ASSERT(i >= 0 && i <= data->length && count + data->length <= data->size);
		auto p = Value() + i;
		if (i < data->length)
			memmove(p + count, p, (data->length - i) * sizeof(T));
		data->length += count;
		return p;
	}

	void Remove(index_t i, index_t count)
	{
		auto v = Value();
		ASSERT(i >= 0 && i + count <= data->length);
		FreeRange(i, count);
		memmove(v + i, v + i + count, (data->length - i - count) * sizeof(T));
		data->length -= count;
	}

	index_t &Length() { return data->length; }
	index_t Capacity() { return data->size; }
	T *Value() { return (T *)(data + 1); }
	operator T *() { return Value(); }
};

template <typename T, typename index_t>
typename FlatVector<T, index_t>::OneT FlatVector<T, index_t>::Empty;


//
// Property: Invoked when a derived object gets/sets the corresponding key.
//

class Property
{
	IObject *mGet = nullptr, *mSet = nullptr, *mCall = nullptr;

	void SetEtter(IObject *&aMemb, IObject *aFunc)
	{
		if (aFunc) aFunc->AddRef();
		if (aMemb) aMemb->Release();
		aMemb = aFunc;
	}

public:
	// MaxParams is cached for performance.  It is used in cases like x.y[z]:=v to
	// determine whether to GET and then apply the parameters to the result, or just
	// invoke SET with parameters.
	int MinParams = -1, MaxParams = -1;

	Property() {}
	~Property()
	{
		if (mGet)
			mGet->Release();
		if (mSet)
			mSet->Release();
		if (mCall)
			mCall->Release();
	}

	IObject *Getter() { return mGet; }
	IObject *Setter() { return mSet; }
	IObject *Method() { return mCall; }

	void SetGetter(IObject *aFunc) { SetEtter(mGet, aFunc); }
	void SetSetter(IObject *aFunc) { SetEtter(mSet, aFunc); }
	void SetMethod(IObject *aFunc) { SetEtter(mCall, aFunc); }
};


//
// Object - Scriptable associative array.
//

//#define ObjParseIntKey(s, endptr) Exp32or64(UorA(wcstol,strtol),UorA(_wcstoi64,_strtoi64))(s, endptr, 10) // Convert string to IntKeyType, setting errno = ERANGE if overflow occurs.
#define ObjParseIntKey(s, endptr) UorA(_wcstoi64,_strtoi64)(s, endptr, 10) // Convert string to IntKeyType, setting errno = ERANGE if overflow occurs.

class Array;

class Object : public ObjectBase
{
public:
	// The type of an array element index or count.
	// Use unsigned to avoid the need to check for negatives.
	typedef UINT index_t;

protected:
	typedef LPTSTR name_t;
	typedef FlatVector<TCHAR> String;

	struct Variant
	{
		union { // Which of its members is used depends on the value of symbol, below.
			__int64 n_int64;	// for SYM_INTEGER
			double n_double;	// for SYM_FLOAT
			IObject *object;	// for SYM_OBJECT
			String string;		// for SYM_STRING
			Property *prop;		// for SYM_DYNAMIC
		};
		SymbolType symbol;
		// key_c contains the first character of key.s. This utilizes space that would
		// otherwise be unused due to 8-byte alignment. See FindField() for explanation.
		TCHAR key_c;

		Variant() = delete;
		~Variant() { Free(); }

		void Minit(); // Perform minimum initialization.
		bool Assign(LPTSTR str, size_t len = -1, bool exact_size = false);
		void AssignEmptyString();
		void AssignMissing();
		bool Assign(ExprTokenType &val);
		bool Assign(ExprTokenType *val) { return Assign(*val); }
		bool InitCopy(Variant &val);
		void ReturnRef(ResultToken &result);
		void ReturnMove(ResultToken &result);
		void Free();
	
		inline void ToToken(ExprTokenType &aToken); // Used when we want the value as is, in a token.  Does not AddRef() or copy strings.
	};

	struct FieldType : Variant
	{
		name_t name;

		FieldType() = delete;
		~FieldType() { free(name); }
	};

	struct MethodType
	{
		name_t name;
		IObject *func;
		MethodType() = delete;
		~MethodType() { func->Release(); free(name); }
	};

	enum EnumeratorType
	{
		Enum_Properties,
		Enum_Methods
	};

	ResultType GetEnumProp(UINT &aIndex, Var *aName, Var *aVal, int aVarCount);

#ifndef _WIN64
	// This is defined in ObjectBase on x64 builds to save space (due to alignment requirements).
	UINT mFlags;
#endif
	enum Flags : decltype(mFlags)
	{
		ClassPrototype = 0x01,
		NativeClassPrototype = 0x02,
		LastObjectFlag = 0x02
	};

	Object *CloneTo(Object &aTo);
	Object() { mFlags = 0; }
	~Object();
	bool Delete() override;

private:
	Object *mBase = nullptr;
	FlatVector<FieldType, index_t> mFields;

	FieldType *FindField(name_t name, index_t &insert_pos);
	FieldType *FindField(name_t name)
	{
		index_t insert_pos;
		return FindField(name, insert_pos);
	}
	
	FieldType *Insert(name_t name, index_t at);

	bool SetInternalCapacity(index_t new_capacity);
	bool Expand()
	// Expands mFields by at least one field.
	{
		return SetInternalCapacity(mFields.Capacity() ? mFields.Capacity() * 2 : 4);
	}

protected:
	ResultType CallAsMethod(ExprTokenType &aFunc, ResultToken &aResultToken, ExprTokenType &aThisToken, ExprTokenType *aParam[], int aParamCount);
	ResultType CallMeta(LPTSTR aName, ResultToken &aResultToken, ExprTokenType &aThisToken, ExprTokenType *aParam[], int aParamCount);
	ResultType CallMetaVarg(int aFlags, LPTSTR aName, ResultToken &aResultToken, ExprTokenType &aThisToken, ExprTokenType *aParam[], int aParamCount);

public:

	static Object *Create();
	static Object *Create(ExprTokenType *aParam[], int aParamCount, ResultToken *apResultToken = nullptr);

	ResultType New(ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount);
	ResultType Construct(ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount);

	bool HasProp(name_t aName);
	bool HasMethod(name_t aName);
	IObject *GetMethod(name_t name);

	bool HasOwnProps() { return mFields.Length(); }
	bool HasOwnProp(name_t aName)
	{
		return FindField(aName) != nullptr;
	}

	enum class PropType
	{
		None = 0,
		Value,
		Object,
		DynamicValue,
		DynamicMethod,
		DynamicMixed
	};
	PropType GetOwnPropType(name_t aName)
	{
		auto field = FindField(aName);
		if (!field)
			return PropType::None;
		switch (field->symbol)
		{
		case SYM_DYNAMIC:
			if (field->prop->Getter() || field->prop->Getter())
				return field->prop->Method() ? PropType::DynamicMixed : PropType::DynamicValue;
			return field->prop->Method() ? PropType::DynamicMethod : PropType::None;
		case SYM_OBJECT: return PropType::Object;
		default: return PropType::Value;
		}
	}

	bool GetOwnProp(ExprTokenType &aToken, name_t aName)
	{
		auto field = FindField(aName);
		if (!field)
			return false;
		field->ToToken(aToken);
		return true;
	}
	
	IObject *GetOwnPropObj(name_t aName)
	{
		auto field = FindField(aName);
		return field && field->symbol == SYM_OBJECT ? field->object : nullptr;
	}

	bool SetOwnProp(name_t aName, ExprTokenType &aValue)
	{
		index_t insert_pos;
		auto field = FindField(aName, insert_pos);
		if (!field && !(field = Insert(aName, insert_pos)))
			return false;
		return field->Assign(aValue);
	}

	bool SetOwnProp(name_t aName, __int64 aValue) { return SetOwnProp(aName, ExprTokenType(aValue)); }
	bool SetOwnProp(name_t aName, IObject *aValue) { return SetOwnProp(aName, ExprTokenType(aValue)); }
	bool SetOwnProp(name_t aName, LPCTSTR aValue) { return SetOwnProp(aName, ExprTokenType(const_cast<LPTSTR>(aValue))); }

	void DeleteOwnProp(name_t aName)
	{
		auto field = FindField(aName);
		if (field)
			mFields.Remove((index_t)(field - mFields), 1);
	}
	
	Property *DefineProperty(name_t aName);
	bool DefineMethod(name_t aName, IObject *aFunc);
	
	bool CanSetBase(Object *aNewBase);
	ResultType SetBase(Object *aNewBase, ResultToken &aResultToken);
	void SetBase(Object *aNewBase)
	{ 
		if (aNewBase)
			aNewBase->AddRef();
		if (mBase)
			mBase->Release();
		mBase = aNewBase;
	}

	bool IsClassPrototype() { return mFlags & ClassPrototype; }
	bool IsNativeClassPrototype() { return mFlags & NativeClassPrototype; }

	Object *GetNativeBase();
	Object *Base() 
	{
		return mBase; // Callers only want to call Invoke(), so no AddRef is done.
	}

	LPTSTR Type();
	bool IsDerivedFrom(IObject *aBase); // Always false for non-Object objects, but IObject* allows dynamic_cast to be avoided.
	bool IsInstanceOf(Object *aClass);

	void EndClassDefinition();
	Object *GetUnresolvedClass(LPTSTR &aName);
	
	ResultType Invoke(IObject_Invoke_PARAMS_DECL);

	static ObjectMember sMembers[];
	static ObjectMember sClassMembers[];
	static ObjectMember sErrorMembers[];
	static Object *sPrototype, *sClass, *sClassPrototype;

	static Object *CreateRootPrototypes();
	static Object *CreateClass(Object *aPrototype);
	static Object *CreatePrototype(LPTSTR aClassName, Object *aBase = nullptr);
	static Object *CreatePrototype(LPTSTR aClassName, Object *aBase, ObjectMember aMember[], int aMemberCount);
	static Object *DefineMembers(Object *aObject, LPTSTR aClassName, ObjectMember aMember[], int aMemberCount);
	static Object *CreateClass(LPTSTR aClassName, Object *aBase, Object *aPrototype, ObjectCtor aCtor);

	ResultType CallBuiltin(int aID, ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount);

	// Only available as functions:
	ResultType GetCapacity(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount);
	ResultType SetCapacity(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount);
	ResultType PropCount(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount);

	// Methods and functions:
	ResultType DeleteProp(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount);
	ResultType DefineProp(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount);
	ResultType GetOwnPropDesc(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount);
	ResultType HasOwnProp(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount);
	ResultType OwnProps(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount);
	ResultType Clone(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount);
	ResultType GetMethod(ResultToken &aResultToken, name_t aName);

	ResultType Error__New(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount);

	// For pseudo-objects:
	static ObjectMember sValueMembers[];
	static Object *sAnyPrototype, *sPrimitivePrototype, *sStringPrototype
		, *sNumberPrototype, *sIntegerPrototype, *sFloatPrototype;
	static Object *sVarRefPrototype;
	static Object *ValueBase(ExprTokenType &aValue);
	static bool HasBase(ExprTokenType &aValue, IObject *aBase);

	static LPTSTR sMetaFuncName[];

	IObject_DebugWriteProperty_Def;
#ifdef CONFIG_DEBUGGER
	friend class Debugger;
#endif
};


//
// Array
//

class Array : public Object
{
private:
	Variant *mItem = nullptr;
	index_t mLength = 0, mCapacity = 0;

	ResultType SetCapacity(index_t aNewCapacity);
	ResultType EnsureCapacity(index_t aRequired);

	index_t ParamToZeroIndex(ExprTokenType &aParam);

	Array() {}
	
public:
	enum : index_t
	{
		BadIndex = UINT_MAX, // Always >= mLength.
		MaxIndex = INT_MAX // This would need 32GB RAM just for mItem, assuming 16 bytes per element.  Not exceeding INT_MAX might avoid some issues.
	};

	index_t Length() { return mLength; }
	index_t Capacity() { return mCapacity; }
	
	ResultType SetLength(index_t aNewLength);

	template<typename TokenT>
	ResultType InsertAt(index_t aIndex, TokenT aValue[], index_t aCount);
	void       RemoveAt(index_t aIndex, index_t aCount);

	bool Append(ExprTokenType &aValue);
	bool Append(LPTSTR aValue, size_t aValueLength = -1) { return Append(ExprTokenType(aValue, aValueLength)); }
	bool Append(__int64 aValue) { return Append(ExprTokenType(aValue)); }

	Array *Clone();

	bool ItemToToken(index_t aIndex, ExprTokenType &aToken);
	ResultType GetEnumItem(UINT &aIndex, Var *, Var *, int);

	~Array();
	static Array *Create(ExprTokenType *aValue[] = nullptr, index_t aCount = 0);
	static Array *FromArgV(LPTSTR *aArgV, int aArgC);
	static Array *FromEnumerable(ExprTokenType &aEnum);
	ResultType ToStrings(LPTSTR *aStrings, int &aStringCount, int aStringsMax);
	void ToParams(ExprTokenType *token, ExprTokenType **param_list, ExprTokenType **aParam, int aParamCount);

	enum MemberID
	{
		P___Item,
		P_Length,
		P_Capacity,
		M_InsertAt,
		M_Push,
		M_RemoveAt,
		M_Pop,
		M_Has,
		M_Delete,
		M_Clone,
		M___Enum
	};
	static ObjectMember sMembers[];
	static Object *sPrototype;
	ResultType Invoke(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount);
};


//
// Map
//

class Map : public Object
{
	union Key // Which of its members is used depends on the field's position in the mItem array.
	{
		LPTSTR s;
		IntKeyType i;
		IObject *p;
	};
	struct Pair : Variant
	{
		Key key;

		Pair() = delete;
		~Pair() = delete;
	};

	enum MapOption : decltype(mFlags)
	{
		MapCaseless = LastObjectFlag << 1,
		MapUseLocale = MapCaseless << 1
	};

	Pair *mItem = nullptr;
	index_t mCount = 0, mCapacity = 0;

	// Holds the index of the first key of a given type within mItem.  Must be in the order: int, object, string.
	// Compared to storing the key-type with each key-value pair, this approach saves 4 bytes per key (excluding
	// the 8 bytes taken by the two fields below) and speeds up lookups since only the section within mItem
	// with the appropriate type of key needs to be searched (and no need to check the type of each key).
	// mKeyOffsetObject should be set to mKeyOffsetInt + the number of int keys.
	// mKeyOffsetString should be set to mKeyOffsetObject + the number of object keys.
	// mKeyOffsetObject-1, mKeyOffsetString-1 and mFieldCount-1 indicate the last index of each prior type.
	static const index_t mKeyOffsetInt = 0;
	index_t mKeyOffsetObject = 0, mKeyOffsetString = 0;

	Map() {}
	void Clear();
	~Map()
	{
		Clear();
		free(mItem);
	}
	 
	Pair *FindItem(LPTSTR val, index_t left, index_t right, index_t &insert_pos);
	Pair *FindItem(IntKeyType val, index_t left, index_t right, index_t &insert_pos);
	Pair *FindItem(SymbolType key_type, Key key, index_t &insert_pos);	
	Pair *FindItem(ExprTokenType &key_token, LPTSTR aBuf, SymbolType &key_type, Key &key, index_t &insert_pos);

	void ConvertKey(ExprTokenType &key_token, LPTSTR buf, SymbolType &key_type, Key &key);

	Pair *Insert(SymbolType key_type, Key key, index_t at);

	bool SetInternalCapacity(index_t new_capacity);
	
	// Expands mItem by at least one field.
	bool Expand()
	{
		return SetInternalCapacity(mCapacity ? mCapacity * 2 : 4);
	}

	Map *CloneTo(Map &aTo);

	ResultType GetEnumItem(UINT &aIndex, Var *, Var *, int);

public:
	static Map *Create(ExprTokenType *aParam[] = NULL, int aParamCount = 0);

	bool HasItem(ExprTokenType &aKey)
	{
		return GetItem(ExprTokenType(), aKey); // Conserves code size vs. calling FindItem() directly and is unlikely to perform worse.
	}

	bool GetItem(ExprTokenType &aToken, ExprTokenType &aKey)
	{
		index_t insert_pos;
		TCHAR buf[MAX_NUMBER_SIZE];
		SymbolType key_type;
		Key key;
		auto item = FindItem(aKey, buf, key_type, key, insert_pos);
		if (!item)
			return false;
		item->ToToken(aToken);
		return true;
	}

	bool GetItem(ExprTokenType &aToken, LPTSTR aKey)
	{
		ExprTokenType key;
		key.symbol = SYM_STRING;
		key.marker = aKey;
		return GetItem(aToken, key);
	}

	bool SetItem(ExprTokenType &aKey, ExprTokenType &aValue)
	{
		index_t insert_pos;
		TCHAR buf[MAX_NUMBER_SIZE];
		SymbolType key_type;
		Key key;
		auto item = FindItem(aKey, buf, key_type, key, insert_pos);
		if (!item && !(item = Insert(key_type, key, insert_pos))) // Relies on short-circuit boolean evaluation.
			return false;
		return item->Assign(aValue);
	}

	bool SetItem(LPTSTR aKey, ExprTokenType &aValue)
	{
		return SetItem(ExprTokenType(aKey), aValue);
	}

	bool SetItem(LPTSTR aKey, __int64 aValue)
	{
		return SetItem(aKey, ExprTokenType(aValue));
	}

	bool SetItem(LPTSTR aKey, IObject *aValue)
	{
		return SetItem(aKey, ExprTokenType(aValue));
	}

	ResultType SetItems(ExprTokenType *aParam[], int aParamCount);

	// Methods callable by script.
	ResultType __Item(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount);
	ResultType Set(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount);
	ResultType Capacity(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount);
	ResultType Count(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount);
	ResultType CaseSense(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount);
	ResultType Clear(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount);
	ResultType Delete(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount);
	ResultType __Enum(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount);
	ResultType Has(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount);
	ResultType Clone(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount);

	static ObjectMember sMembers[];
	static Object *sPrototype;
};


//
// RegExMatchObject:  Returned by RegExMatch via UnquotedOutputVar.
//
class RegExMatchObject : public Object
{
	LPTSTR mHaystack;
	int mHaystackStart;
	int *mOffset;
	LPTSTR *mPatternName;
	int mPatternCount;
	LPTSTR mMark;

	ResultType GetEnumItem(UINT &aIndex, Var *, Var *, int);

	RegExMatchObject() : mHaystack(NULL), mOffset(NULL), mPatternName(NULL), mPatternCount(0), mMark(NULL) {}
	
	~RegExMatchObject()
	{
		if (mHaystack)
			free(mHaystack);
		if (mOffset)
			free(mOffset);
		if (mPatternName)
		{
			// Free the strings:
			for (int p = 1; p < mPatternCount; ++p) // Start at 1 since 0 never has a name.
				if (mPatternName[p])
					free(mPatternName[p]);
			// Free the array:
			free(mPatternName);
		}
		if (mMark)
			free(mMark);
	}

public:
	static ResultType Create(LPCTSTR aHaystack, int *aOffset, LPCTSTR *aPatternName
		, int aPatternCount, int aCapturedPatternCount, LPCTSTR aMark, IObject *&aNewObject);
	
	enum MemberID
	{
		M_Value,
		M_Pos,
		M_Len,
		M_Name,
		M_Count,
		M_Mark,
		M___Enum
	};
	static ObjectMember sMembers[];
	static Object *sPrototype;
	ResultType Invoke(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount);
};


//
// Buffer
//

class BufferObject : public Object
{
protected:
	void *mData;
	size_t mSize;

public:
	void *Data() { return mData; }
	size_t Size() { return mSize; }
	ResultType Resize(size_t aNewSize);

	BufferObject() = delete;
	BufferObject(void *aData, size_t aSize) : mData(aData), mSize(aSize) {}
	~BufferObject() { free(mData); }

	enum MemberID
	{
		P_Ptr,
		P_Size
	};
	static ObjectMember sMembers[];
	static Object *sPrototype;
	ResultType Invoke(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount);
};


//
// ClipboardAll: Represents a blob of clipboard data (all formats retrieved from clipboard).
//

class ClipboardAll : public BufferObject
{
public:
	ClipboardAll() : BufferObject(nullptr, 0) {}
	static ObjectMember sMembers[];
	static Object *sPrototype;
	static Object *Create();
	ResultType __New(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount);
};



void DefineFileClass();



namespace ErrorPrototype
{
	extern Object *Error, *Memory, *Type, *Value, *OS, *ZeroDivision;
	extern Object *Target, *Member, *Property, *Method, *Index, *Key;
	extern Object *Timeout;
}



ResultType GetEnumerator(IObject *&aEnumerator, ExprTokenType &aEnumerable, int aVarCount, bool aDisplayError);
ResultType CallEnumerator(IObject *aEnumerator, ExprTokenType *aParam[], int aParamCount, bool aDisplayError);
