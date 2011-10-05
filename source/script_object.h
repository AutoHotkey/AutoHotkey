#pragma once

#define INVOKE_TYPE			(aFlags & IT_BITMASK)
#define IS_INVOKE_SET		(aFlags & IT_SET)
#define IS_INVOKE_GET		(INVOKE_TYPE == IT_GET)
#define IS_INVOKE_CALL		(aFlags & IT_CALL)
#define IS_INVOKE_META		(aFlags & IF_METAOBJ)
#define SHOULD_INVOKE_METAFUNC (aFlags & IF_METAFUNC)

#define INVOKE_NOT_HANDLED	CONDITION_FALSE

//
// ObjectBase - Common base class, implements reference counting.
//

class DECLSPEC_NOVTABLE ObjectBase : public IObject
{
protected:
	ULONG mRefCount;
	
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
				return 0;
			// Implementor has ensured Delete() returns false only if delete wasn't called (due to
			// remaining references to this), so we must assume mRefCount > 1.  If Delete() really
			// deletes the object and (erroneously) returns false, checking if mRefCount is still
			// 1 may be just as unsafe as decrementing mRefCount as per usual.
		}
		return --mRefCount;
	}

	ObjectBase() : mRefCount(1) {}

	// Declare a virtual destructor for correct 'delete this' behaviour in Delete(),
	// and because it is likely to be more convenient and reliable than overriding
	// Delete(), especially with a chain of derived types.
	virtual ~ObjectBase() {}
};	


//
// EnumBase - Base class for enumerator objects following standard syntax.
//

class EnumBase : public ObjectBase
{
public:
	ResultType STDMETHODCALLTYPE Invoke(ExprTokenType &aResultToken, ExprTokenType &aThisToken, int aFlags, ExprTokenType *aParam[], int aParamCount);
	virtual int Next(Var *aOutputVar1, Var *aOutputVar2) = 0;
};
	

//
// Object - Scriptable associative array.
//

class Object : public ObjectBase
{
protected:
	typedef INT_PTR IntKeyType; // Same size as the other union members.
	typedef INT_PTR IndexType; // Type of index for the internal array.  Must be signed for FindKey to work correctly.
	union KeyType // Which of its members is used depends on the field's position in the mFields array.
	{
		LPTSTR s;
		IntKeyType i;
		IObject *p;
	};
	struct FieldType
	{
		union { // Which of its members is used depends on the value of symbol, below.
			__int64 n_int64;	// for SYM_INTEGER
			double n_double;	// for SYM_FLOAT
			IObject *object;	// for SYM_OBJECT
			struct {
				LPTSTR marker;		// for SYM_OPERAND
				size_t size;		// for SYM_OPERAND; allows reuse of allocated memory. For UNICODE: count in characters
			};
		};
		// key and symbol probably need to be adjacent to each other to conserve memory due to 8-byte alignment.
		KeyType key;
		SymbolType symbol;
		
		inline IntKeyType CompareKey(IntKeyType val) { return val - key.i; }  // Used by both int and object since they are stored separately.
		inline int CompareKey(LPTSTR val) { return _tcsicmp(val, key.s); }

		bool Assign(LPTSTR str, size_t len = -1, bool exact_size = false);
		bool Assign(ExprTokenType &val);
		void Get(ExprTokenType &result);
		void Free();
	
		inline void ToToken(ExprTokenType &aToken) // Used when we want the value as is, in a token.  Does not AddRef() or copy strings.
		{
			aToken.value_int64 = n_int64; // Union copy. Overlaps with buf on x86 builds, so do it first.
			if ((aToken.symbol = symbol) == SYM_OPERAND)
				aToken.buf = NULL; // Indicate that this SYM_OPERAND token LACKS a pre-converted binary integer.
		}
	};

	class Enumerator : public EnumBase
	{
		Object *mObject;
		IndexType mOffset;
	public:
		Enumerator(Object *aObject) : mObject(aObject), mOffset(-1) { mObject->AddRef(); }
		~Enumerator() { mObject->Release(); }
		int Next(Var *aKey, Var *aVal);
	};
	
	IObject *mBase;
	FieldType *mFields;
	IndexType mFieldCount, mFieldCountMax; // Current/max number of fields.

	// Holds the index of first key of a given type within mFields.  Must be in the order: int, object, string.
	// Compared to storing the key-type with each key-value pair, this approach saves 4 bytes per key (excluding
	// the 8 bytes taken by the two fields below) and speeds up lookups since only the section within mFields
	// with the appropriate type of key needs to be searched (and no need to check the type of each key).
	// mKeyOffsetObject should be set to mKeyOffsetInt + the number of int keys.
	// mKeyOffsetString should be set to mKeyOffsetObject + the number of object keys.
	// mKeyOffsetObject-1, mKeyOffsetString-1 and mFieldCount-1 indicate the last index of each prior type.
	static const IndexType mKeyOffsetInt = 0;
	IndexType mKeyOffsetObject, mKeyOffsetString;

#ifdef CONFIG_DEBUGGER
	friend class Debugger;
#endif

	Object()
		: mBase(NULL)
		, mFields(NULL), mFieldCount(0), mFieldCountMax(0)
		, mKeyOffsetObject(0), mKeyOffsetString(0)
	{}

	bool Delete();
	~Object();

	template<typename T>
	FieldType *FindField(T val, IndexType left, IndexType right, IndexType &insert_pos);
	FieldType *FindField(SymbolType key_type, KeyType key, IndexType &insert_pos);	
	FieldType *FindField(ExprTokenType &key_token, LPTSTR aBuf, SymbolType &key_type, KeyType &key, IndexType &insert_pos);
	
	FieldType *Insert(SymbolType key_type, KeyType key, IndexType at);

	bool SetInternalCapacity(IndexType new_capacity);
	bool Expand()
	// Expands mFields by at least one field.
	{
		return SetInternalCapacity(mFieldCountMax ? mFieldCountMax * 2 : 4);
	}
	
	ResultType CallField(FieldType *aField, ExprTokenType &aResultToken, ExprTokenType &aThisToken, int aFlags, ExprTokenType *aParam[], int aParamCount);
	
public:
	static Object *Create(ExprTokenType *aParam[], int aParamCount);

	// Used by Func::Call() for variadic functions/function-calls:
	Object *Clone(INT_PTR aStartOffset = 0);
	ResultType ArrayToParams(void *&aMemToFree, ExprTokenType **&aParam, int &aParamCount, int aMinParams);
	
	inline bool GetNextItem(ExprTokenType &aToken, INT_PTR &aOffset, INT_PTR &aKey)
	{
		if (++aOffset >= mKeyOffsetObject) // i.e. no more integer-keyed items.
			return false;
		FieldType &field = mFields[aOffset];
		aKey = field.key.i;
		field.ToToken(aToken);
		return true;
	}
	
	bool GetItem(ExprTokenType &aToken, LPTSTR aKey)
	{
		KeyType key;
		SymbolType key_type = IsPureNumeric(aKey, FALSE, FALSE, FALSE); // SYM_STRING or SYM_INTEGER.
		if (key_type == SYM_INTEGER)
			key.i = ATOI(aKey);
		else
			key.s = aKey;
		IndexType insert_pos;
		FieldType *field = FindField(key_type, key, insert_pos);
		if (!field)
			return false;
		field->ToToken(aToken);
		return true;
	}
	
	bool SetItem(LPTSTR aKey, ExprTokenType &aValue)
	{
		KeyType key;
		SymbolType key_type = IsPureNumeric(aKey, FALSE, FALSE, FALSE); // SYM_STRING or SYM_INTEGER.
		if (key_type == SYM_INTEGER)
			key.i = ATOI(aKey);
		else
			key.s = aKey;
		IndexType insert_pos;
		FieldType *field = FindField(key_type, key, insert_pos);
		if (  !field && !(field = Insert(key_type, key, insert_pos))  ) // Relies on short-circuit boolean evaluation.
			return false;
		return field->Assign(aValue);
	}

	bool SetItem(LPTSTR aKey, IObject *aValue)
	{
		ExprTokenType token;
		token.symbol = SYM_OBJECT;
		token.object = aValue;
		return SetItem(aKey, token);
	}
	
	void ReduceKeys(INT_PTR aAmount)
	{
		for (IndexType i = 0; i < mKeyOffsetObject; ++i)
			mFields[i].key.i -= aAmount;
	}

	void SetBase(IObject *aNewBase)
	{ 
		if (aNewBase)
			aNewBase->AddRef();
		if (mBase)
			mBase->Release();
		mBase = aNewBase;
	}

	IObject *Base() 
	{
		return mBase; // Callers only want to call Invoke(), so no AddRef is done.
	}
	
	// Used by Object::_Insert() and Func::Call():
	bool InsertAt(INT_PTR aOffset, INT_PTR aKey, ExprTokenType *aValue[], int aValueCount);

	void EndClassDefinition();
	
	ResultType STDMETHODCALLTYPE Invoke(ExprTokenType &aResultToken, ExprTokenType &aThisToken, int aFlags, ExprTokenType *aParam[], int aParamCount);

	ResultType _Insert(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);
	ResultType _Remove(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);
	ResultType _GetCapacity(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);
	ResultType _SetCapacity(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);
	ResultType _GetAddress(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);
	ResultType _MaxIndex(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);
	ResultType _MinIndex(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);
	ResultType _NewEnum(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);
	ResultType _HasKey(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);
	ResultType _Clone(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);

	static LPTSTR sMetaFuncName[];
};


//
// MetaObject:	Used only by g_MetaObject (not every meta-object); see comments below.
//
class MetaObject : public Object
{
public:
	// In addition to ensuring g_MetaObject is never "deleted", this avoids a
	// tiny bit of work when any reference to this object is added or released.
	// Temporary references such as when evaluating "".base.foo are most common.
	ULONG STDMETHODCALLTYPE AddRef() { return 1; }
	ULONG STDMETHODCALLTYPE Release() { return 1; }
	bool Delete() { return false; }

	ResultType STDMETHODCALLTYPE Invoke(ExprTokenType &aResultToken, ExprTokenType &aThisToken, int aFlags, ExprTokenType *aParam[], int aParamCount);
};

extern MetaObject g_MetaObject;		// Defines "object" behaviour for non-object values.


//
// RegExMatchObject:  Returned by RegExMatch via UnquotedOutputVar.
//
class RegExMatchObject : public ObjectBase
{
	LPTSTR mHaystack;
	int *mOffset;
	LPTSTR *mPatternName;
	int mPatternCount;

	RegExMatchObject() : mHaystack(NULL), mOffset(NULL), mPatternName(NULL), mPatternCount(0) {}
	
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
	}

public:
	static RegExMatchObject *Create(LPCTSTR aHaystack, int *aOffset, LPCTSTR *aPatternName
		, int aPatternCount, int aCapturedPatternCount);
	
	ResultType STDMETHODCALLTYPE Invoke(ExprTokenType &aResultToken, ExprTokenType &aThisToken, int aFlags, ExprTokenType *aParam[], int aParamCount);
};