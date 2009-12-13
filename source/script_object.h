
#define IT_GET				0
#define IT_SET				1
#define IT_CALL				2 // L40: MetaObject::Invoke relies on these being mutually-exclusive bits.
#define IT_BITMASK			3 // bit-mask for the above.

#define IF_METAOBJ			0x10000 // Indicates 'this' is a meta-object/base of aThisToken. Restricts some functionality and causes aThisToken to be inserted into the param list of called functions.
#define IF_METAFUNC			0x20000 // Indicates Invoke should call a meta-function before checking the object's fields.
#define IF_META				(IF_METAOBJ | IF_METAFUNC)	// Flags for regular recursion into base object.

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

	virtual bool Delete()
	{
		delete this; // Derived classes MUST be instantiated with 'new' or override this function.
		return true; // See Release() for comments.
	}

	// Declare a virtual destructor for correct 'delete this' behaviour in Delete(),
	// and because it is likely to be more convenient and reliable than overriding
	// Delete(), especially with a chain of derived types.
	virtual ~ObjectBase() {}
};	
	

//
// Object - Scriptable associative array.
//

class Object : public ObjectBase
{
protected:
	union KeyType // Which of its members is used depends on the field's position in the mFields array.
	{
		char *s;
		int i;
		IObject *p;
	};
	struct FieldType
	{
		union { // Which of its members is used depends on the value of symbol, below.
			__int64 n_int64;	// for SYM_INTEGER
			double n_double;	// for SYM_FLOAT
			IObject *object;	// for SYM_OBJECT
			struct {
				char *marker;		// for SYM_OPERAND
				size_t size;		// for SYM_OPERAND; allows reuse of allocated memory.
			};
		};
		// key and symbol probably need to be adjacent to each other to conserve memory due to 8-byte alignment.
		KeyType key;
		SymbolType symbol;
		
		inline int CompareKey(int val) { return val - key.i; }  // Used by both int and object since they are stored separately.  Will needs review for 64-bit compatibility.
		inline int CompareKey(char *val) { return stricmp(val, key.s); }
		
		bool Assign(char *str, size_t len = -1, bool exact_size = false);
		bool Assign(ExprTokenType &val);
		void Get(ExprTokenType &result);
		void Free();
	};
	
	IObject *mBase;
	FieldType *mFields;
	int mFieldCount, mFieldCountMax; // Current/max number of fields.

	// Holds the index of first key of a given type within mFields.  Must be in the order: int, object, string.
	// Compared to an approach using true key-value pairs, this approach saves 8 bytes per key (excluding
	// the 8 bytes taken by the two fields below) and speeds up lookups since only the section within mFields
	// with the appropriate type of key needs to be searched (and no need to check the type of each key).
	// mKeyOffsetObject should be set to mKeyOffsetInt + the number of int keys.
	// mKeyOffsetString should be set to mKeyOffsetObject + the number of object keys.
	// mKeyOffsetObject-1, mKeyOffsetString-1 and mFieldCount-1 indicate the last index of each prior type.
	static const int mKeyOffsetInt = 0;
	int mKeyOffsetObject, mKeyOffsetString;

	Object()
		: mBase(NULL)
		, mFields(NULL), mFieldCount(0), mFieldCountMax(0)
		, mKeyOffsetObject(0), mKeyOffsetString(0)
	{}

	bool Delete();
	~Object();

	template<typename T>
	FieldType *FindField(T val, int left, int right, int &insert_pos);
	FieldType *FindField(SymbolType key_type, KeyType key, int &insert_pos);	
	FieldType *FindField(ExprTokenType &key_token, SymbolType &key_type, KeyType &key, int &insert_pos);
	FieldType *FindField(ExprTokenType &key_token)
	// Overload for callers who will not need to later insert a field.
	{
		SymbolType key_type;
		KeyType key;
		int insert_pos;
		return FindField(key_token, key_type, key, insert_pos);
	}
	
	FieldType *Insert(SymbolType key_type, KeyType key, int at);

	bool SetInternalCapacity(int new_capacity);
	bool Expand()
	// Expands mFields by at least one field.
	{
		return SetInternalCapacity(mFieldCountMax ? mFieldCountMax * 2 : 4);
	}
	
	inline ResultType _Insert(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);
	inline ResultType _Remove(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);
	inline ResultType _GetCapacity(ExprTokenType &aResultToken);
	inline ResultType _SetCapacity(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);
	inline ResultType _MaxIndex(ExprTokenType &aResultToken);
	inline ResultType _MinIndex(ExprTokenType &aResultToken);

	ResultType CallField(FieldType *aField, ExprTokenType &aResultToken, ExprTokenType &aThisToken, int aFlags, ExprTokenType *aParam[], int aParamCount);
	
public:
	ResultType STDMETHODCALLTYPE Invoke(ExprTokenType &aResultToken, ExprTokenType &aThisToken, int aFlags, ExprTokenType *aParam[], int aParamCount);
	static IObject *Create(ExprTokenType *aParam[], int aParamCount);

	static char *sMetaFuncName[];
};


//
// MetaObject:	Used only by g_MetaObject (not every meta-object); see comments below.
//
class MetaObject : public Object
{
public:
	// In addition to ensuring g_MetaObject is never "deleted", this avoids a
	// tiny bit of work when any reference to this object is added or released.
	// Temporary references such as when evaluting "".base.foo are most common.
	ULONG STDMETHODCALLTYPE AddRef() { return 1; }
	ULONG STDMETHODCALLTYPE Release() { return 1; }
	bool Delete() { return false; }

	//ResultType STDMETHODCALLTYPE Invoke(ExprTokenType &aResultToken, ExprTokenType &aThisToken, int aFlags, ExprTokenType *aParam[], int aParamCount);
};

extern MetaObject g_MetaObject;		// Defines "object" behaviour for non-object values.

