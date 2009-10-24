
#define IT_GET				0
#define IT_SET				1
#define IT_CALL				2
#define IT_BITMASK			3 // bit-mask for the above.

#define IF_META				0x10000 // meta-invocation: 'this' is a meta-object/base of aThisToken.
#define IF_METAFUNC			0x20000 // meta-invocation of __Call/__Get/__Set.

#define IT_CALL_METAFUNC	(IT_CALL | IF_META | IF_METAFUNC)

#define INVOKE_TYPE			(aFlags & IT_BITMASK)
#define IS_INVOKE_SET		(aFlags & IT_SET)
#define IS_INVOKE_GET		(INVOKE_TYPE == IT_GET)
#define IS_INVOKE_CALL		(aFlags & IT_CALL)
#define IS_INVOKE_META		(aFlags & IF_META)

#define INVOKE_NOT_HANDLED	CONDITION_FALSE


ExprTokenType g_MetaFuncId[4];	// Initialized by MetaObject constructor.


//
// ObjectBase
//	Common base class, implements reference counting.
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

		// Used by both int and object since they are stored separately.  NOT 64-BIT COMPATIBLE (many other areas would need to be revised anyway).
		inline int CompareKey(int val)
		{
			return val - key.i;
		}

		inline int CompareKey(char *val)
		{
			return stricmp(val, key.s);
		}
		
		bool Assign(char *str, size_t len = -1, bool exact_size = false)
		{
			if (!str || !*str && len < 1) // If empty string or null pointer, free our contents.  Passing len >= 1 allows copying \0, so don't check *str in that case.  Ordered for short-circuit performance (len is usually -1).
			{
				Free();
				marker = Var::sEmptyString;
				size = 0;
				return false;
			}
			
			if (len == -1)
				len = strlen(str);

			if (symbol != SYM_OPERAND || len >= size)
			{
				Free(); // Free object or previous buffer (which was too small).
				symbol = SYM_OPERAND;
				int new_size = len + 1;
				if (!exact_size)
				{
					// Use size calculations equivalent to Var:
					if (new_size < 16) // v1.0.45.03: Added this new size to prevent all local variables in a recursive
						new_size = 16; // function from having a minimum size of MAX_PATH.  16 seems like a good size because it holds nearly any number.  It seems counterproductive to go too small because each malloc, no matter how small, could have around 40 bytes of overhead.
					else if (new_size < MAX_PATH)
						new_size = MAX_PATH;  // An amount that will fit all standard filenames seems good.
					else if (new_size < (160 * 1024)) // MAX_PATH to 160 KB or less -> 10% extra.
						new_size = (size_t)(new_size * 1.1);
					else if (new_size < (1600 * 1024))  // 160 to 1600 KB -> 16 KB extra
						new_size += (16 * 1024);
					else if (new_size < (6400 * 1024)) // 1600 to 6400 KB -> 1% extra
						new_size = (size_t)(new_size * 1.01);
					else  // 6400 KB or more: Cap the extra margin at some reasonable compromise of speed vs. mem usage: 64 KB
						new_size += (64 * 1024);
				}
				if ( !(marker = (char *)malloc(new_size)) )
				{
					marker = Var::sEmptyString;
					size = 0;
					return false; // See "Sanity check" above.
				}
				size = new_size;
			}
			// else we have a buffer with sufficient capacity already.

			memcpy(marker, str, len + 1); // +1 for null-terminator.
			return true; // Success.
		}

		bool Assign(ExprTokenType &val)
		{
			// Currently only SYM_INTEGER/SYM_FLOAT inputs are stored as binary numbers
			// since it seems best to preserve formatting of SYM_OPERAND/SYM_VAR (in case
			// it is important), at the cost of performance in *some* cases.
			if (IS_NUMERIC(val.symbol))
			{
				Free(); // Free string or object, if applicable.
				symbol = val.symbol; // Either SYM_INTEGER or SYM_FLOAT.  Set symbol *after* calling Free().
				n_int64 = val.value_int64; // Also handles value_double via union.
			}
			else
			{
				// String, object or var (can be a numeric string or var with cached binary number; see above).
				IObject *val_as_obj;
				if (val_as_obj = TokenToObject(val)) // SYM_OBJECT or SYM_VAR with var containing object.
				{
					Free(); // Free string or object, if applicable.
					val_as_obj->AddRef();
					symbol = SYM_OBJECT; // Set symbol *after* calling Free().
					object = val_as_obj;
				}
				else
				{
					// Handles setting symbol and allocating or resizing buffer as appropriate:
					return Assign(TokenToString(val));
				}
			}
			return true;
		}

		void Get(ExprTokenType &result)
		{
			result.symbol = symbol;
			result.value_int64 = n_int64; // Union copy.
			if (symbol == SYM_OBJECT)
				object->AddRef();
		}

		void Free()
		// Only the value is freed, since keys only need to be freed when a field is removed
		// entirely or the Object is being deleted.  See Object::Delete.
		// CONTAINED VALUE WILL NOT BE VALID AFTER THIS FUNCTION RETURNS.
		{
			if (symbol == SYM_OPERAND) {
				if (size)
					free(marker);
			} else if (symbol == SYM_OBJECT)
				object->Release();
		}
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
	int /*mKeyOffsetInt,*/ mKeyOffsetObject, mKeyOffsetString;

	Object()
		: mBase(NULL)
		, mFields(NULL), mFieldCount(0), mFieldCountMax(0)
		, mKeyOffsetObject(0), mKeyOffsetString(0)
	{}

	bool Delete(); // Called immediately before mRefCount is decremented to 0.
	
	~Object()
	{
		if (mBase)
			mBase->Release();

		if (mFields)
		{
			int i = mFieldCount - 1;
			// Free keys: first strings, then objects (objects have a lower index in the mFields array).
			for ( ; i >= mKeyOffsetString; --i)
				free(mFields[i].key.s);
			for ( ; i >= mKeyOffsetObject; --i)
				mFields[i].key.p->Release();
			// Free values.
			while (mFieldCount) 
				mFields[--mFieldCount].Free();
			// Free fields array.
			free(mFields);
		}
	}


	inline SymbolType TokenToKey(ExprTokenType &key_token, KeyType &key)
	// Takes a token and outputs a key value (of type indicated by return value) and begin/end offsets where
	// keys of this type are stored in the mFields array.  inline because currently used in only one place.
	{
		if (TokenIsPureNumeric(key_token) == PURE_INTEGER)
		{
			// Treat all integer keys (even numeric strings) as pure integers for consistency and performance.
			key.i = (int)TokenToInt64(key_token, TRUE);
			return SYM_INTEGER;
		}
		else if (key.p = TokenToObject(key_token))
		{
			// SYM_OBJECT or SYM_VAR containing object.
			return SYM_OBJECT;
		}
		// SYM_STRING, SYM_OPERAND or SYM_VAR (all confirmed not to be an integer at this point).
		key.s = TokenToString(key_token);
		return SYM_STRING;
	}

	template<typename T> inline FieldType *FindField(T val, int left, int right, int &insert_pos)
	// left and right must be set by caller to the appropriate bounds within mFields.
	{
		int mid, result;
		while (left <= right)
		{
			mid = (left + right) / 2;
			
			FieldType &field = mFields[mid];
			
			result = field.CompareKey(val);
			
			if (result < 0)
				right = mid - 1;
			else if (result > 0)
				left = mid + 1;
			else
				return &field;
		}
		insert_pos = left;
		return NULL;
	}

	FieldType *FindField(SymbolType key_type, KeyType key, int &insert_pos)
	// Searches for a field with the given key.  If found, a pointer to the field is returned.  Otherwise
	// NULL is returned and insert_pos is set to the index a newly created field should be inserted at.
	// key_type and key are output for creating a new field or removing an existing one correctly.
	// left and right must indicate the appropriate section of mFields to search, based on key type.
	{
		int left, right;

		if (key_type == SYM_STRING)
		{
			left = mKeyOffsetString;
			right = mFieldCount - 1; // String keys are last in the mFields array.

			return FindField<char *>(key.s, left, right, insert_pos);
		}
		else // key_type == SYM_INTEGER || key_type == SYM_OBJECT
		{
			if (key_type == SYM_INTEGER)
			{
				left = mKeyOffsetInt;
				right = mKeyOffsetObject - 1; // Int keys end where Object keys begin.
			}
			else
			{
				left = mKeyOffsetObject;
				right = mKeyOffsetString - 1; // Object keys end where String keys begin.
			}
			// Both may be treated as integer since left/right exclude keys of an incorrect type:
			return FindField<int>(key.i, left, right, insert_pos);
		}
	}

	FieldType *FindField(ExprTokenType &key_token, SymbolType &key_type, KeyType &key, int &insert_pos)
	{
		key_type = TokenToKey(key_token, key);
		return FindField(key_type, key, insert_pos);
	}

	FieldType *FindField(ExprTokenType &key_token)
	// Overload for callers who will not need to later insert a field.
	{
		SymbolType key_type;
		KeyType key;
		int insert_pos;
		return FindField(key_token, key_type, key, insert_pos);
	}
	
	bool Expand()
	// Expands mFields by at least one field.
	{
		return SetInternalCapacity(mFieldCountMax ? mFieldCountMax * 2 : 4);
	}

	bool SetInternalCapacity(int new_capacity)
	// Expands mFields to the specified number if fields.
	// Caller *must* ensure new_capacity >= 1 && new_capacity >= mFieldCount.
	{
		FieldType *new_fields = (FieldType *)realloc(mFields, new_capacity * sizeof(FieldType));
		if (!new_fields)
			return false;
		mFields = new_fields;
		mFieldCountMax = new_capacity;
		return true;
	}

	FieldType *Insert(SymbolType key_type, KeyType key, int at)
	// Inserts a single field with the given key at the given offset.
	// Caller must ensure 'at' is the correct offset for this key.
	{
		if (mFieldCount == mFieldCountMax && !Expand()  // Attempt to expand if at capacity.
			|| key_type == SYM_STRING && !(key.s = _strdup(key.s)))  // Attempt to duplicate key-string.
		{	// Out of memory.
			return NULL;
		}
		// There is now definitely room in mFields for a new field.

		if (key_type == SYM_OBJECT)
			// Keep key object alive:
			key.p->AddRef();

		FieldType &field = mFields[at];
		if (at < mFieldCount)
			// Move existing fields to make room.
			memmove(&field + 1, &field, (mFieldCount - at) * sizeof(FieldType));
		
		// Since we just inserted a field, we must update the key type offsets:
		if (key_type != SYM_STRING)
		{
			// Must be either SYM_INTEGER or SYM_OBJECT, which both precede SYM_STRING.
			++mKeyOffsetString;
			if (key_type != SYM_OBJECT)
				// Must be SYM_INTEGER, which precedes SYM_OBJECT.
				++mKeyOffsetObject;
		}
		++mFieldCount; // Only after memmove above.
		
		field.marker = ""; // Init for maintainability.
		field.size = 0; // Init to ensure safe behaviour in Assign().
		field.key = key; // Above has already copied string or called key.p->AddRef() as appropriate.
		field.symbol = SYM_OPERAND;

		return &field;
	}

	void Remove(FieldType *field, SymbolType key_type)
	// Removes a given field.
	// Caller must ensure 'field' is a location within mFields.
	{
		int remaining_fields = &mFields[mFieldCount-1] - field;
		if (remaining_fields)
			memmove(field, field + 1, remaining_fields * sizeof(FieldType));
		--mFieldCount;
		// Also update key type offsets as necessary -- see related section above for comments:
		if (key_type != SYM_STRING)
		{
			--mKeyOffsetString;
			if (key_type != SYM_OBJECT)
				--mKeyOffsetObject;
		}
	}

	inline ResultType _Insert(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);			// _Insert( key, value )
	inline ResultType _Remove(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);			// _Remove( min_key [, max_key ] )
	inline ResultType _GetCapacity(ExprTokenType &aResultToken);//, ExprTokenType *aParam[], int aParamCount);
	inline ResultType _SetCapacity(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);		// _SetCapacity( new_capacity )
	inline ResultType _MaxIndex(ExprTokenType &aResultToken);
	inline ResultType _MinIndex(ExprTokenType &aResultToken);
	//inline ResultType _GetAddress(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);
	
public:
	ResultType STDMETHODCALLTYPE Invoke(ExprTokenType &aResultToken, ExprTokenType &aThisToken, int aFlags, ExprTokenType *aParam[], int aParamCount);
	static IObject *Create(ExprTokenType *aParam[], int aParamCount);
};


//
// MetaObject:	Used only by g_MetaObject (not every meta-object); see comments below.
//
class MetaObject : public Object
{
public:
	// In addition to ensuring g_MetaObject is never deleted, this avoids a tiny bit of work when
	// any object with mBase == g_MetaObject is deleted (and it calls mBase->Release()).
	// Since every object has mBase == g_MetaObject by default, it seems worth doing.
	ULONG STDMETHODCALLTYPE AddRef() { return 1; }
	ULONG STDMETHODCALLTYPE Release() { return 1; }
	bool Delete() { return false; }

	MetaObject()
	{
		// Initialize array of meta-function "identifier" tokens for quick access.
		g_MetaFuncId[IT_GET].symbol = SYM_STRING;
		g_MetaFuncId[IT_GET].marker = "__Get";
		g_MetaFuncId[IT_SET].symbol = SYM_STRING;
		g_MetaFuncId[IT_SET].marker = "__Set";
		g_MetaFuncId[IT_CALL].symbol = SYM_STRING;
		g_MetaFuncId[IT_CALL].marker = "__Call";
		g_MetaFuncId[3].symbol = SYM_STRING; // This one does not have a corresponding Invoke-type since it is only invoked indirectly via Release().
		g_MetaFuncId[3].marker = "__Delete";
	}

	ResultType STDMETHODCALLTYPE Invoke(ExprTokenType &aResultToken, ExprTokenType &aThisToken, int aFlags, ExprTokenType *aParam[], int aParamCount);
};

MetaObject g_MetaObject;		// Defines "object" behaviour for non-object values.

