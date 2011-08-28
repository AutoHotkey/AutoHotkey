#include "stdafx.h" // pre-compiled headers
#include "defines.h"
#include "globaldata.h"
#include "script.h"

#include "script_object.h"


//
//	Internal: CallFunc - Call a script function with given params.
//

ResultType CallFunc(Func &aFunc, ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
// Caller should pass an aResultToken with the usual setup:
//	buf points to a buffer the called function may use: TCHAR[MAX_NUMBER_SIZE]
//	mem_to_free is NULL; if it is non-NULL on return, caller (or caller's caller) is responsible for it.
// Caller is responsible for making a persistent copy of the result, if appropriate.
{
	if (aParamCount < aFunc.mMinParams)
	{
		aResultToken.symbol = SYM_STRING;
		aResultToken.marker = _T("");
		return FAIL;
	}

	// When this variable goes out of scope, Var::FreeAndRestoreFunctionVars() is called (if appropriate):
	FuncCallData func_call;
	ResultType result;

	// CALL THE FUNCTION.
	if (aFunc.Call(func_call, result, aResultToken, aParam, aParamCount)
		// Make return value persistent if applicable:
		&& aResultToken.symbol == SYM_STRING && !aFunc.mIsBuiltIn)
	{
		// Make a persistent copy of the string in case it is the contents of one of the function's local variables.
		if ( !*aResultToken.marker || !TokenSetResult(aResultToken, aResultToken.marker) )
			// Above failed or the result is an empty string, so make sure it remains valid after we return:
			aResultToken.marker = _T("");
	}

	return result;
}
	

//
// Object::Create - Called by BIF_ObjCreate to create a new object, optionally passing key/value pairs to set.
//

Object *Object::Create(ExprTokenType *aParam[], int aParamCount)
{
	if (aParamCount & 1)
		return NULL; // Odd number of parameters - reserved for future use.

	Object *obj = new Object();
	if (obj && aParamCount)
	{
		ExprTokenType result_token, this_token;
		TCHAR buf[MAX_NUMBER_SIZE];

		this_token.symbol = SYM_OBJECT;
		this_token.object = obj;
		
		for (int i = 0; i + 1 < aParamCount; i += 2)
		{
			result_token.symbol = SYM_STRING;
			result_token.marker = _T("");
			result_token.mem_to_free = NULL;
			result_token.buf = buf;

			// This is used rather than a more direct approach to ensure it is equivalent to assignment.
			// For instance, Object("base",MyBase,"a",1,"b",2) invokes meta-functions contained by MyBase.
			// For future consideration: Maybe it *should* bypass the meta-mechanism?
			obj->Invoke(result_token, this_token, IT_SET, aParam + i, 2);

			if (result_token.symbol == SYM_OBJECT) // L33: Bugfix.  Invoke must assume the result will be used and as a result we must account for this object reference:
				result_token.object->Release();
			if (result_token.mem_to_free) // Comment may be obsolete: Currently should never happen, but may happen in future.
				free(result_token.mem_to_free);
		}
	}
	return obj;
}


//
// Object::Clone - Used for variadic function-calls.
//

Object *Object::Clone(INT_PTR aStartOffset)
// Creates an object and copies to it the fields at and after the given offset.
{
	Object *objptr = new Object();
	if (!objptr|| aStartOffset >= mFieldCount)
		return objptr;
	
	Object &obj = *objptr;

	// Allocate space in destination object.
	IndexType field_count = mFieldCount - aStartOffset;
	if (!obj.SetInternalCapacity(field_count))
	{
		obj.Release();
		return NULL;
	}

	FieldType *fields = obj.mFields; // Newly allocated by above.
	int failure_count = 0; // See comment below.
	IndexType i;

	obj.mFieldCount = field_count;
	obj.mKeyOffsetObject = mKeyOffsetObject - aStartOffset;
	obj.mKeyOffsetString = mKeyOffsetString - aStartOffset;
	if (obj.mKeyOffsetObject < 0) // Currently might always evaluate to false.
	{
		obj.mKeyOffsetObject = 0; // aStartOffset excluded all integer and some or all object keys.
		if (obj.mKeyOffsetString < 0)
			obj.mKeyOffsetString = 0; // aStartOffset also excluded some string keys.
	}
	//else no need to check mKeyOffsetString since it should always be >= mKeyOffsetObject.

	for (i = 0; i < field_count; ++i)
	{
		FieldType &dst = fields[i];
		FieldType &src = mFields[aStartOffset + i];

		// Copy key.
		if (i >= obj.mKeyOffsetString)
		{
			if ( !(dst.key.s = _tcsdup(src.key.s)) )
			{
				// Key allocation failed. At this point, all int and object keys
				// have been set and values for previous fields have been copied.
				// Rather than trying to set up the object so that what we have
				// so far is valid in order to break out of the loop, continue,
				// make all fields valid and then allow them to be freed. 
				++failure_count;
			}
		}
		else if (i >= obj.mKeyOffsetObject)
			(dst.key.p = src.key.p)->AddRef();
		else
			dst.key.i = src.key.i;

		// Copy value.
		switch (dst.symbol = src.symbol)
		{
		case SYM_OPERAND:
			if (dst.size = src.size)
			{
				if (dst.marker = tmalloc(dst.size))
				{
					// Since user may have stored binary data, copy the entire field:
					tmemcpy(dst.marker, src.marker, src.size);
					continue;
				}
				// Since above didn't continue: allocation failed.
				++failure_count; // See failure comment further above.
			}
			dst.marker = Var::sEmptyString;
			dst.size = 0;
			break;

		case SYM_OBJECT:
			(dst.object = src.object)->AddRef();
			break;

		//case SYM_INTEGER:
		//case SYM_FLOAT:
		default:
			dst.n_int64 = src.n_int64; // Union copy.
		}
	}
	if (failure_count)
	{
		// One or more memory allocations failed.  It seems best to return a clear failure
		// indication rather than an incomplete copy.  Now that the loop above has finished,
		// the object's contents are at least valid and it is safe to free the object:
		obj.Release();
		return NULL;
	}
	return &obj;
}


//
// Object::ArrayToParams - Used for variadic function-calls.
//

ResultType Object::ArrayToParams(void *&aMemToFree, ExprTokenType **&aParam, int &aParamCount, int aMinParams)
// Expands this object's contents into the parameter list. Due to the nature
// of the parameter list, this function has the following limitations:
//	- Only fields with integer keys are used (named parameters aren't supported).
//	- Keys must be contiguous; if keys are non-contiguous, there must be a parameter
//		missing somewhere.  We could assume an empty string for omitted parameters,
//		but that would probably be counter-intuitive and scripts which come to rely
//		on it would break if this limitation is ever removed.
// RETURN VALUE: true if new memory was allocated for aParam; false otherwise.
{
	aMemToFree = NULL; // Init for maintainability.
	int field_count = (int)(mKeyOffsetObject - mKeyOffsetInt);
	if (aParamCount + field_count < aMinParams)
		return FAIL; // There still won't be enough params, so give up.  Caller relies on us to do this check.
	if (!field_count)
		return OK; // Nothing to expand into the param list, so our work here is done.
	FieldType *field = mFields + mKeyOffsetInt;
	if (field->key.i != 1 || field[field_count-1].key.i != (IndexType)field_count)
		// Keys are non-contiguous or don't start at 1.  See comments above.
		return FAIL;
	// Calculate space required for ...
	size_t space_needed = field_count * sizeof(ExprTokenType) // ... new param tokens
		+ (aParamCount + field_count) * sizeof(ExprTokenType *); // ... existing and new param pointers
	// Allocate new param list and tokens; tokens first for convenience.
	ExprTokenType *token = (ExprTokenType *)malloc(space_needed);
	if (!token)
		return FAIL;
	ExprTokenType **param_list = (ExprTokenType **)(token + field_count);
	memcpy(param_list, aParam, aParamCount * sizeof(ExprTokenType *)); // Copy the existing param pointers.
	aParam = param_list; // Update caller's pointer.
	param_list += aParamCount; // Point param_list at our target; past the existing params.
	// Convert each field to a token, putting the token's address into the appropriate slot:
	for (int i = 0; i < field_count; ++i)
		field[i].ToToken(*(param_list[i] = &token[i]));
	aParamCount += field_count; // Update caller's count.
	aMemToFree = token; // Return the memory block for caller to take care of.
	return OK;
}


//
// Object::Delete - Called immediately before the object is deleted.
//					Returns false if object should not be deleted yet.
//

bool Object::Delete()
{
	if (mBase)
	{
		KeyType key;
		IndexType insert_pos;
		key.s = _T("__Class");
		if (FindField(SYM_STRING, key, insert_pos))
			// This object appears to be a class definition, so it would probably be
			// undesirable to call the super-class' __Delete() meta-function for this.
			return ObjectBase::Delete();

		ExprTokenType result_token, this_token, param_token, *param;
		
		TCHAR dummy_buf[MAX_NUMBER_SIZE];
		result_token.marker = _T("");
		result_token.symbol = SYM_STRING;
		result_token.mem_to_free = NULL;
		result_token.buf = dummy_buf; // This is required in the case where __Delete returns a short string (although there's no good reason to ever do that).

		this_token.symbol = SYM_OBJECT;
		this_token.object = this;

		param_token.symbol = SYM_STRING;
		param_token.marker = sMetaFuncName[3]; // "__Delete"
		param = &param_token;

		// L33: Privatize the last recursion layer's deref buffer in case it is in use by our caller.
		// It's done here rather than in Var::FreeAndRestoreFunctionVars or CallFunc (even though the
		// below might not actually call any script functions) because this function is probably
		// executed much less often in most cases.
		PRIVATIZE_S_DEREF_BUF;

		mBase->Invoke(result_token, this_token, IT_CALL | IF_METAOBJ, &param, 1);

		DEPRIVATIZE_S_DEREF_BUF; // L33: See above.

		// L33: Release result if given, although typically there should not be one:
		if (result_token.mem_to_free)
			free(result_token.mem_to_free);
		if (result_token.symbol == SYM_OBJECT)
			result_token.object->Release();

		// Above may pass the script a reference to this object to allow cleanup routines to free any
		// associated resources.  Deleting it is only safe if the script no longer holds any references
		// to it.  Since cleanup routines may (intentionally or unintentionally) copy this reference,
		// ensure this object really has no more references before proceeding with deletion:
		if (mRefCount > 1)
			return false;
	}
	return ObjectBase::Delete();
}


Object::~Object()
{
	if (mBase)
		mBase->Release();

	if (mFields)
	{
		if (mFieldCount)
		{
			IndexType i = mFieldCount - 1;
			// Free keys: first strings, then objects (objects have a lower index in the mFields array).
			for ( ; i >= mKeyOffsetString; --i)
				free(mFields[i].key.s);
			for ( ; i >= mKeyOffsetObject; --i)
				mFields[i].key.p->Release();
			// Free values.
			while (mFieldCount) 
				mFields[--mFieldCount].Free();
		}
		// Free fields array.
		free(mFields);
	}
}


//
// Object::Invoke - Called by BIF_ObjInvoke when script explicitly interacts with an object.
//

ResultType STDMETHODCALLTYPE Object::Invoke(
                                            ExprTokenType &aResultToken,
                                            ExprTokenType &aThisToken,
                                            int aFlags,
                                            ExprTokenType *aParam[],
                                            int aParamCount
                                            )
// L40: Revised base mechanism for flexibility and to simplify some aspects.
//		obj[] -> obj.base.__Get -> obj.base[] -> obj.base.__Get etc.
{
	SymbolType key_type;
	KeyType key;
    FieldType *field;
	IndexType insert_pos;

	// If this is some object's base and is being invoked in that capacity, call
	//	__Get/__Set/__Call as defined in this base object before searching further.
	if (SHOULD_INVOKE_METAFUNC)
	{
		key.s = sMetaFuncName[INVOKE_TYPE];
		// Look for a meta-function definition directly in this base object.
		if (field = FindField(SYM_STRING, key, /*out*/ insert_pos))
		{
			// Seems more maintainable to copy params rather than assume aParam[-1] is always valid.
			ExprTokenType **meta_params = (ExprTokenType **)_alloca((aParamCount + 1) * sizeof(ExprTokenType *));
			// Shallow copy; points to the same tokens.  Leave a space for param[0], which must be the param which
			// identified the field (or in this case an empty space) to replace with aThisToken when appropriate.
			memcpy(meta_params + 1, aParam, aParamCount * sizeof(ExprTokenType*));

			ResultType r = CallField(field, aResultToken, aThisToken, aFlags, meta_params, aParamCount + 1);
			if (r == EARLY_RETURN)
				// Propagate EARLY_RETURN in case this was the __Call meta-function of a
				// "function object" which is used as a meta-function of some other object.
				return EARLY_RETURN; // TODO: Detection of 'return' vs 'return empty_value'.
		}
	}
	
	int param_count_excluding_rvalue = aParamCount;

	if (IS_INVOKE_SET)
	{
		// Prior validation of ObjSet() param count ensures the result won't be negative:
		--param_count_excluding_rvalue;
		
		if (IS_INVOKE_META)
		{
			if (param_count_excluding_rvalue == 1)
				// Prevent below from searching for or setting a field, since this is a base object of aThisToken.
				// Relies on mBase->Invoke recursion using aParamCount and not param_count_excluding_rvalue.
				param_count_excluding_rvalue = 0;
			//else: Allow SET to operate on a field of an object stored in the target's base.
			//		For instance, x[y,z]:=w may operate on x[y][z], x.base[y][z], x[y].base[z], etc.
		}
	}
	
	if (param_count_excluding_rvalue)
	{
		field = FindField(*aParam[0], aResultToken.buf, /*out*/ key_type, /*out*/ key, /*out*/ insert_pos);
	}
	else
	{
		key_type = SYM_INVALID; // Allow key_type checks below without requiring that param_count_excluding_rvalue also be checked.
		field = NULL;
	}
	
	if (!field)
	{
		// This field doesn't exist, so let our base object define what happens:
		//		1) __Get, __Set or __Call.  If these don't return a value, processing continues.
		//		2) For GET and CALL only, check the base object's own fields.
		//		3) Repeat 1 through 3 for the base object's own base.
		if (mBase)
		{
			// aFlags: If caller specified IF_METAOBJ but not IF_METAFUNC, they want to recursively
			// find and execute a specific meta-function (__new or __delete) but don't want any base
			// object to invoke __call.  So if this is already a meta-invocation, don't change aFlags.
			ResultType r = mBase->Invoke(aResultToken, aThisToken, aFlags | (IS_INVOKE_META ? 0 : IF_META), aParam, aParamCount);
			if (r != INVOKE_NOT_HANDLED)
				return r;

			// Since the above may have inserted or removed fields (including the specified one),
			// insert_pos may no longer be correct or safe.  Updating field also allows a meta-function
			// to initialize a field and allow processing to continue as if it already existed.
			if (param_count_excluding_rvalue)
				field = FindField(key_type, key, /*out*/ insert_pos);
		}

		// Since the base object didn't handle this op, check for built-in properties/methods.
		// This must apply only to the original target object (aThisToken), not one of its bases.
		if (!IS_INVOKE_META && key_type == SYM_STRING)
		{
			//
			// BUILT-IN METHODS
			//
			if (IS_INVOKE_CALL)
			{
				// Since above has not handled this call and no field exists, check for built-in methods.
				LPTSTR name = key.s;
				if (*name == '_')
					++name; // ++ to exclude '_' from further consideration.
				++aParam; --aParamCount; // Exclude the method identifier.  A prior check ensures there was at least one param in this case.
				if (!_tcsicmp(name, _T("Insert")))
					return _Insert(aResultToken, aParam, aParamCount);
				if (!_tcsicmp(name, _T("Remove")))
					return _Remove(aResultToken, aParam, aParamCount);
				if (!_tcsicmp(name, _T("HasKey")))
					return _HasKey(aResultToken, aParam, aParamCount);
				if (!_tcsicmp(name, _T("MaxIndex")))
					return _MaxIndex(aResultToken, aParam, aParamCount);
				if (!_tcsicmp(name, _T("NewEnum")))
					return _NewEnum(aResultToken, aParam, aParamCount);
				if (!_tcsicmp(name, _T("GetAddress")))
					return _GetAddress(aResultToken, aParam, aParamCount);
				if (!_tcsicmp(name, _T("SetCapacity")))
					return _SetCapacity(aResultToken, aParam, aParamCount);
				if (!_tcsicmp(name, _T("GetCapacity")))
					return _GetCapacity(aResultToken, aParam, aParamCount);
				if (!_tcsicmp(name, _T("MinIndex")))
					return _MinIndex(aResultToken, aParam, aParamCount);
				if (!_tcsicmp(name, _T("Clone")))
					return _Clone(aResultToken, aParam, aParamCount);
				// For maintainability: explicitly return since above has done ++aParam, --aParamCount.
				return INVOKE_NOT_HANDLED;
			}
			//
			// BUILT-IN "BASE" PROPERTY
			//
			else if (param_count_excluding_rvalue == 1 && !_tcsicmp(key.s, _T("base")))
			{
				if (IS_INVOKE_SET)
				// "base" must be handled before inserting a new field.
				{
					IObject *obj = TokenToObject(*aParam[1]);
					if (obj)
					{
						obj->AddRef(); // for mBase
						obj->AddRef(); // for aResultToken
						aResultToken.symbol = SYM_OBJECT;
						aResultToken.object = obj;
					}
					// else leave as empty string.
					if (mBase)
						mBase->Release();
					mBase = obj; // May be NULL.
					return OK;
				}
				else // GET
				{
					if (mBase)
					{
						aResultToken.symbol = SYM_OBJECT;
						aResultToken.object = mBase;
						mBase->AddRef();
					}
					// else leave as empty string.
					return OK;
				}
			}
		} // if (!IS_INVOKE_META && key_type == SYM_STRING)
	} // if (!field)

	//
	// OPERATE ON A FIELD WITHIN THIS OBJECT
	//

	// CALL
	if (IS_INVOKE_CALL)
	{
		if (field)
			return CallField(field, aResultToken, aThisToken, aFlags, aParam, aParamCount);
	}

	// MULTIPARAM[x,y] -- may be SET[x,y]:=z or GET[x,y], but always treated like GET[x].
	else if (param_count_excluding_rvalue > 1)
	{
		// This is something like this[x,y] or this[x,y]:=z.  Since it wasn't handled by a meta-mechanism above,
		// handle only the "x" part (automatically creating and storing an object if this[x] didn't already exist
		// and an assignment is being made) and recursively invoke.  This has at least two benefits:
		//	1) Objects natively function as multi-dimensional arrays.
		//	2) Automatic initialization of object-fields.
		//		For instance, this["base","__Get"]:="MyObjGet" does not require a prior this.base:=Object().
		IObject *obj = NULL;
		if (field)
		{
			if (field->symbol == SYM_OBJECT)
				// AddRef not used.  See below.
				obj = field->object;
		}
		else if (!IS_INVOKE_META)
		{
			// This section applies only to the target object (aThisToken) and not any of its base objects.
			// Allow obj["base",x] to access a field of obj.base; L40: This also fixes obj.base[x] which was broken by L36.
			if (key_type == SYM_STRING && !_tcsicmp(key.s, _T("base")))
			{
				if (!mBase && IS_INVOKE_SET)
					mBase = new Object();
				obj = mBase; // If NULL, above failed and below will detect it.
			}
			// Automatically create a new object for the x part of obj[x,y]:=z.
			else if (IS_INVOKE_SET)
			{
				Object *new_obj = new Object();
				if (new_obj)
				{
					if ( field = Insert(key_type, key, insert_pos) )
					{	// Don't do field->Assign() since it would do AddRef() and we would need to counter with Release().
						field->symbol = SYM_OBJECT;
						field->object = obj = new_obj;
					}
					else
					{	// Create() succeeded but Insert() failed, so free the newly created obj.
						new_obj->Release();
					}
				}
			}
		}
		if (obj) // Object was successfully found or created.
		{
			// obj now contains a pointer to the object contained by this field, possibly newly created above.
			ExprTokenType obj_token;
			obj_token.symbol = SYM_OBJECT;
			obj_token.object = obj;
			// References in obj_token and obj weren't counted (AddRef wasn't called), so Release() does not
			// need to be called before returning, and accessing obj after calling Invoke() would not be safe
			// since it could Release() the object (by overwriting our field via script) as a side-effect.
			// Recursively invoke obj, passing remaining parameters; remove IF_META to correctly treat obj as target:
			return obj->Invoke(aResultToken, obj_token, aFlags & ~IF_META, aParam + 1, aParamCount - 1);
			// Above may return INVOKE_NOT_HANDLED in cases such as obj[a,b] where obj[a] exists but obj[a][b] does not.
		}
	} // MULTIPARAM[x,y]

	// SET
	else if (IS_INVOKE_SET)
	{
		if (!IS_INVOKE_META && param_count_excluding_rvalue)
		{
			ExprTokenType &value_param = *aParam[1];
			// L34: Assigning an empty string no longer removes the field.
			if ( (field || (field = Insert(key_type, key, insert_pos))) && field->Assign(value_param) )
			{
				if (field->symbol == SYM_OPERAND)
				{
					// Use value_param since our copy may be freed prematurely in some (possibly rare) cases:
					aResultToken.symbol = SYM_STRING;
					aResultToken.marker = TokenToString(value_param);
					// Below: no longer used as other areas expect string results to always be SYM_STRING.
					// If it is ever used in future, we MUST copy marker and buf separately for x64 support.
					// However, it seems appropriate *not* to return a SYM_OPERAND with cached binary integer
					// since above has stored only the string part of it.
					//aResultToken.symbol		 = value_param.symbol;
					//aResultToken.value_int64 = value_param.value_int64; // Copy marker and buf (via union) in case it is SYM_OPERAND with a cached integer.
				}
				else
					field->Get(aResultToken); // L34: Corrected this to be aResultToken instead of value_param (broken by L33).
			}
			return OK;
		}
	}

	// GET
	else if (field)
	{
		if (field->symbol == SYM_OPERAND)
		{
			// Use SYM_STRING and not SYM_OPERAND, since SYM_OPERAND's use of aResultToken.buf
			// would conflict with the use of mem_to_free/buf to return a memory allocation.
			aResultToken.symbol = SYM_STRING;
			// L33: Make a persistent copy; our copy might be freed indirectly by releasing this object.
			//		Prior to L33, callers took care of this UNLESS this was the last op in an expression.
			if (!TokenSetResult(aResultToken, field->marker))
				aResultToken.marker = _T("");
		}
		else
			field->Get(aResultToken);

		return OK;
	}

	return INVOKE_NOT_HANDLED;
}

//
// Internal: Object::CallField - Used by Object::Invoke to call a function/method stored in this object.
//

ResultType Object::CallField(FieldType *aField, ExprTokenType &aResultToken, ExprTokenType &aThisToken, int aFlags, ExprTokenType *aParam[], int aParamCount)
// aParam[0] contains the identifier of this field or an empty space (for __Get etc.).
{
	if (aField->symbol == SYM_OBJECT)
	{
		ExprTokenType field_token;
		field_token.symbol = SYM_OBJECT;
		field_token.object = aField->object;
		ExprTokenType *tmp = aParam[0];
		// Something must be inserted into the parameter list to remove any ambiguity between an intentionally
		// and directly called function of 'that' object and one of our parameters matching an existing name.
		// Rather than inserting something like an empty string, it seems more useful to insert 'this' object,
		// allowing 'that' to change (via __Call) the behaviour of a "function-call" which operates on 'this'.
		// Consequently, if 'that[this]' contains a value, it is invoked; seems obscure but rare, and could
		// also be of use (for instance, as a means to remove the 'this' parameter or replace it with 'that').
		aParam[0] = &aThisToken;
		ResultType r = aField->object->Invoke(aResultToken, field_token, IT_CALL | IF_FUNCOBJ, aParam, aParamCount);
		aParam[0] = tmp;
		return r;
	}
	else if (aField->symbol == SYM_OPERAND)
	{
		Func *func = g_script.FindFunc(aField->marker);
		if (func)
		{
			// At this point, aIdCount == 1 and aParamCount includes only the explicit parameters for the call.
			if (IS_INVOKE_META)
			{
				ExprTokenType *tmp = aParam[0];
				// Called indirectly by means of the meta-object mechanism (mBase); treat it as a "method call".
				// For this type of call, "this" object is included as the first parameter.  To do this, aParam[0] is
				// temporarily overwritten with a pointer to aThisToken.  Note that aThisToken contains the original
				// object specified in script, not the real "this" which is actually a meta-object/base of that object.
				aParam[0] = &aThisToken;
				ResultType r = CallFunc(*func, aResultToken, aParam, aParamCount);
				aParam[0] = tmp;
				return r;
			}
			else
				// This object directly contains a function name.  Assume this object is intended
				// as a simple array of functions; do not pass aThisToken as is done above.
				// aParam + 1 vs aParam because aParam[0] is the key which was used to find this field, not a parameter of the call.
				return CallFunc(*func, aResultToken, aParam + 1, aParamCount - 1);
		}
	}
	return INVOKE_NOT_HANDLED;
}
	

//
// Object:: Built-in Methods
//

bool Object::InsertAt(INT_PTR aOffset, INT_PTR aKey, ExprTokenType *aValue[], int aValueCount)
{
	IndexType value_count = (IndexType)aValueCount;
	IndexType need_capacity = mFieldCount + value_count;
	if (need_capacity > mFieldCountMax && !SetInternalCapacity(need_capacity))
		// Fail.
		return false;
	FieldType *field = mFields + aOffset;
	if (aOffset < mFieldCount)
		memmove(field + value_count, field, (mFieldCount - aOffset) * sizeof(FieldType));
	mFieldCount += value_count;
	mKeyOffsetObject += value_count; // ints before objects
	mKeyOffsetString += value_count; // and strings
	FieldType *field_end;
	// Set keys and copy value params into the fields.
	for (field_end = field + value_count; field < field_end; ++field)
	{
		field->key.i = aKey++;
		field->symbol = SYM_INTEGER; // Must be init'd for Assign().
		field->Assign(**(aValue++));
	}
	// Adjust keys of fields which have been moved.
	for (field_end = mFields + mKeyOffsetObject; field < field_end; ++field)
	{
		field->key.i += value_count; // NOT =++key.i since keys might not be contiguous.
	}
	return true;
}

ResultType Object::_Insert(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
// _Insert( key, value )
{
	if (!aParamCount)
		return OK; // Error.

	SymbolType key_type;
	KeyType key;
	IndexType insert_pos, pos;
	FieldType *field = NULL;

	if (aParamCount == 1)
	{
		// Insert at the end when no key is supplied, since that is typically most useful
		// and is also most efficient (because no int-keyed fields are moved or adjusted).
		insert_pos = mKeyOffsetObject; // int keys end here.
		key.i = insert_pos ? mFields[insert_pos - 1].key.i + 1 : 1;
		key_type = SYM_INTEGER;
	}
	else
	{
		field = FindField(**aParam, aResultToken.buf, /*out*/ key_type, /*out*/ key, /*out*/ insert_pos);
		if (key_type == SYM_INTEGER)
		{
			if (field)
			{	// Don't overwrite this key's value; instead insert a new field.
				insert_pos = field - mFields; // insert_pos wasn't set in this case.
				field = NULL;
			}

			if (aParamCount > 2) // Multiple value params.  Could also handle aParamCount == 2, but the simpler method is faster.
			{
				if (InsertAt(insert_pos, key.i, aParam + 1, aParamCount - 1))
				{
					aResultToken.symbol = SYM_INTEGER;
					aResultToken.value_int64 = 1;
				}
				return OK;
			}
		}
		else
			if (aParamCount > 2)
				// Error: multiple values but not an integer key.
				return OK;
		++aParam; // See below.
	}
	// If we were passed only one parameter, aParam points to it.  Otherwise it
	// was interpreted as the key and aParam now points to the next parameter.
	
	if ( field || (field = Insert(key_type, key, insert_pos)) )
	{
		// Assign this field its new value:
		field->Assign(**aParam);
		// Increment any numeric keys following this one.  At this point, insert_pos always indicates the position of a field just inserted.
		if (key_type == SYM_INTEGER)
			for (pos = insert_pos + 1; pos < mKeyOffsetObject; ++pos)
				++mFields[pos].key.i;
		// Return indication of success.  Probably isn't useful to return the caller-specified index; zero can be a valid index (and may be mistaken for "fail").
		aResultToken.symbol = SYM_INTEGER;
		aResultToken.value_int64 = 1;
	}
	// else insert failed; leave aResultToken at default, empty string.  Numeric indices are *not* adjusted in this case.
	return OK;
}

ResultType Object::_Remove(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
// _Remove( [ min_key, max_key ] )
{
	if (aParamCount > 2)
		return OK;

	FieldType *min_field, *max_field;
	IndexType min_pos, max_pos, pos;
	SymbolType min_key_type, max_key_type;
	KeyType min_key, max_key;

	// Find the position of "min".
	if (!aParamCount)
	{
		if (mKeyOffsetObject) // i.e. at least one int field; use _MaxIndex()
		{
			min_field = &mFields[min_pos = mKeyOffsetObject - 1];
			min_key = min_field->key;
			min_key_type = SYM_INTEGER;
		}
		else // No appropriate field to remove, just return "".
			return OK;
	}
	else
		if (min_field = FindField(*aParam[0], aResultToken.buf, min_key_type, min_key, min_pos))
			min_pos = min_field - mFields; // else min_pos was already set by FindField.
	
	if (aParamCount > 1)
	{
		// Find the position following "max".
		if (max_field = FindField(*aParam[1], aResultToken.buf, max_key_type, max_key, max_pos))
			max_pos = max_field - mFields + 1;
		// Since the order of key-types in mFields is of no logical consequence, require that both keys be the same type.
		// Do not allow removing a range of object keys since there is probably no meaning to their order.
		if (max_key_type != min_key_type || max_key_type == SYM_OBJECT || max_pos < min_pos
			// min and max are different types, are objects, or max < min.
			|| (max_pos == min_pos && (max_key_type == SYM_INTEGER ? max_key.i < min_key.i : _tcsicmp(max_key.s, min_key.s) < 0)))
			// max < min, but no keys exist in that range so (max_pos < min_pos) check above didn't catch it.
		{
			if (min_key_type == SYM_INTEGER && max_key_type == SYM_STRING && !*max_key.s)
				// Allow Remove(i,"") to mean "remove [i] but don't adjust keys".
				aParamCount = 1;
			else
				return OK;
		}
		//else if (max_pos == min_pos): specified range is valid, but doesn't match any keys.
		//	Continue on, adjust integer keys as necessary and return 0 instead of "".
	}
	else
		max_key_type = min_key_type;

	if (aParamCount < 2)
	{
		if (!min_field) // Nothing to remove.
		{
			// L34: Must not continue since min_pos points at the wrong key or an invalid location.
			// As of L50, our return value when (aParamCount < 2) is the value which was removed, not
			// the number of removed items. Since this[min_field] would return "", an empty string is
			// the only sensible value we can return here. An earlier check already ensured Remove()
			// with no params returned "" if there was nothing to remove.
			aResultToken.symbol = SYM_STRING;	
			aResultToken.marker = _T(""); // L61: Changed from 0 to "". See comment above.
			return OK;
		}
		// Since only one field (at maximum) can be removed in this mode, it
		// seems more useful to return the field being removed than a count.
		switch (aResultToken.symbol = min_field->symbol)
		{
		case SYM_OPERAND:
			if (min_field->size)
			{
				// Detach the memory allocated for this field's string and pass it back to caller.
				aResultToken.mem_to_free = aResultToken.marker = min_field->marker;
				aResultToken.marker_length = _tcslen(aResultToken.marker); // NOT min_field->size, which is the allocation size.
				min_field->size = 0; // Prevent Free() from freeing min_field->marker.
			}
			//else aResultToken already contains an empty string.
			break;
		case SYM_OBJECT:
			aResultToken.object = min_field->object;
			min_field->symbol = SYM_INTEGER; // Prevent Free() from calling object->Release(), instead of calling AddRef().
			break;
		default:
			aResultToken.value_int64 = min_field->n_int64; // Effectively also value_double = n_double.
		}
		// Set these up as if caller did _Remove(min_key, min_key):
		max_pos = min_pos + 1;
		max_key.i = min_key.i; // Union copy. Used only if min_key_type == SYM_INTEGER; has no effect in other cases.
	}

	for (pos = min_pos; pos < max_pos; ++pos)
		// Free each field in the range being removed.
		mFields[pos].Free();

	IndexType remaining_fields = mFieldCount - max_pos;
	if (remaining_fields)
		// Move remaining fields left to fill the gap left by the removed range.
		memmove(mFields + min_pos, mFields + max_pos, remaining_fields * sizeof(FieldType));
	// Adjust count by the actual number of fields in the removed range.
	IndexType actual_count_removed = max_pos - min_pos;
	mFieldCount -= actual_count_removed;
	// Adjust key offsets and numeric keys as necessary.
	if (min_key_type != SYM_STRING) // i.e. SYM_OBJECT or SYM_INTEGER
	{
		mKeyOffsetString -= actual_count_removed;
		if (min_key_type == SYM_INTEGER)
		{
			mKeyOffsetObject -= actual_count_removed;
			if (max_key_type == SYM_INTEGER)
			{
				// Regardless of whether any fields were removed, min_pos contains the position of the field which
				// immediately followed the specified range.  Decrement each numeric key from this position onward.
				IntKeyType logical_count_removed = max_key.i - min_key.i + 1;
				if (logical_count_removed > 0)
					for (pos = min_pos; pos < mKeyOffsetObject; ++pos)
						mFields[pos].key.i -= logical_count_removed;
			}
			//else max_key must be "", a special flag caller may pass to indicate keys shouldn't be adjusted.
		}
	}
	if (aParamCount > 1)
	{
		// Return actual number of fields removed:
		aResultToken.symbol = SYM_INTEGER;
		aResultToken.value_int64 = actual_count_removed;
	}
	//else result was set above.
	return OK;
}

ResultType Object::_MinIndex(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	if (aParamCount)
		return OK;

	if (mKeyOffsetObject) // i.e. there are fields with integer keys
	{
		aResultToken.symbol = SYM_INTEGER;
		aResultToken.value_int64 = (__int64)mFields[0].key.i;
	}
	// else no integer keys; leave aResultToken at default, empty string.
	return OK;
}

ResultType Object::_MaxIndex(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	if (aParamCount)
		return OK;

	if (mKeyOffsetObject) // i.e. there are fields with integer keys
	{
		aResultToken.symbol = SYM_INTEGER;
		aResultToken.value_int64 = (__int64)mFields[mKeyOffsetObject - 1].key.i;
	}
	// else no integer keys; leave aResultToken at default, empty string.
	return OK;
}

ResultType Object::_GetCapacity(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	if (aParamCount == 1)
	{
		SymbolType key_type;
		KeyType key;
		IndexType insert_pos;
		FieldType *field;

		if ( (field = FindField(*aParam[0], aResultToken.buf, /*out*/ key_type, /*out*/ key, /*out*/ insert_pos))
			&& field->symbol == SYM_OPERAND )
		{
			aResultToken.symbol = SYM_INTEGER;
			aResultToken.value_int64 = field->size ? _TSIZE(field->size - 1) : 0; // -1 to exclude null-terminator.
		}
		// else wrong type of field; leave aResultToken at default, empty string.
	}
	else if (aParamCount == 0)
	{
		aResultToken.symbol = SYM_INTEGER;
		aResultToken.value_int64 = mFieldCountMax;
	}
	return OK;
}

ResultType Object::_SetCapacity(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
// _SetCapacity( [field_name,] new_capacity )
{
	if ((aParamCount != 1 && aParamCount != 2) || !TokenIsPureNumeric(*aParam[aParamCount - 1]))
		// Invalid or missing param(s); return default empty string.
		return OK;
	__int64 desired_capacity = TokenToInt64(*aParam[aParamCount - 1], TRUE);
	if (aParamCount == 2) // Field name was specified.
	{
		if (desired_capacity < 0) // Check before sign is dropped.
			// Bad param.
			return OK;
		size_t desired_size = (size_t)desired_capacity;

		SymbolType key_type;
		KeyType key;
		IndexType insert_pos;
		FieldType *field;
		LPTSTR buf;

		if ( (field = FindField(*aParam[0], aResultToken.buf, /*out*/ key_type, /*out*/ key, /*out*/ insert_pos))
			|| (field = Insert(key_type, key, insert_pos)) )
		{	
			// Field was successfully found or inserted.
			if (field->symbol != SYM_OPERAND)
				// Wrong type of field.
				return OK;
			if (!desired_size)
			{	// Caller specified zero - empty the field but do not remove it.
				field->Assign(NULL);
				aResultToken.symbol = SYM_INTEGER;
				aResultToken.value_int64 = 0;
				return OK;
			}
#ifdef UNICODE
			// Convert size in bytes to size in chars.
			desired_size = (desired_size >> 1) + (desired_size & 1);
#endif
			// Like VarSetCapacity, always reserve one char for the null-terminator.
			++desired_size;
			// Unlike VarSetCapacity, allow fields to shrink; preserve existing data up to min(new size, old size).
			// size is checked because if it is 0, marker is Var::sEmptyString which we can't pass to realloc.
			if (buf = trealloc(field->size ? field->marker : NULL, desired_size))
			{
				buf[desired_size - 1] = '\0'; // Terminate at the new end of data.
				field->marker = buf;
				field->size = desired_size;
				// Return new size, minus one char reserved for null-terminator.
				aResultToken.symbol = SYM_INTEGER;
				aResultToken.value_int64 = _TSIZE(desired_size - 1);
			}
			//else out of memory.
		}
		return OK;
	}
	IndexType desired_count = (IndexType)desired_capacity;
	// else aParamCount == 1: set the capacity of this object.
	if (desired_count < mFieldCount)
	{	// It doesn't seem intuitive to allow _SetCapacity to truncate the fields array.
		desired_count = mFieldCount;
	}
	if (!desired_count)
	{	// Caller wants to shrink object to current contents but there aren't any, so free mFields.
		if (mFields)
		{
			free(mFields);
			mFields = NULL;
			mFieldCountMax = 0;
		}
		//else mFieldCountMax should already be 0.
		// Since mFieldCountMax and desired_size are both 0, below will return 0 and won't call SetInternalCapacity.
	}
	if (desired_count == mFieldCountMax || SetInternalCapacity(desired_count))
	{
		aResultToken.symbol = SYM_INTEGER;
		aResultToken.value_int64 = mFieldCountMax;
	}
	return OK;
}

ResultType Object::_GetAddress(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
// _GetAddress( key )
{
	if (aParamCount == 1)
	{
		SymbolType key_type;
		KeyType key;
		IndexType insert_pos;
		FieldType *field;

		if ( (field = FindField(*aParam[0], aResultToken.buf, /*out*/ key_type, /*out*/ key, /*out*/ insert_pos))
			&& field->symbol == SYM_OPERAND && field->size )
		{
			aResultToken.symbol = SYM_INTEGER;
			aResultToken.value_int64 = (__int64)field->marker;
		}
		// else field has no memory allocated; leave aResultToken at default, empty string.
	}
	return OK;
}

ResultType Object::_NewEnum(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	if (aParamCount == 0)
	{
		IObject *newenum;
		if (newenum = new Enumerator(this))
		{
			aResultToken.symbol = SYM_OBJECT;
			aResultToken.object = newenum;
		}
	}
	return OK;
}

ResultType Object::_HasKey(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	if (aParamCount == 1)
	{
		SymbolType key_type;
		KeyType key;
		INT_PTR insert_pos;
		FieldType *field = FindField(**aParam, aResultToken.buf, key_type, key, insert_pos);
		aResultToken.symbol = SYM_INTEGER;
		aResultToken.value_int64 = (field != NULL);
	}
	return OK;
}

ResultType Object::_Clone(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	if (aParamCount == 0)
	{
		Object *clone = Clone();
		if (clone)
		{
			if (mBase)
				(clone->mBase = mBase)->AddRef();
			aResultToken.object = clone;
			aResultToken.symbol = SYM_OBJECT;
		}
	}
	return OK;
}


//
// Object::FieldType
//

bool Object::FieldType::Assign(LPTSTR str, size_t len, bool exact_size)
{
	if (!str || !*str && (len == -1 || !len)) // If empty string or null pointer, free our contents.  Passing len >= 1 allows copying \0, so don't check *str in that case.  Ordered for short-circuit performance (len is usually -1).
	{
		Free();
		symbol = SYM_OPERAND;
		marker = Var::sEmptyString;
		size = 0;
		return true;
	}
	
	if (len == -1)
		len = _tcslen(str);

	if (symbol != SYM_OPERAND || len >= size)
	{
		Free(); // Free object or previous buffer (which was too small).
		symbol = SYM_OPERAND;
		size_t new_size = len + 1;
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
		if ( !(marker = tmalloc(new_size)) )
		{
			marker = Var::sEmptyString;
			size = 0;
			return false; // See "Sanity check" above.
		}
		size = new_size;
	}
	// else we have a buffer with sufficient capacity already.

	tmemcpy(marker, str, len + 1); // +1 for null-terminator.
	return true; // Success.
}

bool Object::FieldType::Assign(ExprTokenType &aParam)
{
	ExprTokenType temp, *val; // Seems more maintainable to use a copy; avoid any possible side-effects.
	if (aParam.symbol == SYM_VAR)
	{
		// Primary reason for this check: If var has a cached binary integer, we want to use it and
		// not the stringified copy of it.  It seems unlikely that scripts will depend on the string
		// format of a literal number such as 0x123 or 00001, and even less likely for a number stored
		// in an object (even implicitly via a variadic function).  If the value is eventually passed
		// to a COM method call, it can be important that it is passed as VT_I4 and not VT_BSTR.
		aParam.var->TokenToContents(temp);
		val = &temp;
	}
	else
		val = &aParam;

	switch (val->symbol)
	{
	case SYM_OPERAND:
		if (val->buf)
		{
			// Store this integer literal as a pure integer.  See above for comments.
			Free();
			symbol = SYM_INTEGER;
			n_int64 = *(__int64 *)val->buf;
			break;
		}
		// FALL THROUGH to the next case.
	case SYM_STRING:
		return Assign(val->marker);
	case SYM_INTEGER:
	case SYM_FLOAT:
		Free(); // Free string or object, if applicable.
		symbol = val->symbol; // Either SYM_INTEGER or SYM_FLOAT.  Set symbol *after* calling Free().
		n_int64 = val->value_int64; // Also handles value_double via union.
		break;
	case SYM_OBJECT:
		Free(); // Free string or object, if applicable.
		symbol = SYM_OBJECT; // Set symbol *after* calling Free().
		object = val->object;
		if (aParam.symbol != SYM_VAR)
			object->AddRef();
		// Otherwise, take ownership of the ref in temp.
		break;
	default:
		ASSERT(FALSE);
	}
	return true;
}

void Object::FieldType::Get(ExprTokenType &result)
{
	result.symbol = symbol;
	result.value_int64 = n_int64; // Union copy.
	if (symbol == SYM_OBJECT)
		object->AddRef();
}

void Object::FieldType::Free()
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


//
// Enumerator
//

ResultType STDMETHODCALLTYPE EnumBase::Invoke(ExprTokenType &aResultToken, ExprTokenType &aThisToken, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	if (IS_INVOKE_SET)
		return INVOKE_NOT_HANDLED;

	if (IS_INVOKE_CALL)
	{
		if (aParamCount && !_tcsicmp(TokenToString(*aParam[0]), _T("Next")))
		{	// This is something like enum.Next(var); exclude "Next" so it is treated below as enum[var].
			++aParam; --aParamCount;
		}
		else
			return INVOKE_NOT_HANDLED;
	}
	Var *var0 = NULL, *var1 = NULL;
	if (aParamCount)
	{
		if (aParam[0]->symbol != SYM_VAR)
			return OK;
		if (aParamCount > 1)
		{
			if (aParam[1]->symbol != SYM_VAR)
				return OK;
			var1 = aParam[1]->var;
		}
		var0 = aParam[0]->var;
	}
	aResultToken.symbol = SYM_INTEGER;
	aResultToken.value_int64 = Next(var0, var1);
	return OK;
}

int Object::Enumerator::Next(Var *aKey, Var *aVal)
{
	if (++mOffset < mObject->mFieldCount)
	{
		FieldType &field = mObject->mFields[mOffset];
		if (aKey)
		{
			if (mOffset < mObject->mKeyOffsetObject) // mKeyOffsetInt < mKeyOffsetObject
				aKey->Assign(field.key.i);
			else if (mOffset < mObject->mKeyOffsetString) // mKeyOffsetObject < mKeyOffsetString
				aKey->Assign(field.key.p);
			else // mKeyOffsetString < mFieldCount
				aKey->Assign(field.key.s);
		}
		if (aVal)
		{
			switch (field.symbol)
			{
			case SYM_OPERAND:	aVal->AssignString(field.marker);	break;
			case SYM_INTEGER:	aVal->Assign(field.n_int64);			break;
			case SYM_FLOAT:		aVal->Assign(field.n_double);		break;
			case SYM_OBJECT:	aVal->Assign(field.object);			break;
			}
		}
		return true;
	}
	return false;
}

	

//
// Object:: Internal Methods
//

template<typename T>
Object::FieldType *Object::FindField(T val, INT_PTR left, INT_PTR right, INT_PTR &insert_pos)
// Template used below.  left and right must be set by caller to the appropriate bounds within mFields.
{
	INT_PTR mid, result;
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

Object::FieldType *Object::FindField(SymbolType key_type, KeyType key, IndexType &insert_pos)
// Searches for a field with the given key.  If found, a pointer to the field is returned.  Otherwise
// NULL is returned and insert_pos is set to the index a newly created field should be inserted at.
// key_type and key are output for creating a new field or removing an existing one correctly.
// left and right must indicate the appropriate section of mFields to search, based on key type.
{
	IndexType left, right;

	if (key_type == SYM_STRING)
	{
		left = mKeyOffsetString;
		right = mFieldCount - 1; // String keys are last in the mFields array.

		return FindField<LPTSTR>(key.s, left, right, insert_pos);
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
		return FindField<IntKeyType>(key.i, left, right, insert_pos);
	}
}

Object::FieldType *Object::FindField(ExprTokenType &key_token, LPTSTR aBuf, SymbolType &key_type, KeyType &key, IndexType &insert_pos)
// Searches for a field with the given key, where the key is a token passed from script.
{
	if (TokenIsPureNumeric(key_token) == PURE_INTEGER)
	{	// Treat all integer keys (even numeric strings) as pure integers for consistency and performance.
		key.i = (IntKeyType)TokenToInt64(key_token, TRUE);
		key_type = SYM_INTEGER;
	}
	else if (key.p = TokenToObject(key_token))
	{	// SYM_OBJECT or SYM_VAR containing object.
		key_type = SYM_OBJECT;
	}
	else
	{	// SYM_STRING, SYM_FLOAT, SYM_OPERAND or SYM_VAR (all confirmed not to be an integer at this point).
		key.s = TokenToString(key_token, aBuf); // L41: Pass aBuf to allow float->string conversion as documented (but not previously working).
		key_type = SYM_STRING;
	}
	return FindField(key_type, key, insert_pos);
}
	
bool Object::SetInternalCapacity(IndexType new_capacity)
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
	
Object::FieldType *Object::Insert(SymbolType key_type, KeyType key, IndexType at)
// Inserts a single field with the given key at the given offset.
// Caller must ensure 'at' is the correct offset for this key.
{
	if (mFieldCount == mFieldCountMax && !Expand()  // Attempt to expand if at capacity.
		|| key_type == SYM_STRING && !(key.s = _tcsdup(key.s)))  // Attempt to duplicate key-string.
	{	// Out of memory.
		return NULL;
	}
	// There is now definitely room in mFields for a new field.

	FieldType &field = mFields[at];
	if (at < mFieldCount)
		// Move existing fields to make room.
		memmove(&field + 1, &field, (mFieldCount - at) * sizeof(FieldType));
	++mFieldCount; // Only after memmove above.
	
	// Update key-type offsets based on where and what was inserted; also update this key's ref count:
	if (key_type != SYM_STRING)
	{
		// Must be either SYM_INTEGER or SYM_OBJECT, which both precede SYM_STRING.
		++mKeyOffsetString;

		if (key_type != SYM_OBJECT)
			// Must be SYM_INTEGER, which precedes SYM_OBJECT.
			++mKeyOffsetObject;
		else
			key.p->AddRef();
	}
	
	field.marker = _T(""); // Init for maintainability.
	field.size = 0; // Init to ensure safe behaviour in Assign().
	field.key = key; // Above has already copied string or called key.p->AddRef() as appropriate.
	field.symbol = SYM_OPERAND;

	return &field;
}


//
// Func: Script interface, accessible via "function reference".
//

ResultType STDMETHODCALLTYPE Func::Invoke(ExprTokenType &aResultToken, ExprTokenType &aThisToken, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	if (!aParamCount)
		return INVOKE_NOT_HANDLED;

	LPTSTR member = TokenToString(*aParam[0]);

	if (!IS_INVOKE_CALL)
	{
		if (IS_INVOKE_SET || aParamCount > 1)
			return INVOKE_NOT_HANDLED;

		if (!_tcsicmp(member, _T("Name")))
		{
			aResultToken.symbol = SYM_STRING;
			aResultToken.marker = mName;
		}
		else if (!_tcsicmp(member, _T("MinParams")))
		{
			aResultToken.symbol = SYM_INTEGER;
			aResultToken.value_int64 = mMinParams;
		}
		else if (!_tcsicmp(member, _T("MaxParams")))
		{
			aResultToken.symbol = SYM_INTEGER;
			aResultToken.value_int64 = mParamCount;
		}
		else if (!_tcsicmp(member, _T("IsBuiltIn")))
		{
			aResultToken.symbol = SYM_INTEGER;
			aResultToken.value_int64 = mIsBuiltIn;
		}
		else if (!_tcsicmp(member, _T("IsVariadic")))
		{
			aResultToken.symbol = SYM_INTEGER;
			aResultToken.value_int64 = mIsVariadic;
		}
		else
			return INVOKE_NOT_HANDLED;

		return OK;
	}
	
	if (  !(aFlags & IF_FUNCOBJ)  )
	{
		if (!_tcsicmp(member, _T("IsOptional")) && aParamCount <= 2)
		{
			if (aParamCount == 2)
			{
				int param = (int)TokenToInt64(*aParam[1]); // One-based.
				if (param > 0 && (param <= mParamCount || mIsVariadic))
				{
					aResultToken.symbol = SYM_INTEGER;
					aResultToken.value_int64 = param > mMinParams;
				}
			}
			else
			{
				aResultToken.symbol = SYM_INTEGER;
				aResultToken.value_int64 = mMinParams != mParamCount || mIsVariadic; // True if any params are optional.
			}
			return OK;
		}
		else if (!_tcsicmp(member, _T("IsByRef")) && aParamCount <= 2 && !mIsBuiltIn)
		{
			if (aParamCount == 2)
			{
				int param = (int)TokenToInt64(*aParam[1]); // One-based.
				if (param > 0 && (param <= mParamCount || mIsVariadic))
				{
					aResultToken.symbol = SYM_INTEGER;
					aResultToken.value_int64 = param <= mParamCount && mParam[param-1].is_byref;
				}
			}
			else
			{
				aResultToken.symbol = SYM_INTEGER;
				aResultToken.value_int64 = FALSE;
				for (int param = 0; param < mParamCount; ++param)
					if (mParam[param].is_byref)
					{
						aResultToken.value_int64 = TRUE;
						break;
					}
			}
			return OK;
		}
		if (!TokenIsEmptyString(*aParam[0]))
			return INVOKE_NOT_HANDLED; // Reserved.
		// Called explicitly by script, such as by "obj.funcref.()" or "x := obj.funcref, x.()"
		// rather than implicitly, like "obj.funcref()".
		++aParam;		// Discard the "method name" parameter.
		--aParamCount;	// 
	}
	return CallFunc(*this, aResultToken, aParam, aParamCount);
}
	

//
// MetaObject - Defines behaviour of object syntax when used on a non-object value.
//

MetaObject g_MetaObject;

LPTSTR Object::sMetaFuncName[] = { _T("__Get"), _T("__Set"), _T("__Call"), _T("__Delete"), _T("__New") };

ResultType STDMETHODCALLTYPE MetaObject::Invoke(ExprTokenType &aResultToken, ExprTokenType &aThisToken, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	// For something like base.Method() in a class-defined method:
	// It seems more useful and intuitive for this special behaviour to take precedence over
	// the default meta-functions, especially since "base" may become a reserved word in future.
	if (aThisToken.symbol == SYM_VAR && !_tcsicmp(aThisToken.var->mName, _T("base"))
		&& !aThisToken.var->HasContents() // Since scripts are able to assign to it, may as well let them use the assigned value.
		&& g->CurrentFunc)
	{
		LPCTSTR full_name = g->CurrentFunc->mName;
		LPCTSTR end_marker = _tcsrchr(full_name, '.');
		Object *this_class;
		Var *this_var;
		if (   end_marker // Appears to be a class definition.
			&& (this_class = g_script.FindClass(full_name, end_marker - full_name)) // Can fail if the class var was reassigned at run-time.
			&& (this_var = g->CurrentFunc->mParam[0].var)->IsObject()   ) // Valid 'this' param.
		{
			ExprTokenType this_token;
			this_token.symbol = SYM_VAR;
			this_token.var = this_var;
			if (IObject *this_class_base = this_class->Base())
			{
				return this_class_base->Invoke(aResultToken, this_token, (aFlags & ~IF_METAFUNC) | IF_METAOBJ, aParam, aParamCount);
			}
			return OK;
		}
	}

	// Allow script-defined meta-functions to override the default behaviour:
	ResultType result = Object::Invoke(aResultToken, aThisToken, aFlags, aParam, aParamCount);
	
	if (result != INVOKE_NOT_HANDLED || !aParamCount)
		return result;

	if (IS_INVOKE_CALL && TokenIsEmptyString(*aParam[0]))
	{
		// Support func_var.(params) as a means to call either a function or an object-function.
		// This can be done fairly easily in script, but not in a way that supports ByRef;
		// and as a default behaviour, it might be more appealing for func lib authors.
		LPCTSTR func_name = TokenToString(aThisToken, aResultToken.buf);
		Func *func = g_script.FindFunc(func_name, EXPR_TOKEN_LENGTH((&aThisToken), func_name));
		if (func)
			return CallFunc(*func, aResultToken, aParam + 1, aParamCount - 1);
	}

	return result;
}
