#include "stdafx.h" // pre-compiled headers
#include "defines.h"
#include "globaldata.h"
#include "script.h"

#include "script_object.h"
#include "script_func_impl.h"


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
		return g_script.ScriptError(ERR_TOO_FEW_PARAMS, aFunc.mName);

	// When this variable goes out of scope, Var::FreeAndRestoreFunctionVars() is called (if appropriate):
	FuncCallData func_call;
	ResultType result;

	// CALL THE FUNCTION.
	if (aFunc.Call(func_call, result, aResultToken, aParam, aParamCount)
		// Make return value persistent if applicable:
		&& aResultToken.symbol == SYM_STRING && !aFunc.mIsBuiltIn)
	{
		// Make a persistent copy of the string in case it is the contents of one of the function's local variables.
		if (!*aResultToken.marker)
			aResultToken.marker = _T("");
		else if (!TokenSetResult(aResultToken, aResultToken.marker))
			result = FAIL;
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
		if (aParamCount > 8)
			// Set initial capacity to avoid multiple expansions.
			// For simplicity, failure is handled by the loop below.
			obj->SetInternalCapacity(aParamCount >> 1);
		// Otherwise, there are 4 or less key-value pairs.  When the first
		// item is inserted, a default initial capacity of 4 will be set.

		TCHAR buf[MAX_NUMBER_SIZE];
		FieldType *field;
		SymbolType key_type;
		KeyType key;
		IndexType insert_pos;
		
		for (int i = 0; i + 1 < aParamCount; i += 2)
		{
			if (aParam[i]->symbol == SYM_MISSING || aParam[i+1]->symbol == SYM_MISSING)
				continue; // For simplicity.

			field = obj->FindField(*aParam[i], buf, key_type, key, insert_pos);
			if (!field // Probably NULL, but calling FindField() first avoids an extra call to TokenToString() below.
				&& key_type == SYM_STRING && !_tcsicmp(key.s, _T("base")))
			{
				// For consistency with assignments, the following is allowed to overwrite a previous
				// base object (although having "base" occur twice in the parameter list would be quite
				// useless) or set mBase to NULL if the value parameter is not an object.
				if (obj->mBase)
					obj->mBase->Release();
				if (obj->mBase = TokenToObject(*aParam[i + 1]))
					obj->mBase->AddRef();
				continue;
			}
			if (  !(field
				 || (field = obj->Insert(key_type, key, insert_pos)))
				|| !field->Assign(*aParam[i + 1])  )
			{	// Out of memory.
				obj->Release();
				return NULL;
			}
		}
	}
	return obj;
}


//
// Object::Clone - Used for variadic function-calls.
//

Object *Object::Clone(BOOL aExcludeIntegerKeys)
// Creates an object and copies to it the fields at and after the given offset.
{
	IndexType aStartOffset = aExcludeIntegerKeys ? mKeyOffsetObject : 0;

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
		case SYM_STRING:
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

void Object::ArrayToParams(ExprTokenType *token, ExprTokenType **param_list, int extra_params
	, ExprTokenType **&aParam, int &aParamCount)
// Expands this object's contents into the parameter list.  Due to the nature
// of the parameter list, only fields with integer keys are used (named params
// aren't supported).
// Return value is FAIL if a required parameter was omitted or malloc() failed.
{
	// Find the first and last field to be used.
	int start = (int)mKeyOffsetInt;
	int end = (int)mKeyOffsetObject; // For readability.
	while (start < end && mFields[start].key.i < 1)
		++start; // Skip any keys <= 0 (consistent with UDF-calling behaviour).
	
	int param_index;
	IndexType field_index;

	// For each extra param...
	for (field_index = start, param_index = 0; field_index < end; ++field_index, ++param_index)
	{
		for ( ; param_index + 1 < (int)mFields[field_index].key.i; ++param_index)
		{
			token[param_index].symbol = SYM_MISSING;
			token[param_index].marker = _T("");
		}
		mFields[field_index].ToToken(token[param_index]);
	}
	
	ExprTokenType **param_ptr = param_list;

	// Init the array of param token pointers.
	for (param_index = 0; param_index < aParamCount; ++param_index)
		*param_ptr++ = aParam[param_index]; // Caller-supplied param token.
	for (param_index = 0; param_index < extra_params; ++param_index)
		*param_ptr++ = &token[param_index]; // New param.

	aParam = param_list; // Update caller's pointer.
	aParamCount += extra_params; // Update caller's count.
}


//
// Object::ArrayToStrings - Used by BIF_StrSplit.
//

ResultType Object::ArrayToStrings(LPTSTR *aStrings, int &aStringCount, int aStringsMax)
{
	int i, j;
	for (i = 0, j = 0; i < aStringsMax && j < mKeyOffsetObject; ++j)
		if (SYM_STRING == mFields[j].symbol)
			aStrings[i++] = mFields[j].marker;
		else
			return FAIL;
	aStringCount = i;
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
		
		// Since defining base[key] prevents base.base.__Get and __Call from being invoked, it seems best
		// to have it also block __Set. The code below is disabled to achieve this, with a slight cost to
		// performance when assigning to a new key in any object which has a base object. (The cost may
		// depend on how many key-value pairs each base object has.) Note that this doesn't affect meta-
		// functions defined in *this* base object, since they were already invoked if present.
		//if (IS_INVOKE_META)
		//{
		//	if (param_count_excluding_rvalue == 1)
		//		// Prevent below from unnecessarily searching for a field, since it won't actually be assigned to.
		//		// Relies on mBase->Invoke recursion using aParamCount and not param_count_excluding_rvalue.
		//		param_count_excluding_rvalue = 0;
		//	//else: Allow SET to operate on a field of an object stored in the target's base.
		//	//		For instance, x[y,z]:=w may operate on x[y][z], x.base[y][z], x[y].base[z], etc.
		//}
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
				return CallBuiltin(key.s, aResultToken, aParam + 1, aParamCount - 1); // +/- 1 to exclude the method identifier.
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
			else if (!IS_INVOKE_META)
				return g_script.ScriptError(_T("Array is not multi-dimensional."));
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
						new_obj = NULL;
					}
				}
				if (!new_obj)
					return g_script.ScriptError(ERR_OUTOFMEM);
			}
			else if (IS_INVOKE_GET)
			{
				// Treat x[y,z] like x[y] when x[y] is not set: just return "", don't throw an exception.
				// On the other hand, if x[y] is set to something which is not an object, the "if (field)"
				// section above raises an error.
				return OK;
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
				if (field->symbol == SYM_STRING)
				{
					// Use value_param since our copy may be freed prematurely in some (possibly rare) cases:
					aResultToken.symbol = SYM_STRING;
					aResultToken.marker = TokenToString(value_param);
				}
				else
					field->Get(aResultToken); // L34: Corrected this to be aResultToken instead of value_param (broken by L33).
			}
			else
				return g_script.ScriptError(ERR_OUTOFMEM);
			return OK;
		}
	}

	// GET
	else
	{
		if (field)
		{
			if (field->symbol == SYM_STRING)
			{
				aResultToken.symbol = SYM_STRING;
				// L33: Make a persistent copy; our copy might be freed indirectly by releasing this object.
				//		Prior to L33, callers took care of this UNLESS this was the last op in an expression.
				return TokenSetResult(aResultToken, field->marker);
			}
			field->Get(aResultToken);
			return OK;
		}
		// If 'this' is the target object (not its base), produce OK so that something like if(!foo.bar) is
		// considered valid even when foo.bar has not been set.
		if (!IS_INVOKE_META && aParamCount)
			return OK;
	}

	// Fell through from one of the sections above: invocation was not handled.
	return INVOKE_NOT_HANDLED;
}


ResultType Object::CallBuiltin(LPCTSTR aName, ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	switch (toupper(*aName))
	{
	case 'I':
		if (!_tcsicmp(aName, _T("InsertAt")))
			return _InsertAt(aResultToken, aParam, aParamCount);
		break;
	case 'P':
		if (!_tcsicmp(aName, _T("Push")))
			return _Push(aResultToken, aParam, aParamCount);
		if (!_tcsicmp(aName, _T("Pop")))
			return _Pop(aResultToken, aParam, aParamCount);
		break;
	case 'R':
		if (!_tcsnicmp(aName, _T("Remove"), 6))
		{
			aName += 6;
			if (!*aName)
				return _Remove(aResultToken, aParam, aParamCount);
			if (!_tcsicmp(aName, _T("At")))
				return _RemoveAt(aResultToken, aParam, aParamCount);
		}
		break;
	case 'H':
		if (!_tcsicmp(aName, _T("HasKey")))
			return _HasKey(aResultToken, aParam, aParamCount);
		break;
	case 'M':
		if (!_tcsicmp(aName, _T("MaxIndex")))
			return _MaxIndex(aResultToken, aParam, aParamCount);
		if (!_tcsicmp(aName, _T("MinIndex")))
			return _MinIndex(aResultToken, aParam, aParamCount);
		break;
	case '_':
		if (!_tcsicmp(aName, _T("_NewEnum")))
			return _NewEnum(aResultToken, aParam, aParamCount);
		break;
	case 'G':
		if (!_tcsicmp(aName, _T("GetAddress")))
			return _GetAddress(aResultToken, aParam, aParamCount);
		if (!_tcsicmp(aName, _T("GetCapacity")))
			return _GetCapacity(aResultToken, aParam, aParamCount);
		break;
	case 'S':
		if (!_tcsicmp(aName, _T("SetCapacity")))
			return _SetCapacity(aResultToken, aParam, aParamCount);
		break;
	case 'C':
		if (!_tcsicmp(aName, _T("Clone")))
			return _Clone(aResultToken, aParam, aParamCount);
		break;
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
	else if (aField->symbol == SYM_STRING)
	{
		Func *func = g_script.FindFunc(aField->marker);
		if (func)
		{
			ExprTokenType *tmp = aParam[0];
			// v2: Always pass "this" as the first parameter.  The old behaviour of passing it only when called
			// indirectly via mBase was confusing to many users, and isn't needed now that the script can do
			// %this.func%() instead of this.func() if they don't want to pass "this".
			// For this type of call, "this" object is included as the first parameter.  To do this, aParam[0] is
			// temporarily overwritten with a pointer to aThisToken.  Note that aThisToken contains the original
			// object specified in script, not the C++ "this" which is actually a meta-object/base of that object.
			aParam[0] = &aThisToken;
			ResultType r = CallFunc(*func, aResultToken, aParam, aParamCount);
			aParam[0] = tmp;
			return r;
		}
	}
	// The field's value is neither a function reference nor the name of a known function.
	ExprTokenType tok;
	aField->ToToken(tok);
	return g_script.ScriptError(ERR_NONEXISTENT_FUNCTION, TokenToString(tok, aResultToken.buf));
}


//
// Helper function for WinMain()
//

Object *Object::CreateFromArgV(LPTSTR *aArgV, int aArgC)
{
	Object *args;
	if (  !(args = Create(NULL, 0))  )
		return NULL;
	if (aArgC < 1)
		return args;
	ExprTokenType *token = (ExprTokenType *)_alloca(aArgC * sizeof(ExprTokenType));
	ExprTokenType **param = (ExprTokenType **)_alloca(aArgC * sizeof(ExprTokenType*));
	for (int j = 0; j < aArgC; ++j)
	{
		token[j].symbol = SYM_STRING;
		token[j].marker = aArgV[j];
		param[j] = &token[j];
	}
	if (!args->InsertAt(0, 1, param, aArgC))
	{
		args->Release();
		return NULL;
	}
	return args;
}



//
// Helper function for StringSplit()
//

bool Object::Append(LPTSTR aValue, size_t aValueLength)
{
	if (mFieldCount == mFieldCountMax && !Expand()) // Attempt to expand if at capacity.
		return false;

	if (aValueLength == -1)
		aValueLength = _tcslen(aValue);
	
	FieldType &field = mFields[mKeyOffsetObject];
	if (mKeyOffsetObject < mFieldCount)
		// For maintainability. This might never be done, because our caller
		// doesn't use string/object keys. Move existing fields to make room:
		memmove(&field + 1, &field, (mFieldCount - mKeyOffsetObject) * sizeof(FieldType));
	++mFieldCount; // Only after memmove above.
	++mKeyOffsetObject;
	++mKeyOffsetString;
	
	// The following relies on the fact that callers of this function ONLY use
	// this function, so the last integer key == the number of integer keys.
	field.key.i = mKeyOffsetObject;

	field.symbol = SYM_STRING;
	if (aValueLength) // i.e. a non-empty string was supplied.
	{
		++aValueLength; // Convert length to size.
		if (field.marker = tmalloc(aValueLength))
		{
			tmemcpy(field.marker, aValue, aValueLength);
			field.marker[aValueLength-1] = '\0';
			field.size = aValueLength;
			return true;
		}
		// Otherwise, mem alloc failed; assign an empty string.
	}
	field.marker = Var::sEmptyString;
	field.size = 0;
	return (aValueLength == 0); // i.e. true if caller supplied an empty string.
}


//
// Helper function used with class definitions.
//

void Object::EndClassDefinition()
{
	// Instance variables were previously created as keys in the class object to prevent duplicate or
	// conflicting declarations.  Since these variables will be added at run-time to the derived objects,
	// we don't want them in the class object.  So delete any key-value pairs with the special marker
	// value (currently any integer, since static initializers haven't been evaluated yet).
	for (IndexType i = mFieldCount - 1; i >= 0; --i)
		if (mFields[i].symbol == SYM_INTEGER)
			if (i < --mFieldCount)
				memmove(mFields + i, mFields + i + 1, (mFieldCount - i) * sizeof(FieldType));
}


//
// Helper function for 'is' operator: is aBase a direct or indirect base object of this?
//

bool Object::IsDerivedFrom(IObject *aBase)
{
	IObject *ibase;
	Object *base;
	for (ibase = mBase; ; ibase = base->mBase)
	{
		if (ibase == aBase)
			return true;
		if (  !(base = dynamic_cast<Object *>(ibase))  )  // ibase may be NULL.
			return false;
	}
}
	

//
// Object:: Built-in Methods
//

bool Object::InsertAt(INT_PTR aOffset, INT_PTR aKey, ExprTokenType *aValue[], int aValueCount)
{
	IndexType actual_count = (IndexType)aValueCount;
	for (int i = 0; i < aValueCount; ++i)
		if (aValue[i]->symbol == SYM_MISSING)
			actual_count--;
	IndexType need_capacity = mFieldCount + actual_count;
	if (need_capacity > mFieldCountMax && !SetInternalCapacity(need_capacity))
		// Fail.
		return false;
	FieldType *field = mFields + aOffset;
	if (aOffset < mFieldCount)
		memmove(field + actual_count, field, (mFieldCount - aOffset) * sizeof(FieldType));
	mFieldCount += actual_count;
	mKeyOffsetObject += actual_count; // ints before objects
	mKeyOffsetString += actual_count; // and strings
	FieldType *field_end;
	// Set keys and copy value params into the fields.
	for (int i = 0; i < aValueCount; ++i, ++aKey)
	{
		ExprTokenType &value = *(aValue[i]);
		if (value.symbol != SYM_MISSING)
		{
			field->key.i = aKey;
			field->symbol = SYM_INTEGER; // Must be init'd for Assign().
			field->Assign(value);
			field++;
		}
	}
	// Adjust keys of fields which have been moved.
	for (field_end = mFields + mKeyOffsetObject; field < field_end; ++field)
	{
		field->key.i += aValueCount; // NOT =++key.i since keys might not be contiguous.
	}
	return true;
}

ResultType Object::_InsertAt(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
// InsertAt(index, value1, ...)
{
	if (aParamCount < 2)
		return g_script.ScriptError(ERR_TOO_FEW_PARAMS);

	SymbolType key_type;
	KeyType key;
	IndexType insert_pos;
	FieldType *field = FindField(**aParam, aResultToken.buf, /*out*/ key_type, /*out*/ key, /*out*/ insert_pos);
	if (key_type != SYM_INTEGER)
		return g_script.ScriptError(ERR_PARAM1_INVALID, key_type == SYM_STRING ? key.s : _T(""));
		
	if (field)
	{
		insert_pos = field - mFields; // insert_pos wasn't set in this case.
		field = NULL; // Insert, don't overwrite.
	}

	if (!InsertAt(insert_pos, key.i, aParam + 1, aParamCount - 1))
		return g_script.ScriptError(ERR_OUTOFMEM);
	
	// Leave aResultToken at its default empty value.
	return OK;
}

ResultType Object::_Push(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
// Push(value1, ...)
{
	IndexType insert_pos = mKeyOffsetObject; // int keys end here.;
	IntKeyType start_index = (insert_pos ? mFields[insert_pos - 1].key.i + 1 : 1);
	if (!InsertAt(insert_pos, start_index, aParam, aParamCount))
		return g_script.ScriptError(ERR_OUTOFMEM);

	// Return the new "length" of the array.
	aResultToken.symbol = SYM_INTEGER;
	aResultToken.value_int64 = start_index + aParamCount - 1;
	return OK;
}

ResultType Object::_Remove_impl(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount, RemoveMode aMode)
// Remove(first_key [, last_key := first_key])
// RemoveAt(index [, virtual_count := 1])
// Pop()
{
	FieldType *min_field;
	IndexType min_pos, max_pos, pos;
	SymbolType min_key_type;
	KeyType min_key, max_key;
	IntKeyType logical_count_removed = 1;

	// Find the position of "min".
	if (aMode == RM_Pop)
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
	{
		if (!aParamCount)
			return g_script.ScriptError(ERR_TOO_FEW_PARAMS);

		if (min_field = FindField(*aParam[0], aResultToken.buf, min_key_type, min_key, min_pos))
			min_pos = min_field - mFields; // else min_pos was already set by FindField.
		
		if (min_key_type != SYM_INTEGER && aMode != RM_RemoveKey)
			return g_script.ScriptError(ERR_PARAM1_INVALID);
	}
	
	if (aParamCount > 1) // Removing a range of keys.
	{
		SymbolType max_key_type;
		FieldType *max_field;
		if (aMode == RM_RemoveAt)
		{
			logical_count_removed = (IntKeyType)TokenToInt64(*aParam[1]);
			// Find the next position >= [aParam[1] + Count].
			max_key_type = SYM_INTEGER;
			max_key.i = min_key.i + logical_count_removed;
			if (max_field = FindField(max_key_type, max_key, max_pos))
				max_pos = max_field - mFields;
		}
		else
		{
			// Find the next position > [aParam[1]].
			if (max_field = FindField(*aParam[1], aResultToken.buf, max_key_type, max_key, max_pos))
				max_pos = max_field - mFields + 1;
		}
		// Since the order of key-types in mFields is of no logical consequence, require that both keys be the same type.
		// Do not allow removing a range of object keys since there is probably no meaning to their order.
		if (max_key_type != min_key_type || max_key_type == SYM_OBJECT || max_pos < min_pos
			// min and max are different types, are objects, or max < min.
			|| (max_pos == min_pos && (max_key_type == SYM_INTEGER ? max_key.i < min_key.i : _tcsicmp(max_key.s, min_key.s) < 0)))
			// max < min, but no keys exist in that range so (max_pos < min_pos) check above didn't catch it.
		{
			return g_script.ScriptError(ERR_PARAM2_INVALID);
		}
		//else if (max_pos == min_pos): specified range is valid, but doesn't match any keys.
		//	Continue on, adjust integer keys as necessary and return 0.
	}
	else // Removing a single item.
	{
		if (!min_field) // Nothing to remove.
		{
			if (aMode == RM_RemoveAt)
				for (pos = min_pos; pos < mKeyOffsetObject; ++pos)
					mFields[pos].key.i--;
			// Our return value when only one key is given is supposed to be the value
			// previously at this[key], which has just been removed.  Since this[key]
			// would return "", it makes sense to return an empty string in this case.
			aResultToken.symbol = SYM_STRING;	
			aResultToken.marker = _T("");
			return OK;
		}
		// Since only one field (at maximum) can be removed in this mode, it
		// seems more useful to return the field being removed than a count.
		switch (aResultToken.symbol = min_field->symbol)
		{
		case SYM_STRING:
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
		// If the key is an object, release it now because Free() doesn't.
		// Note that object keys can only be removed in the single-item mode.
		if (min_key_type == SYM_OBJECT)
			min_field->key.p->Release();
		// Set these up as if caller did Remove(min_key, min_key):
		max_pos = min_pos + 1;
		max_key.i = min_key.i; // Union copy. Used only if min_key_type == SYM_INTEGER; has no effect in other cases.
	}

	for (pos = min_pos; pos < max_pos; ++pos)
		// Free each field in the range being removed.
		mFields[pos].Free();

	if (min_key_type == SYM_STRING)
		// Free all string keys in the range being removed.
		for (pos = min_pos; pos < max_pos; ++pos)
			free(mFields[pos].key.s);

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
			if (aMode == RM_RemoveAt)
			{
				// Regardless of whether any fields were removed, min_pos contains the position of the field which
				// immediately followed the specified range.  Decrement each numeric key from this position onward.
				if (logical_count_removed > 0)
					for (pos = min_pos; pos < mKeyOffsetObject; ++pos)
						mFields[pos].key.i -= logical_count_removed;
			}
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

ResultType Object::_Remove(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	return _Remove_impl(aResultToken, aParam, aParamCount, RM_RemoveKey);
}

ResultType Object::_RemoveAt(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	return _Remove_impl(aResultToken, aParam, aParamCount, RM_RemoveAt);
}

ResultType Object::_Pop(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	// Unwanted parameters are ignored, as is conventional for dynamic calls.
	// _Remove_impl relies on aParamCount == 0 for Pop().
	return _Remove_impl(aResultToken, NULL, 0, RM_Pop);
}


ResultType Object::_MinIndex(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
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
	if (aParamCount)
	{
		SymbolType key_type;
		KeyType key;
		IndexType insert_pos;
		FieldType *field;

		if ( (field = FindField(*aParam[0], aResultToken.buf, /*out*/ key_type, /*out*/ key, /*out*/ insert_pos))
			&& field->symbol == SYM_STRING )
		{
			aResultToken.symbol = SYM_INTEGER;
			aResultToken.value_int64 = field->size ? _TSIZE(field->size - 1) : 0; // -1 to exclude null-terminator.
		}
		// else wrong type of field; leave aResultToken at default, empty string.
	}
	else // aParamCount == 0
	{
		aResultToken.symbol = SYM_INTEGER;
		aResultToken.value_int64 = mFieldCountMax;
	}
	return OK;
}

ResultType Object::_SetCapacity(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
// SetCapacity([field_name,] new_capacity)
{
	if (!aParamCount || !TokenIsNumeric(*aParam[aParamCount - 1]))
		return g_script.ScriptError(ERR_PARAM_INVALID);

	__int64 desired_capacity = TokenToInt64(*aParam[aParamCount - 1]);
	if (aParamCount >= 2) // Field name was specified.
	{
		if (desired_capacity < 0) // Check before sign is dropped.
			return g_script.ScriptError(ERR_PARAM2_INVALID);
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
			if (field->symbol != SYM_STRING)
				// Wrong type of field.
				return g_script.ScriptError(ERR_INVALID_VALUE);
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
				return OK;
			}
		}
		// Out of memory.  Throw an error, otherwise scripts might assume success and end up corrupting the heap:
		return g_script.ScriptError(ERR_OUTOFMEM);
	}
	// else aParamCount == 1: set the capacity of this object.
	IndexType desired_count = (IndexType)desired_capacity;
	if (desired_count < mFieldCount)
	{
		// It doesn't seem intuitive to allow SetCapacity to truncate the fields array, so just reallocate
		// as necessary to remove any unused space.  Allow negative values since SetCapacity(-1) seems more
		// intuitive than SetCapacity(0) when the contents aren't being discarded.
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
		return OK;
	}
	// At this point, failure isn't critical since nothing is being stored yet.  However, it might be easier to
	// debug if an error is thrown here rather than possibly later, when the array attempts to resize itself to
	// fit new items.  This also avoids the need for scripts to check if the return value is less than expected:
	return g_script.ScriptError(ERR_OUTOFMEM);
}

ResultType Object::_GetAddress(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
// GetAddress(key)
{
	if (!aParamCount)
		return g_script.ScriptError(ERR_TOO_FEW_PARAMS);
	
	SymbolType key_type;
	KeyType key;
	IndexType insert_pos;
	FieldType *field;

	if ( (field = FindField(*aParam[0], aResultToken.buf, /*out*/ key_type, /*out*/ key, /*out*/ insert_pos))
		&& field->symbol == SYM_STRING && field->size )
	{
		aResultToken.symbol = SYM_INTEGER;
		aResultToken.value_int64 = (__int64)field->marker;
	}
	// else field has no memory allocated; leave aResultToken at default, empty string.
	return OK;
}

ResultType Object::_NewEnum(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	IObject *newenum = new Enumerator(this);
	if (!newenum)
		return g_script.ScriptError(ERR_OUTOFMEM);
	aResultToken.symbol = SYM_OBJECT;
	aResultToken.object = newenum;
	return OK;
}

ResultType Object::_HasKey(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	if (!aParamCount)
		return g_script.ScriptError(ERR_TOO_FEW_PARAMS);
	
	SymbolType key_type;
	KeyType key;
	INT_PTR insert_pos;
	FieldType *field = FindField(**aParam, aResultToken.buf, key_type, key, insert_pos);
	aResultToken.symbol = SYM_INTEGER;
	aResultToken.value_int64 = (field != NULL);
	return OK;
}

ResultType Object::_Clone(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	Object *clone = Clone();
	if (!clone)
		return g_script.ScriptError(ERR_OUTOFMEM);	
	if (mBase)
		(clone->mBase = mBase)->AddRef();
	aResultToken.object = clone;
	aResultToken.symbol = SYM_OBJECT;
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
		symbol = SYM_STRING;
		marker = Var::sEmptyString;
		size = 0;
		return true;
	}
	
	if (len == -1)
		len = _tcslen(str);

	if (symbol != SYM_STRING || len >= size)
	{
		Free(); // Free object or previous buffer (which was too small).
		symbol = SYM_STRING;
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
		aParam.var->ToToken(temp);
		val = &temp;
	}
	else
		val = &aParam;

	switch (val->symbol)
	{
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
	if (symbol == SYM_STRING) {
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
		if (aParamCount && !_tcsicmp(ParamIndexToString(0), _T("Next")))
		{	// This is something like enum.Next(var); exclude "Next" so it is treated below as enum[var].
			++aParam; --aParamCount;
		}
		else
			return INVOKE_NOT_HANDLED;
	}
	Var *var0 = ParamIndexToOptionalVar(0);
	Var *var1 = ParamIndexToOptionalVar(1);
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
			case SYM_STRING:	aVal->AssignString(field.marker);	break;
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
	if (key_token.symbol == SYM_INTEGER
		|| (key_token.symbol == SYM_VAR && key_token.var->IsPureNumeric() == PURE_INTEGER))
	{	// Pure integer.
		key.i = (IntKeyType)TokenToInt64(key_token);
		key_type = SYM_INTEGER;
	}
	else if (key.p = TokenToObject(key_token))
	{	// SYM_OBJECT or SYM_VAR which contains an object.
		key_type = SYM_OBJECT;
	}
	else
	{	// SYM_STRING or SYM_VAR (confirmed not to be pure integer); or SYM_FLOAT.
		key.s = TokenToString(key_token, aBuf); // Pass aBuf to allow float -> string conversion.
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
	field.symbol = SYM_STRING;

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
		// Called explicitly by script, such as by "%obj.funcref%()" or "x := obj.funcref, x.()"
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
		&& g->CurrentFunc && g->CurrentFunc->mClass) // We're in a function defined within a class (i.e. a method).
	{
		if (IObject *this_class_base = g->CurrentFunc->mClass->Base())
		{
			ExprTokenType this_token;
			this_token.symbol = SYM_VAR;
			this_token.var = g->CurrentFunc->mParam[0].var;
			ResultType result = this_class_base->Invoke(aResultToken, this_token, (aFlags & ~IF_METAFUNC) | IF_METAOBJ, aParam, aParamCount);
			// Avoid returning INVOKE_NOT_HANDLED in this case so that our caller never
			// shows an "uninitialized var" warning for base.Foo() in a class method.
			if (result != INVOKE_NOT_HANDLED)
				return result;
		}
		return OK;
	}

	// Allow script-defined meta-functions to override the default behaviour:
	return Object::Invoke(aResultToken, aThisToken, aFlags, aParam, aParamCount);
}


#ifdef CONFIG_DEBUGGER

void ObjectBase::DebugWriteProperty(IDebugProperties *aDebugger, int aPage, int aPageSize, int aMaxDepth)
{
	DebugCookie cookie;
	aDebugger->BeginProperty(NULL, "object", 0, cookie);
	//if (aPage == 0)
	//{
	//	// This is mostly a workaround for debugger clients which make it difficult to
	//	// tell when a property contains an object with no child properties of its own:
	//	aDebugger->WriteProperty("Note", _T("This object doesn't support debugging."));
	//}
	aDebugger->EndProperty(cookie);
}

void Func::DebugWriteProperty(IDebugProperties *aDebugger, int aPage, int aPageSize, int aMaxDepth)
{
	DebugCookie cookie;
	aDebugger->BeginProperty(NULL, "object", 1, cookie);
	if (aPage == 0)
		aDebugger->WriteProperty("Name", mName);
	aDebugger->EndProperty(cookie);
}

#endif
