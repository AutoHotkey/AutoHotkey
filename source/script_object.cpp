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
//	buf points to a buffer the called function may use: char[MAX_NUMBER_SIZE]
//	circuit_token is NULL; if it is non-NULL on return, caller (or caller's caller) is responsible for it.
// Caller is responsible for making a persistent copy of the result, if appropriate.
{
	if (aParamCount < aFunc.mMinParams)
	{
		aResultToken.symbol = SYM_STRING;
		aResultToken.marker = "";
		return FAIL;
	}
	ResultType result = OK;

	// Code heavily based on SYM_FUNC handling in script_expression.cpp; see there for detailed comments.
	if (aFunc.mIsBuiltIn)
	{
		aResultToken.symbol = SYM_INTEGER; // Set default return type so that functions don't have to do it if they return INTs.
		aResultToken.marker = aFunc.mName;  // Inform function of which built-in function called it (allows code sharing/reduction). Can't use circuit_token because it's value is still needed later below.

		// CALL THE BUILT-IN FUNCTION:
		aFunc.mBIF(aResultToken, aParam, aParamCount);
	}
	else // It's not a built-in function.
	{
		int j, count_of_actuals_that_have_formals, var_backup_count;
		VarBkp *var_backup = NULL;

		// L: Set a default here in case we return early/abort.
		aResultToken.symbol = SYM_STRING;
		aResultToken.marker = "";

		count_of_actuals_that_have_formals = (aParamCount > aFunc.mParamCount)
					? aFunc.mParamCount  // Omit any actuals that lack formals (this can happen when a dynamic call passes too many parameters).
					: aParamCount;

		if (aFunc.mInstances > 0)
		{
			// Backup/restore of function's variables is needed.
			for (j = 0; j < count_of_actuals_that_have_formals; ++j) // For each actual parameter than has a formal.
			{
				ExprTokenType &this_stack_token = *aParam[j];
				if (this_stack_token.symbol == SYM_VAR && !aFunc.mParam[j].is_byref)
					this_stack_token.var->TokenToContents(this_stack_token);
			}
			if (!Var::BackupFunctionVars(aFunc, var_backup, var_backup_count)) // Out of memory.
			{
				//LineError(ERR_OUTOFMEM ERR_ABORT, FAIL, aFunc.mName);
				return FAIL;
			}
		}

		for (j = aParamCount; j < aFunc.mParamCount; ++j) // For each formal parameter that lacks an actual, provide a default value.
		{
			FuncParam &this_formal_param = aFunc.mParam[j];
			if (this_formal_param.is_byref)
				this_formal_param.var->ConvertToNonAliasIfNecessary();
			switch(this_formal_param.default_type)
			{
			case PARAM_DEFAULT_STR:   this_formal_param.var->Assign(this_formal_param.default_str);    break;
			case PARAM_DEFAULT_INT:   this_formal_param.var->Assign(this_formal_param.default_int64);  break;
			case PARAM_DEFAULT_FLOAT: this_formal_param.var->Assign(this_formal_param.default_double); break;
			//case PARAM_DEFAULT_NONE: Not possible due to the nature of this loop and due to earlier validation.
			}
		}

		for (j = 0; j < count_of_actuals_that_have_formals; ++j) // For each actual parameter than has a formal, assign the actual to the formal.
		{
			ExprTokenType &token = *aParam[j];
			if (!IS_OPERAND(token.symbol)) // Haven't found a way to produce this situation yet, but safe to assume it's possible.
			{
				Var::FreeAndRestoreFunctionVars(aFunc, var_backup, var_backup_count);
				return FAIL;
			}
			if (aFunc.mParam[j].is_byref)
			{
				if (token.symbol != SYM_VAR)
				{
					if (j < aFunc.mMinParams || token.value_int64 != aFunc.mParam[j].default_int64)
					{
						Var::FreeAndRestoreFunctionVars(aFunc, var_backup, var_backup_count);
						return FAIL;
					}
					aFunc.mParam[j].var->ConvertToNonAliasIfNecessary(); // L.
				}
				else
				{
					aFunc.mParam[j].var->UpdateAlias(token.var);
					continue;
				}
			}
			aFunc.mParam[j].var->Assign(token);
		}

		result = aFunc.Call(&aResultToken); // Call the UDF.

		if ( !(result == EARLY_EXIT || result == FAIL) )
		{
			if (aResultToken.symbol == SYM_STRING || aResultToken.symbol == SYM_OPERAND) // SYM_VAR is not currently possible.
			{
				char *buf;
				size_t len;
				// Make a persistent copy of the string in case it is the contents of one of the function's local variables.
				if ( *aResultToken.marker && (buf = (char *)malloc(1 + (len = strlen(aResultToken.marker)))) )
				{
					aResultToken.marker = (char *)memcpy(buf, aResultToken.marker, len + 1);
					aResultToken.circuit_token = (ExprTokenType *)buf;
					aResultToken.buf = (char *)len; // L33: Bugfix - buf is the length of the string, not the size of the memory allocation.
				}
				else
					aResultToken.marker = "";
			}
		}
		Var::FreeAndRestoreFunctionVars(aFunc, var_backup, var_backup_count);
	}
	return result;
}
	

//
// Object::Create - Called by BIF_Object to create a new object, optionally passing key/value pairs to set.
//

IObject *Object::Create(ExprTokenType *aParam[], int aParamCount)
{
	if (aParamCount & 1)
		return NULL; // Odd number of parameters - reserved for future use.

	Object *obj = new Object();
	if (obj && aParamCount)
	{
		ExprTokenType result_token, this_token;
		char buf[MAX_NUMBER_SIZE];

		this_token.symbol = SYM_OBJECT;
		this_token.object = obj;
		
		for (int i = 0; i + 1 < aParamCount; i += 2)
		{
			result_token.symbol = SYM_STRING;
			result_token.marker = "";
			result_token.circuit_token = NULL;
			result_token.buf = buf;

			// This is used rather than a more direct approach to ensure it is equivalent to assignment.
			// For instance, Object("base",MyBase,"a",1,"b",2) invokes meta-functions contained by MyBase.
			// For future consideration: Maybe it *should* bypass the meta-mechanism?
			obj->Invoke(result_token, this_token, IT_SET, aParam + i, 2);

			if (result_token.symbol == SYM_OBJECT) // L33: Bugfix.  Invoke must assume the result will be used and as a result we must account for this object reference:
				result_token.object->Release();
			if (result_token.circuit_token) // Currently should never happen, but may happen in future.
				free(result_token.circuit_token);
		}
	}
	return obj;
}


//
// Object::Delete - Called immediately before the object is deleted.
//					Returns false if object should not be deleted yet.
//

bool Object::Delete()
{
	if (mBase)
	{
		ExprTokenType result_token, this_token, param_token, *param;
		
		result_token.marker = "";
		result_token.symbol = SYM_STRING;
		result_token.circuit_token = NULL;

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
		if (result_token.circuit_token)
			free(result_token.circuit_token);
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
	int insert_pos;

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
				return OK; // TODO: Detection of 'return' vs 'return empty_value'.
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
			ResultType r = mBase->Invoke(aResultToken, aThisToken, aFlags | IF_META, aParam, aParamCount);
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
				// TODO: Move these predefined methods to built-in functions (which can be reimplemented as methods via a base object).
				if (aParamCount < 4 && *key.s == '_')
				{
					char *name = key.s + 1; // + 1 to exclude '_' from further consideration.
					++aParam; --aParamCount; // Exclude the method identifier.  A prior check ensures there was at least one param in this case.
					if (aParamCount) {
						if (!stricmp(name, "Insert"))
							return _Insert(aResultToken, aParam, aParamCount);
						if (!stricmp(name, "Remove"))
							return _Remove(aResultToken, aParam, aParamCount);
						if (!stricmp(name, "SetCapacity"))
							return _SetCapacity(aResultToken, aParam, aParamCount);
					} else { // aParamCount == 0
						if (!stricmp(name, "MaxIndex"))
							return _MaxIndex(aResultToken);
						if (!stricmp(name, "MinIndex"))
							return _MinIndex(aResultToken);
						if (!stricmp(name, "GetCapacity"))
							return _GetCapacity(aResultToken);
					}
					// For maintability: explicitly return since above has done ++aParam, --aParamCount.
					return INVOKE_NOT_HANDLED;
				}
				// Fall through and return INVOKE_NOT_HANDLED.
			}
			//
			// BUILT-IN "BASE" PROPERTY
			//
			else if (param_count_excluding_rvalue == 1 && !stricmp(key.s, "base"))
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
			if (key_type == SYM_STRING && !stricmp(key.s, "base"))
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
		if (!IS_INVOKE_META)
		{
			ExprTokenType &value_param = *aParam[1];
			// L34: Assigning an empty string no longer removes the field.
			if ( (field || (field = Insert(key_type, key, insert_pos))) && field->Assign(value_param) )
			{
				if (value_param.symbol == SYM_OPERAND || value_param.symbol == SYM_STRING)
				{
					// L33: Use value_param since our copy may be freed prematurely in some (possibly rare) cases:
					aResultToken.symbol = value_param.symbol;
					aResultToken.marker = value_param.marker;
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
			char *buf;
			size_t len;
			// L33: Make a persistent copy; our copy might be freed indirectly by releasing this object.
			//		Prior to L33, callers took care of this UNLESS this was the last op in an expression.
			if (buf = (char *)malloc(1 + (len = strlen(field->marker))))
			{
				aResultToken.marker = (char *)memcpy(buf, field->marker, len + 1);
				aResultToken.circuit_token = (ExprTokenType *)buf;
				aResultToken.buf = (char *)len;
			}
			else
				aResultToken.marker = "";

			aResultToken.symbol = SYM_OPERAND;
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
		ResultType r = aField->object->Invoke(aResultToken, field_token, IT_CALL, aParam, aParamCount);
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

ResultType Object::_Insert(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
// _Insert( key, value )
{
	if (aParamCount != 2)
		return OK;

	SymbolType key_type;
	KeyType key;
	int insert_pos, pos;
	FieldType *field;

	if ( (field = FindField(*aParam[0], aResultToken.buf, key_type, key, insert_pos)) && key_type == SYM_INTEGER )
	{
		// Since we were given a numeric key, we want to insert a new field here and increment this and any subsequent keys.
		insert_pos = field - mFields;
		// Signal below to insert a new field:
		field = NULL;
	}
	//else: specified field doesn't exist or has a non-numeric key; in the latter case we will simply overwrite it.

	if ( field || (field = Insert(key_type, key, insert_pos)) )
	{
		// Assign this field its new value:
		field->Assign(*aParam[1]);
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
// _Remove( min_key [, max_key ] )
{
	if (aParamCount < 1 || aParamCount > 2)
		return OK;

	FieldType *min_field, *max_field;
	int min_pos, max_pos, pos;
	SymbolType min_key_type, max_key_type;
	KeyType min_key, max_key;

	// Find the position of "min".
	if (min_field = FindField(*aParam[0], aResultToken.buf, min_key_type, min_key, min_pos))
		min_pos = min_field - mFields;
	
	if (aParamCount > 1)
	{
		// Find the position following "max".
		if (max_field = FindField(*aParam[1], aResultToken.buf, max_key_type, max_key, max_pos))
			max_pos = max_field - mFields + 1;
		// Since the order of key-types in mFields is of no logical consequence, require that both keys be the same type.
		// Do not allow removing a range of object keys since there is probably no meaning to their order.
		if (max_key_type != min_key_type || max_key_type == SYM_OBJECT || max_pos < min_pos
			// min and max are different types, are objects, or max < min.
			|| (max_pos == min_pos && (max_key_type == SYM_INTEGER ? max_key.i < min_key.i : stricmp(max_key.s, min_key.s) < 0)))
			// max < min, but no keys exist in that range so (max_pos < min_pos) check above didn't catch it.
			return OK;
		//else if (max_pos == min_pos): specified range is valid, but doesn't match any keys.
		//	Continue on, adjust integer keys as necessary and return 0 instead of "".
	}
	else
	{
		if (!min_field) // Nothing to remove.
		{
			// L34: Must not continue since min_pos points at the wrong key or an invalid location.
			// Empty result is reserved for invalid parameters; zero indicates no key(s) were found.
			aResultToken.symbol = SYM_INTEGER;	
			aResultToken.value_int64 = 0;
			return OK;
		}
		max_pos = min_pos + 1;
		max_key.i = min_key.i; // Used only if min_key_type == SYM_INTEGER; safe even in other cases.
	}

	for (pos = min_pos; pos < max_pos; ++pos)
		// Free each field in the range being removed.
		mFields[pos].Free();

	int remaining_fields = mFieldCount - max_pos;
	if (remaining_fields)
		// Move remaining fields left to fill the gap left by the removed range.
		memmove(mFields + min_pos, mFields + max_pos, remaining_fields * sizeof(FieldType));
	// Adjust count by the actual number of fields in the removed range.
	int actual_count_removed = max_pos - min_pos;
	mFieldCount -= actual_count_removed;
	// Adjust key offsets and numeric keys as necessary.
	if (min_key_type != SYM_STRING)
	{
		mKeyOffsetString -= actual_count_removed;
		if (min_key_type != SYM_OBJECT) // min_key_type == SYM_INTEGER
		{
			mKeyOffsetObject -= actual_count_removed;
			// Regardless of whether any fields were removed, min_pos contains the position of the field which
			// immediately followed the specified range.  Decrement each numeric key from this position onward.
			int logical_count_removed = max_key.i - min_key.i + 1;
			if (logical_count_removed > 0)
				for (pos = min_pos; pos < mKeyOffsetObject; ++pos)
					mFields[pos].key.i -= logical_count_removed;
		}
	}
	// Return actual number of fields removed:
	aResultToken.symbol = SYM_INTEGER;
	aResultToken.value_int64 = actual_count_removed;
	return OK;
}

ResultType Object::_MinIndex(ExprTokenType &aResultToken)
{
	if (mKeyOffsetObject) // i.e. there are fields with integer keys
	{
		aResultToken.symbol = SYM_INTEGER;
		aResultToken.value_int64 = (__int64)mFields[0].key.i;
	}
	// else no integer keys; leave aResultToken at default, empty string.
	return OK;
}

ResultType Object::_MaxIndex(ExprTokenType &aResultToken)
{
	if (mKeyOffsetObject) // i.e. there are fields with integer keys
	{
		aResultToken.symbol = SYM_INTEGER;
		aResultToken.value_int64 = (__int64)mFields[mKeyOffsetObject - 1].key.i;
	}
	// else no integer keys; leave aResultToken at default, empty string.
	return OK;
}

ResultType Object::_GetCapacity(ExprTokenType &aResultToken)
{
	aResultToken.symbol = SYM_INTEGER;
	aResultToken.value_int64 = mFieldCountMax;
	return OK;
}

ResultType Object::_SetCapacity(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
// _SetCapacity( new_capacity )
{
	if (aParamCount != 1 || !TokenIsPureNumeric(*aParam[aParamCount - 1]))
		// Invalid param(s); return default empty string.
		return OK;
	size_t desired_size = (size_t)TokenToInt64(*aParam[aParamCount - 1], TRUE);
	if (desired_size < (size_t)mFieldCount)
	{	// It doesn't seem intuitive to allow _SetCapacity to truncate the fields array.
		desired_size = (size_t)mFieldCount;
	}
	if (desired_size == mFieldCountMax || SetInternalCapacity((int)desired_size))
	{
		aResultToken.symbol = SYM_INTEGER;
		aResultToken.value_int64 = mFieldCountMax;
	}
	return OK;
}
	

// TODO: Enumeration/direct indexing/field-counting methods.


//
// Object::FieldType
//

bool Object::FieldType::Assign(char *str, size_t len, bool exact_size)
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

bool Object::FieldType::Assign(ExprTokenType &val)
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
// Object:: Internal Methods
//

template<typename T>
Object::FieldType *Object::FindField(T val, int left, int right, int &insert_pos)
// Template used below.  left and right must be set by caller to the appropriate bounds within mFields.
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

Object::FieldType *Object::FindField(SymbolType key_type, KeyType key, int &insert_pos)
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

Object::FieldType *Object::FindField(ExprTokenType &key_token, char *aBuf, SymbolType &key_type, KeyType &key, int &insert_pos)
// Searches for a field with the given key, where the key is a token passed from script.
{
	if (TokenIsPureNumeric(key_token) == PURE_INTEGER)
	{	// Treat all integer keys (even numeric strings) as pure integers for consistency and performance.
		key.i = (int)TokenToInt64(key_token, TRUE);
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
	
bool Object::SetInternalCapacity(int new_capacity)
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
	
Object::FieldType *Object::Insert(SymbolType key_type, KeyType key, int at)
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
	

//
// MetaObject::Invoke - Defines behaviour of object syntax when used on a non-object value.
//

//ResultType STDMETHODCALLTYPE MetaObject::Invoke(ExprTokenType &aResultToken, ExprTokenType &aThisToken, int aFlags, ExprTokenType *aParam[], int aParamCount)
//{
//	// Allow script-defined behaviour to take precedence:
//	ResultType r = Object::Invoke(aResultToken, aThisToken, aFlags | IF_META, aParam, aParamCount);
//
//	// TODO: Fix behaviour of "".base[x], ""["base",x] etc. or replace with built-in function to get g_MetaObject.
//	if (r == INVOKE_NOT_HANDLED && !IS_INVOKE_META && IS_INVOKE_GET && aParamCount == 1)
//	{
//		if (!stricmp(TokenToString(*aParam[0]), "base")) // L40: Allow any string value for consistency (previously only SYM_OPERAND was accepted).
//		{
//			// In this case, script did something like ("".base); i.e. wants a reference to g_MetaObject itself.
//			aResultToken.symbol = SYM_OBJECT;
//			aResultToken.object = this;
//			// No need to AddRef in this case since neither AddRef nor Release do anything.
//			return OK;
//		}
//	}
//
//	return r;
//}

MetaObject g_MetaObject;

char* Object::sMetaFuncName[] = { "__Get", "__Set", "__Call", "__Delete" };