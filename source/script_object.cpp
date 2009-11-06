#include "stdafx.h" // pre-compiled headers
#include "defines.h"
#include "globaldata.h"
#include "script.h"

#include "script_object.h"


/*
MISC IMPLEMENTATION IDEAS

Member Function Declarations:

	ProtoType.func() {
		...
	}

	This would be a convenient way to declare an anonymous function and assign it to a meta-object.
	There are a few possibilities for how this assignment takes place:

		1)	At load-time.  This would be consistent with regular function definitions, but would require 'ProtoType'
			to be a global/static variable (not an issue until function-in-function support is added).  If script does
			ProtoType := Object(), the function is lost.

		2)	At run-time.  It would resolve to something like:  ProtoType.func := <token containing anonymous Func>.
			Method definitions would need to be in the auto-execute section or an initialization function (which would
			require function-in-function support); this is somewhat inconsistent with regular function definitions.

		3)	Both.  Since the actual anonymous function would be created at load-time, it could be both assigned to
			the (automatically created) ProtoType object at load-time and at run-time (in case the ProtoType reference
			is overwritten.)
*/


//
//	DIRECT BUILT-IN FUNCTIONS
//


void BIF_ObjCreate(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	IObject *obj = NULL;

	if (aParamCount == 1) // L33: POTENTIALLY UNSAFE - Cast IObject address to object reference.
	{
		obj = (IObject *)TokenToInt64(*aParam[0]);
		if (obj < (IObject *)1024) // Prevent some obvious errors.
			obj = NULL;
		else
			obj->AddRef();
	}
	else
	{
		//char *type_name = aResultToken.marker + 3; // Omit "Obj" prefix
		//for (int t = 0; t < g_ObjTypeCount; ++t)
		//{
		//	if (!stricmp(type_name, g_ObjTypes[t].name))
		//	{
				//obj = g_ObjTypes[t].ctor(aParam, aParamCount);
				obj = Object::Create(aParam, aParamCount);
		//	}
		//}
	}

	if (obj)
	{
		aResultToken.symbol = SYM_OBJECT;
		aResultToken.object = obj;
		// DO NOT ADDREF: after we return, the only reference will be in aResultToken.
	}
	else
	{
		aResultToken.symbol = SYM_STRING;
		aResultToken.marker = _T("");
	}
}


void BIF_IsObject(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
// IsObject(obj) is currently equivalent to (obj && obj=""), but much more intuitive.
{
	int i;
	for (i = 0; i < aParamCount && TokenToObject(*aParam[i]); ++i);
	aResultToken.value_int64 = (__int64)(i == aParamCount); // TRUE if all are objects.
}


void BIF_ObjInvoke(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
    int invoke_type;
    IObject *obj;
    ExprTokenType *obj_param;

	// Since ObjGet/ObjSet/ObjCall are "pre-loaded" before the script begins executing,
	// marker always contains the correct case (vs. whatever the first call in script used).
	switch (aResultToken.marker[3])
	{
	case 'G': invoke_type = IT_GET; break;
	case 'S': invoke_type = IT_SET; break;
	default: invoke_type = IT_CALL;
	}

	// Set default return value; ONLY AFTER THE ABOVE.
	aResultToken.symbol = SYM_STRING;
	aResultToken.marker = _T("");
    
    obj_param = *aParam; // aParam[0].  Load-time validation has ensured at least one parameter was specified.
    
    if ( obj = TokenToObject(*obj_param) )
	{
		bool param_is_var = obj_param->symbol == SYM_VAR;
		if (param_is_var)
			// Since the variable may be cleared as a side-effect of the invocation, call AddRef to ensure the object does not expire prematurely.
			// This is not necessary for SYM_OBJECT since that reference is already counted and cannot be released before we return.  Each object
			// could take care not to delete itself prematurely, but it seems more proper, more reliable and more maintainable to handle it here.
			obj->AddRef();
        obj->Invoke(aResultToken, *obj_param, invoke_type, aParam + 1, aParamCount - 1);
		if (param_is_var)
			obj->Release();
	}
	else
	{
		if ( g_MetaObject.Invoke(aResultToken, *obj_param, invoke_type | IF_META, aParam + 1, aParamCount - 1) == INVOKE_NOT_HANDLED )
		{
			// Temporarily insert "__Get" or "__Set" or "__Call" into aParam to invoke the meta-function.
			// Note that *aParam (aParam[0]) points to the object-value, which is passed as a separate param.
			*aParam = &g_MetaFuncId[invoke_type];
			g_MetaObject.Invoke(aResultToken, *obj_param, IT_CALL_METAFUNC, aParam, aParamCount);
			*aParam = obj_param;
		}
	}
}



//
//	INTERNAL HELPER FUNCTIONS
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
		aResultToken.marker = _T("");
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
		aResultToken.marker = _T("");

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
				LPTSTR buf;
				size_t len;
				// Make a persistent copy of the string in case it is the contents of one of the function's local variables.
				if ( *aResultToken.marker && (buf = tmalloc(1 + (len = _tcslen(aResultToken.marker)))) )
				{
					aResultToken.marker = tmemcpy(buf, aResultToken.marker, len + 1);
					aResultToken.circuit_token = (ExprTokenType *)buf;
					aResultToken.buf = (LPTSTR)len; // L33: Bugfix - buf is the length of the string, not the size of the memory allocation.
				}
				else
					aResultToken.marker = _T("");
			}
		}
		Var::FreeAndRestoreFunctionVars(aFunc, var_backup, var_backup_count);
	}
	return result;
}
	


//
// Object
//


IObject *Object::Create(ExprTokenType *aParam[], int aParamCount)
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


bool Object::Delete()
// Called immediately before the object is deleted.  Allows deletion to be prevented if appropriate (see below).
{
	if (mBase)
	{
		ExprTokenType result_token, this_token, *param;
		
		result_token.marker = _T("");
		result_token.symbol = SYM_STRING;
		result_token.circuit_token = NULL;

		this_token.symbol = SYM_OBJECT;
		this_token.object = this;

		param = &g_MetaFuncId[3]; // "__Delete"

		// L33: Privatize the last recursion layer's deref buffer in case it is in use by our caller.
		// It's done here rather than in Var::FreeAndRestoreFunctionVars or CallFunc (even though the
		// below might not actually call any script functions) because this function is probably
		// executed much less often in most cases.
		PRIVATIZE_S_DEREF_BUF;

		mBase->Invoke(result_token, this_token, IT_CALL_METAFUNC, &param, 1);

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


ResultType STDMETHODCALLTYPE Object::Invoke(
                                            ExprTokenType &aResultToken,
                                            ExprTokenType &aThisToken,
                                            int aFlags,
                                            ExprTokenType *aParam[],
                                            int aParamCount
                                            )
{
	SymbolType key_type;
	KeyType key;
    FieldType *field;
	int insert_pos;
	
	int param_count_excluding_rvalue = aParamCount;
	if (IS_INVOKE_SET)
		--param_count_excluding_rvalue; // Prior ObjSet validation ensures the result won't be negative.

	if (param_count_excluding_rvalue)
	{
		key_type = TokenToKey(*aParam[0], /*out*/ key);
		field = FindField(key_type, key, /*out*/ insert_pos);
	}
	else
	{
		key_type = SYM_INVALID; // Allow key_type checks below without requiring that param_count_excluding_rvalue also be checked.
		field = NULL;
	}
    
	if (!field)
	{
		if (mBase)
		{
			ResultType r;
			if (!IS_INVOKE_SET)
			{
				r = mBase->Invoke(aResultToken, aThisToken, aFlags | IF_META, aParam, aParamCount);
				if (r != INVOKE_NOT_HANDLED)
					return r;
			}
			//else: IS_INVOKE_SET.  Probably not useful to pass it to mBase, except via __Set below.
			// Also seems best to avoid doing anything that could indirectly invalidate insert_pos.

			if (!IS_INVOKE_META) // This is the real, top-level object.  Since no mBase contains this field, let it be handled via __Get/__Set/__Call.
			{
				ExprTokenType *tmp = aParam[-1];
				// Relies on the fact that aParam always points at a location in (not at the beginning of) caller's aParam:
				aParam[-1] = &g_MetaFuncId[INVOKE_TYPE];
				r = mBase->Invoke(aResultToken, aThisToken, IT_CALL_METAFUNC, aParam - 1, aParamCount + 1);
				aParam[-1] = tmp;
				
				if (r == EARLY_RETURN)
					return OK; // TODO: Detection of 'return' vs 'return empty_value'.

				if (!param_count_excluding_rvalue)
					// This might result in slightly smaller code than letting it fall through and return below:
					return INVOKE_NOT_HANDLED;

				// Since the above may have inserted or removed fields (including the specified one),
				// insert_pos may no longer be correct or safe (and field=NULL may also be incorrect).
				field = FindField(key_type, key, /*out*/ insert_pos);
			}
		}
		if (IS_INVOKE_META // Since there's no field to get or call and a meta-invoke should not directly create a field, go no further.
			// OTHER SECTIONS BELOW RELY ON THIS RETURNING AS field IS UNINITIALIZED:
			|| !param_count_excluding_rvalue) // Something like obj[] or obj[]:=v which we only support through __Get/__Set/__Call above, go no further.
			return INVOKE_NOT_HANDLED;
	}

	if (IS_INVOKE_CALL)
	{
		if (field)
		{
			if (field->symbol == SYM_OBJECT)
			{
				ExprTokenType field_token;
				field_token.symbol = SYM_OBJECT;
				field_token.object = field->object;
				ExprTokenType *tmp = aParam[0];
				// Something must be inserted into the parameter list to remove any ambiguity between an intentionally
				// and directly called function of 'that' object and one of our parameters matching an existing name.
				// Rather than inserting something like an empty string, it seems more useful to insert 'this' object,
				// allowing 'that' to change (via __Call) the behaviour of a "function-call" which operates on 'this'.
				// Consequently, if 'that[this]' contains a value, it is invoked; seems obscure but rare, and could
				// also be of use (for instance, as a means to remove the 'this' parameter or replace it with 'that').
				aParam[0] = &aThisToken;
				ResultType r = field->object->Invoke(aResultToken, field_token, IT_CALL, aParam, aParamCount);
				aParam[0] = tmp;
				return r;
			}
			else if (field->symbol == SYM_OPERAND)
			{
				Func *func = g_script.FindFunc(field->marker);
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
		}
		// Since above has not handled this call, check for built-in methods.
		else if (!IS_INVOKE_META && aParamCount < 4 && key_type == SYM_STRING && *key.s == '_')
		// !IS_INVOKE_META: these methods must operate only on the real/direct object, not indirectly on a meta-object.
		{
			LPTSTR name = key.s + 1; // + 1 to exclude '_' from further consideration.
			++aParam; --aParamCount; // Exclude the method identifier.  A prior check ensures there was at least one param in this case.
			if (aParamCount) {
				if (!_tcsicmp(name, _T("Insert")))
					return _Insert(aResultToken, aParam, aParamCount);
				if (!_tcsicmp(name, _T("Remove")))
					return _Remove(aResultToken, aParam, aParamCount);
				if (!_tcsicmp(name, _T("SetCapacity")))
					return _SetCapacity(aResultToken, aParam, aParamCount);
				//if (!_tcsicmp(name, _T("GetAddress")))
				//	return _GetAddress(aResultToken, aParam, aParamCount);
			} else { // aParamCount == 0
				if (!_tcsicmp(name, _T("MaxIndex")))
					return _MaxIndex(aResultToken);
				if (!_tcsicmp(name, _T("MinIndex")))
					return _MinIndex(aResultToken);
				if (!_tcsicmp(name, _T("GetCapacity")))
					return _GetCapacity(aResultToken);
			}
		}
		return INVOKE_NOT_HANDLED;
	}
	else if (param_count_excluding_rvalue > 1)
	{
		// This is something like this[x,y] or this[x,y]:=z.  Since it wasn't handled by a meta-mechanism above,
		// handle only the "x" part (automatically creating and storing an object if this[x] didn't already exist
		// and an assignment is being made) and recursively invoke.  This has at least two benefits:
		//	1) Objects natively function as multi-dimensional arrays.
		//	2) Automatic initialization of object-fields.  For instance, this["base","__Get"]:="MyObjGet" does not require a prior this.base:=Object().
		IObject *obj = NULL;
		if (field)
		{
			if (field->symbol == SYM_OBJECT)
				// AddRef not used.  See below.
				obj = field->object;
		}
		else if (IS_INVOKE_SET)
		{
			if (key_type == SYM_STRING && !_tcsicmp(key.s, _T("base")))
			{
				if (!mBase)
					mBase = new Object();
				obj = mBase; // If NULL, above failed and below will detect it.
			}
			else
			{
				Object *new_obj = new Object();
				if (new_obj)
				{
					if ( field = Insert(key_type, key, insert_pos) )
					{	// Don't do field->Assign() since it would do AddRef() and we would need to counter with Release().
						field->symbol = SYM_OBJECT;
						field->object = obj = new_obj;
						if (mBase)
						{
							mBase->AddRef();
							new_obj->mBase = mBase;
						}
					}
					else
					{	// Create() succeeded but Insert() failed, so free the newly created obj.
						new_obj->Release();
					}
				}
			}
		}
		if (!obj) // Object was not successfully found or created.
			return OK;
		// obj now contains a pointer to the object contained by this field, possibly newly created above.
		ExprTokenType obj_token;
		obj_token.symbol = SYM_OBJECT;
		obj_token.object = obj;
		// References in obj_token and obj weren't counted (AddRef wasn't called), so Release() does not
		// need to be called before returning, and accessing obj after calling Invoke() would not be safe
		// since it could Release() the object (by overwriting our field via script) as a side-effect.
		return obj->Invoke(aResultToken, obj_token, aFlags & ~IF_META, aParam + 1, aParamCount - 1); // L34: Remove IF_META to allow the object to invoke its own meta-functions.
		// Above may return INVOKE_NOT_HANDLED in cases such as obj[a,b] where obj[a] exists but obj[a][b] does not.
		// NEEDS CONFIRMATION: This means that if obj.base[a] exists, [b] may exist in either obj.base[a] or obj[a].
	}
	else if (IS_INVOKE_SET)
	{
		if (IS_INVOKE_META) // For now it does not seem intuitive to allow obj.n:=v to alter obj's meta-object's fields.
			return INVOKE_NOT_HANDLED; // Note this doesn't apply to obj.base.base:=foo since that is not via the meta-mechanism.

		ExprTokenType &value_param = *aParam[1];

		// Must be handled before inserting a new field:
		if (!field && key_type == SYM_STRING && !_tcsicmp(key.s, _T("base")))
		{
			IObject *obj = TokenToObject(value_param);
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

		//if ( TokenIsEmptyString(value_param)
		//	|| !(field || (field = Insert(key_type, key, insert_pos)))
		//	|| !field->Assign(value_param) ) // THIS LINE MAY HANDLE THE ASSIGNMENT
		//{
		//	// A string assignment failed due to low memory. Remove this field;
		//	// attempting to retrieve it later will return "" unless overridden by meta-object.
		//	if (field)
		//		Remove(field, key_type);

		//	// Ensure an empty string is returned (this may not be necessary if caller set a default).
		//	aResultToken.symbol = SYM_STRING;
		//	aResultToken.marker = "";
		//}
		//else
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
	//else: IS_INVOKE_GET
	if (field)
	{
		if (field->symbol == SYM_OPERAND)
		{
			LPTSTR buf;
			size_t len;
			// L33: Make a persistent copy; our copy might be freed indirectly by releasing this object.
			//		Prior to L33, callers took care of this UNLESS this was the last op in an expression.
			if (buf = tmalloc(1 + (len = _tcslen(field->marker))))
			{
				aResultToken.marker = tmemcpy(buf, field->marker, len + 1);
				aResultToken.circuit_token = (ExprTokenType *)buf;
				aResultToken.buf = (LPTSTR)len;
			}
			else
				aResultToken.marker = _T("");

			aResultToken.symbol = SYM_OPERAND;
		}
		else
			field->Get(aResultToken);

		return OK;
	}
	if (key_type == SYM_STRING && !_tcsicmp(key.s, _T("base")))
	{
		// Above already returned if (!field && IS_INVOKE_META), so 'this' is the real object.
		if (mBase)
		{
			aResultToken.symbol = SYM_OBJECT;
			aResultToken.object = mBase;
			mBase->AddRef();
		}
		// else leave as empty string.
		return OK;
	}
	return INVOKE_NOT_HANDLED;
}


ResultType Object::_Insert(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
// _Insert( key, value )
{
	if (aParamCount != 2)
		return OK;

	SymbolType key_type;
	KeyType key;
	int insert_pos, pos;
	FieldType *field;

	if ( (field = FindField(*aParam[0], key_type, key, insert_pos)) && key_type == SYM_INTEGER )
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
	if (min_field = FindField(*aParam[0], min_key_type, min_key, min_pos))
		min_pos = min_field - mFields;
	
	if (aParamCount > 1)
	{
		// Find the position following "max".
		if (max_field = FindField(*aParam[1], max_key_type, max_key, max_pos))
			max_pos = max_field - mFields + 1;
		// Since the order of key-types in mFields is of no logical consequence, require that both keys be the same type.
		// Do not allow removing a range of object keys since there is probably no meaning to their order.
		if (max_key_type != min_key_type || max_key_type == SYM_OBJECT || max_pos < min_pos
			// min and max are different types, are objects, or max < min.
			|| (max_pos == min_pos && (max_key_type == SYM_INTEGER ? max_key.i < min_key.i : _tcsicmp(max_key.s, min_key.s) < 0)))
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

ResultType Object::_GetCapacity(ExprTokenType &aResultToken)//, ExprTokenType *aParam[], int aParamCount)
{
	//if (aParamCount == 1)
	//{
	//	FieldType *field = FindField(*aParam[0]);
	//	if (field && field->symbol == SYM_OPERAND)
	//	{
	//		aResultToken.symbol = SYM_INTEGER;
	//		aResultToken.value_int64 = field->size - 1; // -1 for reserved null-terminator byte.
	//	}
	//	// else wrong type of field; leave aResultToken at default, empty string.
	//}
	//else if (aParamCount == 0)
	//{
		aResultToken.symbol = SYM_INTEGER;
		aResultToken.value_int64 = mFieldCountMax;
	//}
	return OK;
}

ResultType Object::_SetCapacity(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
// _SetCapacity( new_capacity )
{
	if (aParamCount != 1 || !TokenIsPureNumeric(*aParam[aParamCount - 1]))
		// Invalid param(s); return default empty string.
		return OK;
	size_t desired_size = (size_t)TokenToInt64(*aParam[aParamCount - 1], TRUE);
	//if (desired_size < 0) // Sanity check.
	//	return OK;

	//if (aParamCount > 1)
	//{
	//	SymbolType key_type;
	//	KeyType key;
	//	int insert_pos;
	//	FieldType *field;
	//	char *buf;

	//	if (field = FindField(*aParam[0], key_type, key, insert_pos))
	//	{
	//		if (field->symbol != SYM_OPERAND)
	//			// Wrong type of field.
	//			return OK;
	//		if ( !(desired_size || (desired_size = strlen(field->marker))) )
	//		{
	//			field->Free();
	//			Remove(field, key_type);
	//			aResultToken.symbol = SYM_INTEGER;
	//			aResultToken.value_int64 = 0;
	//			return OK;
	//		}
	//		++desired_size; // Like VarSetCapacity, always reserve one byte for null-terminator.
	//		// Unlike VarSetCapacity, allow fields to shrink; preserve existing data up to the lesser of the new size and old size.
	//		if (buf = (char *)realloc(field->marker, desired_size))
	//		{
	//			// Ensure the data is null-terminated.
	//			if (field->size < desired_size)
	//				buf[field->size] = '\0'; // Terminate at end of existing data, in case it is a string.
	//			else
	//				buf[desired_size - 1] = '\0'; // Terminate at end of new data; data was truncated.

	//			field->marker = buf;
	//			field->size = desired_size;
	//			// Return new size, minus one byte reserved for null-terminator.
	//			aResultToken.symbol = SYM_INTEGER;
	//			aResultToken.value_int64 = desired_size - 1;
	//		}
	//		//else out of memory.
	//		return OK;
	//	}
	//	else if (desired_size)
	//	{
	//		++desired_size; // Always reserve one byte for null-terminator.
	//		if (buf = (char *)malloc(desired_size))
	//		{
	//			if (field = Insert(key_type, key, insert_pos))
	//			{
	//				field->symbol = SYM_OPERAND;
	//				field->marker = buf;
	//				field->size = desired_size;
	//				*buf = '\0';
	//				// Return new size, minus one byte reserved for null-terminator.
	//				aResultToken.symbol = SYM_INTEGER;
	//				aResultToken.value_int64 = desired_size - 1;
	//			}
	//			else // Insertion failed.
	//				free(buf);
	//		}
	//		return OK;
	//	}
	//	// else desired_size == 0 but field doesn't exist, so nothing to do.
	//}
	//else // aParamCount == 0
	//{
		if (desired_size < (size_t)mFieldCount)
		{	// It doesn't seem intuitive to allow _SetCapacity to truncate the fields array.
			desired_size = (size_t)mFieldCount;
		}
		if (desired_size == mFieldCountMax || SetInternalCapacity((int)desired_size))
		{
			aResultToken.symbol = SYM_INTEGER;
			aResultToken.value_int64 = mFieldCountMax;
		}
	//}
	return OK;
}

//ResultType Object::_GetAddress(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
//// _GetAddress( key )
//{
//	if (aParamCount == 1)
//	{
//		FieldType *field = FindField(*aParam[0]);
//		if (field && field->symbol == SYM_OPERAND)
//		{
//			aResultToken.symbol = SYM_INTEGER;
//			aResultToken.value_int64 = (__int64)field->marker;
//		}
//		// else wrong type of field; leave aResultToken at default, empty string.
//	}
//	return OK;
//}
	

// TODO: Enumeration/direct indexing/field-counting methods.



//
// MetaObject: Defines behaviour of ObjGet/ObjSet/ObjCall when used with a non-object value.
//

ResultType STDMETHODCALLTYPE MetaObject::Invoke(ExprTokenType &aResultToken, ExprTokenType &aThisToken, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	// Allow script-defined behaviour to take precedence:
	ResultType r = Object::Invoke(aResultToken, aThisToken, aFlags, aParam, aParamCount);

	// If not already handled, check if user is attempting to set a field (or field of a field etc.) of an empty var.
	/*if (r == INVOKE_NOT_HANDLED && IS_INVOKE_SET && aThisToken.symbol == SYM_VAR && !aThisToken.var->HasContents()) // HasContents() vs HasObject() because it doesn't seem wise to automatically overwrite a non-empty value.
	{
		IObject *obj;
		// Create a new empty object and assign it to the var, then invoke the object.
		if (obj = Object::Create(NULL, 0))
		{
			aThisToken.var->Assign(obj);
			r = obj->Invoke(aResultToken, aThisToken, aFlags & ~IF_META, aIdCount, aParam, aParamCount);
			// Above a) completely handled the assignment, b) performed some action defined by g_MetaObject (this)
			// recursively or c) created a new field and object and returned it to our caller for further processing.
			// For instance, in an expression "emptyvar.x.y:=z", we have created obj and put it into "emptyvar" -
			// obj may have created a field "x" containing an object and returned it for our caller to process ".y:=z".
			obj->Release();
		}
		// Out of memory.  Probably best to abort.
		else return OK;
	}*/

	if (r == INVOKE_NOT_HANDLED && IS_INVOKE_GET && aParamCount == 1 && aParam[0]->symbol == SYM_OPERAND) // aParam[0]->symbol *should* always be SYM_OPERAND, but check anyway.
	{
		if (!_tcsicmp(aParam[0]->marker, _T("base")))
		{
			// In this case, script did something like ("".base); i.e. wants a reference to g_MetaObject itself.
			aResultToken.symbol = SYM_OBJECT;
			aResultToken.object = this;
			// No need to AddRef in this case since neither AddRef nor Release do anything.
			return OK;
		}
	}

	return r;
}