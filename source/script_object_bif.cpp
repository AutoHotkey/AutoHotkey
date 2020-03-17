#include "stdafx.h" // pre-compiled headers
#include "defines.h"
#include "globaldata.h"
#include "script.h"
#include "ScriptModules.h"
#include "script_object.h"
#include "script_func_impl.h"


extern BuiltInFunc *OpFunc_GetProp, *OpFunc_GetItem, *OpFunc_SetProp, *OpFunc_SetItem;

//
// Object()
//

BIF_DECL(BIF_Object)
{
	IObject *obj = Object::Create(aParam, aParamCount, &aResultToken);
	if (obj)
	{
		// DO NOT ADDREF: the caller takes responsibility for the only reference.
		_f_return(obj);
	}
	//else: an error was already thrown.
}


//
// BIF_Array - Array(items*)
//

BIF_DECL(BIF_Array)
{
	if (auto arr = Array::Create(aParam, aParamCount))
		_f_return(arr);
	_f_throw(ERR_OUTOFMEM);
}


//
// Map()
//

BIF_DECL(BIF_Map)
{
	if (aParamCount & 1)
		_f_throw(ERR_PARAM_COUNT_INVALID);
	auto obj = Map::Create(aParam, aParamCount);
	if (!obj)
		_f_throw(ERR_OUTOFMEM);
	_f_return(obj);
}
	

//
// IsObject
//

BIF_DECL(BIF_IsObject)
{
	_f_return_b(TokenToObject(*aParam[0]) != nullptr);
}
	

//
// Op_ObjInvoke - Handles the object operators: x.y, x[y], x.y(), x.y := z, etc.
//

BIF_DECL(Op_ObjInvoke)
{
    int invoke_type = _f_callee_id;
	bool must_be_handled = true;
    IObject *obj;
    ExprTokenType *obj_param;

	
	// Set default return value for Invoke().
	aResultToken.symbol = SYM_STRING;
	aResultToken.marker = _T("");
    obj_param = *aParam; // aParam[0].  Load-time validation has ensured at least one parameter was specified.
	++aParam;
	--aParamCount;

	// The following is used in place of TokenToObject to bypass #Warn UseUnset:
	if (obj_param->symbol == SYM_OBJECT)
		obj = obj_param->object;
	else if (obj_param->symbol == SYM_VAR && obj_param->var->HasObject())
		obj = obj_param->var->Object();
	else if (obj_param->symbol == SYM_SUPER)
	{
		if (!g->CurrentFunc || !g->CurrentFunc->mClass) // We're in a function defined within a class (i.e. a method).
			_f_return_FAIL; // Should have been detected at load time, so just abort.
		obj = g->CurrentFunc->mClass->Base();
		ASSERT(obj != nullptr); // Should always pass for classes created by a class definition.
		obj_param = (ExprTokenType *)alloca(sizeof(ExprTokenType));
		obj_param->symbol = SYM_VAR;
		obj_param->var = g->CurrentFunc->mParam[0].var; // this
		invoke_type |= IF_NO_SET_PROPVAL;
		// Maybe not the best for error-detection, but this allows calls such as base.__delete()
		// to work when the superclass has no definition, which avoids the need to check whether
		// the superclass defines it, and ensures that any definition added later is called.
		must_be_handled = false;
	}
	else if (obj_param->symbol == SYM_MODULE) // It is "dynamic" scope resolution.
	{
		aParam--;
		aParamCount++;
		Op_ModuleInvoke(aResultToken, aParam, aParamCount);
		return;
	}
	else // Non-object value.
	{
		// Op_ModuleInvoke uses code copied from this branch, maintain together
		obj = Object::ValueBase(*obj_param);
		invoke_type |= IF_NO_SET_PROPVAL;
	}
	TCHAR name_buf[MAX_NUMBER_SIZE];
	LPTSTR name = nullptr;
	if (invoke_type & IF_DEFAULT)
		invoke_type &= ~IF_DEFAULT;
	else
	{
		if (aParam[0]->symbol != SYM_MISSING)
		{
			name = TokenToString(*aParam[0], name_buf);
			if (!*name && TokenToObject(*aParam[0]))
				_f_throw_type(_T("String"), *aParam[0]);
		}
		++aParam;
		--aParamCount;
	}
    
	ResultType result;
	bool param_is_var = obj_param->symbol == SYM_VAR;
	if (param_is_var)
	{
		if (obj_param->var->IsUninitializedNormalVar())
			// Treat this as an error for now.
			_f_throw(WARNING_USE_UNSET_VARIABLE, obj_param->var->mName);
		obj->AddRef(); // Ensure obj isn't deleted during the call if the variable is reassigned.
	}
    result = obj->Invoke(aResultToken, invoke_type, name, *obj_param, aParam, aParamCount);
	if (param_is_var)
		obj->Release();
	// Op_ModuleInvoke uses code copied from the below if-else ladder, maintain together
	if (result == INVOKE_NOT_HANDLED && must_be_handled)
	{
		_f__ret(aResultToken.UnknownMemberError(*obj_param, invoke_type, name));
	}
	else if (result == FAIL || result == EARLY_EXIT) // For maintainability: SetExitResult() might not have been called.
	{
		aResultToken.SetExitResult(result);
	}
	else if (invoke_type & IT_SET)
	{
		aResultToken.Free();
		aResultToken.mem_to_free = NULL;
		auto &value = *aParam[aParamCount - 1];
		switch (value.symbol)
		{
		case SYM_VAR:
			value.var->ToToken(aResultToken);
			break;
		case SYM_OBJECT:
			value.object->AddRef();
		default:
			aResultToken.CopyValueFrom(value);
		}
	}
}

//
//	Called by Op_ObjInvoke to handle dynamic scope resolution, eg, MyModule.%expr%[p*] := value, MyModule.%'OtherModule'%.%'var'%, etc...
//

BIF_DECL(Op_ModuleInvoke)
{
	// Op_ObjInvoke have already set the result to "".
	int invoke_type = _f_callee_id;

	ExprTokenType* mod_param = *aParam; // Load-time validation has ensured at least one parameter was specified.
	++aParam;
	--aParamCount;

	if ( (invoke_type & IF_DEFAULT) // Invalid, eg, %'MyModule'%[x]
		|| !mod_param->mod )		// Invalid, referencing an optional module which wasn't loaded.
		_f_throw(ERR_SMODULES_INVALID_SCOPE_RESOLUTION);

	size_t name_length;
	LPTSTR name = ParamIndexToString(0, NULL, &name_length);	// The name of the module, variable or function to resolve.

	++aParam;	// Remove the function name from consideration. Might be put back below.
	--aParamCount;

	if (invoke_type & IT_CALL) // It is a function call, find the func and call it.
	{
		// This must be done before searching for a module since functions and modules can have the same name.
		Func* func = g_script.FindFunc(name, name_length, 0, mod_param->mod);
		if (!func)
		{
			// Invoke value base. Because MyModule.%'NonExistentFunc'%(p*) "should" be consistent with %'NonExistentFunc'%(p*).
			aParam--;
			aParamCount++;
			aResultToken.func->mFID = (BuiltInFunctionID)(invoke_type | IF_DEFAULT);
			Op_ObjInvoke(aResultToken, aParam, aParamCount);
			return;
		}
		func->Call(aResultToken, aParam, aParamCount, false); // "false" since variadic parameters has already been expanded.
		return;
	}
	// Trying to resolve either a module or a variable.
	ScriptModule* found = NULL;
	if ( (found = mod_param->mod->GetNestedModule(name, true))
		|| mod_param->mod->IsOptionalModule(name))
	{
		// It is a module.
		if ( (invoke_type & IT_SET) // Invalid, eg, MyModule.%MyOtherModule% := val
			|| !found)				// Invalid, referencing an optional module which wasn't loaded.
			_f_throw(ERR_SMODULES_INVALID_SCOPE_RESOLUTION);

		aResultToken.symbol = SYM_MODULE;
		aResultToken.mod = found;
	}
	else
	{	// It is a variable.
		if (!*name) // Eg, MyModule.%''%
			_f_throw(ERR_DYNAMIC_BLANK);
		Var* var = g_script.FindOrAddVar(name, name_length, VAR_GLOBAL, mod_param->mod);
		if (!var)
		{	// FindOrAddVar already displayed the error.
			aResultToken.SetExitResult(FAIL);
			return;
		}
		if (var->Type() == VAR_BUILTIN) // Currently doesn't support built-in vars.
			_f_throw(ERR_SMODULES_INVALID_SCOPE_RESOLUTION);
		if (invoke_type & IT_SET
			&& aParamCount == 1)
		{
			// Assiging a variable. Eg, MyModule.%expr% := value
			if (VAR_IS_READONLY(*var)) // For maintainability
				_f_throw(ERR_VAR_IS_READONLY);
			// Assign the value to the variable
			if (!(var->Assign(**aParam)))
			{
				// The error message should have been displayed, but the exit result must be set.
				aResultToken.SetExitResult(FAIL);
				return;
			}
			// Proceed to set the variable as the result of the assignment.
		}
		else if (aParamCount)
		{
			// Parameters were passed, eg MyModule.%expr%[p*] [ := value ]
			// This branch contains some code from Op_ObjInvoke, maintain together. Done this ways for convenince.
			ExprTokenType this_token;
			this_token.symbol = SYM_VAR;
			this_token.var = var;

			// Get the obj which will be invoked:
			IObject* obj;
			if (var->HasObject())
				obj = var->Object();
			else
			{
				// MyModule.%expr%[p*] should be equivalent to %expr%[p*] also when %expr% is resolved to a var which doesn't contain an object.
				obj = Object::ValueBase(*aParam[-1]);
				invoke_type |= IF_NO_SET_PROPVAL;
			}
			// Addref and Release unconditionally for brevity. Done to avoid the object being deleted in case the variable is reasigned during the operation.
			obj->AddRef();
			ResultType result = obj->Invoke(aResultToken, invoke_type, NULL, this_token, aParam, aParamCount);
			obj->Release();
			if (result == INVOKE_NOT_HANDLED)
			{
				_f__ret(aResultToken.UnknownMemberError(*aParam[-1], invoke_type, name));
			}
			else if (result == FAIL || result == EARLY_EXIT) // For maintainability: SetExitResult() might not have been called.
			{
				aResultToken.SetExitResult(result);
			}
			else if (invoke_type & IT_SET)
			{
				aResultToken.Free();
				aResultToken.mem_to_free = NULL;
				auto& value = *aParam[aParamCount - 1];
				switch (value.symbol)
				{
				case SYM_VAR:
					value.var->ToToken(aResultToken);
					break;
				case SYM_OBJECT:
					value.object->AddRef();
				default:
					aResultToken.CopyValueFrom(value);
				}
			}
			return;
		}
		// The result is a variable, eg, MyModule.%'var'% [:= value].
		aResultToken.symbol = SYM_VAR;
		aResultToken.var = var;
	}
	return;
}

//
// Op_ObjGetInPlace - Handles part of a compound assignment like x.y += z.
//

BIF_DECL(Op_ObjGetInPlace)
{
	// Since the most common cases have two params, the "param count" param is omitted in
	// those cases. Otherwise we have one visible parameter, which indicates the number of
	// actual parameters below it on the stack.
	aParamCount = aParamCount ? (int)TokenToInt64(*aParam[0]) : 2; // x[<n-1 params>] : x.y
	Op_ObjInvoke(aResultToken, aParam - aParamCount, aParamCount);
}


//
// Op_ObjIncDec - Handles pre/post-increment/decrement for object fields, such as ++x[y].
//

BIF_DECL(Op_ObjIncDec)
{
	SymbolType op = SymbolType(_f_callee_id & ~IF_DEFAULT);
	
	bool square_brackets = _f_callee_id & IF_DEFAULT;
	auto *get_func = square_brackets ? OpFunc_GetItem : OpFunc_GetProp;
	auto *set_func = square_brackets ? OpFunc_SetItem : OpFunc_SetProp;

	ResultToken temp_result;
	// Set the defaults expected by Op_ObjInvoke:
	temp_result.InitResult(aResultToken.buf);
	temp_result.symbol = SYM_INTEGER;
	temp_result.func = get_func;

	// Retrieve the current value.  Do it this way instead of calling Object::Invoke
	// so that if aParam[0] is not an object, ValueBase() is correctly invoked.
	Op_ObjInvoke(temp_result, aParam, aParamCount);

	if (temp_result.Exited()) // Implies no return value.
	{
		aResultToken.SetExitResult(temp_result.Result());
		return;
	}

	ExprTokenType current_value, value_to_set;
	switch (value_to_set.symbol = current_value.symbol = TokenIsNumeric(temp_result))
	{
	case PURE_INTEGER:
		value_to_set.value_int64 = (current_value.value_int64 = TokenToInt64(temp_result))
			+ ((op == SYM_POST_INCREMENT || op == SYM_PRE_INCREMENT) ? +1 : -1);
		break;

	case PURE_FLOAT:
		value_to_set.value_double = (current_value.value_double = TokenToDouble(temp_result))
			+ ((op == SYM_POST_INCREMENT || op == SYM_PRE_INCREMENT) ? +1 : -1);
		break;

	default: // PURE_NOT_NUMERIC == SYM_STRING.
		// Value is non-numeric, so throw.
		aResultToken.TypeError(_T("Number"), temp_result);
		temp_result.Free();
		return;
	}

	// Free the object or string returned by Op_ObjInvoke, if applicable.
	// It's theoretically possible for a numeric string to require this.
	temp_result.Free();

	// Although it's likely our caller's param array has enough space to hold the extra
	// parameter, there's no way to know for sure whether it's safe, so we allocate our own:
	ExprTokenType **param = (ExprTokenType **)_alloca((aParamCount + 1) * sizeof(ExprTokenType *));
	memcpy(param, aParam, aParamCount * sizeof(ExprTokenType *)); // Copy caller's param pointers.
	param[aParamCount++] = &value_to_set; // Append new value as the last parameter.

	if (op == SYM_PRE_INCREMENT || op == SYM_PRE_DECREMENT)
	{
		aResultToken.func = set_func;
		// Set the new value and pass the return value of the invocation back to our caller.
		// This should be consistent with something like x.y := x.y + 1.
		Op_ObjInvoke(aResultToken, param, aParamCount);
	}
	else // SYM_POST_INCREMENT || SYM_POST_DECREMENT
	{
		// Must be re-initialized (and must use SET rather than GET):
		temp_result.InitResult(aResultToken.buf);
		temp_result.symbol = SYM_INTEGER;
		temp_result.func = set_func;
		
		// Set the new value.
		Op_ObjInvoke(temp_result, param, aParamCount);

		if (temp_result.Exited()) // Implies no return value.
		{
			aResultToken.SetExitResult(temp_result.Result());
			return;
		}
		
		// Dispose of the result safely.
		temp_result.Free();

		// Return the previous value.
		aResultToken.symbol = current_value.symbol;
		aResultToken.value_int64 = current_value.value_int64; // Union copy.  Includes marker_length on x86.
#ifdef _WIN64
		aResultToken.marker_length = current_value.marker_length; // For simplicity, symbol isn't checked.
#endif
	}
}


//
// Functions for accessing built-in methods (even if obscured by a user-defined method).
//

BIF_DECL(BIF_ObjXXX)
{
	aResultToken.symbol = SYM_STRING;
	aResultToken.marker = _T(""); // Set default for CallBuiltin().
	
	Object *obj = dynamic_cast<Object*>(TokenToObject(*aParam[0]));
	if (obj)
		obj->CallBuiltin(_f_callee_id, aResultToken, aParam + 1, aParamCount - 1);
	else
		_f_throw_type(_T("Object"), *aParam[0]);
}


//
// ObjAddRef/ObjRelease - used with pointers rather than object references.
//

BIF_DECL(BIF_ObjAddRefRelease)
{
	IObject *obj = (IObject *)TokenToInt64(*aParam[0]);
	if (obj < (IObject *)65536) // Rule out some obvious errors.
		_f_throw(ERR_PARAM1_INVALID);
	if (_f_callee_id == FID_ObjAddRef)
		_f_return_i(obj->AddRef());
	else
		_f_return_i(obj->Release());
}


//
// ObjBindMethod(Obj, Method, Params...)
//

BIF_DECL(BIF_ObjBindMethod)
{
	IObject *func, *bound_func;
	if (  !(func = TokenToFunctor(*aParam[0]))  )
		_f_throw(ERR_PARAM1_INVALID);
	LPCTSTR name = nullptr;
	if (aParamCount > 1)
	{
		if (aParam[1]->symbol != SYM_MISSING)
			name = TokenToString(*aParam[1], _f_number_buf);
		aParam += 2;
		aParamCount -= 2;
	}
	else
		aParamCount = 0;
	bound_func = BoundFunc::Bind(func, IT_CALL, name, aParam, aParamCount);
	func->Release();
	if (!bound_func)
		_f_throw(ERR_OUTOFMEM);
	_f_return(bound_func);
}


//
// ObjPtr/ObjPtrAddRef/ObjFromPtr - Convert between object reference and IObject pointer.
//

BIF_DECL(BIF_ObjPtr)
{
	if (_f_callee_id >= FID_ObjFromPtr)
	{
		auto obj = (IObject *)ParamIndexToInt64(0);
		if (obj < (IObject *)65536) // Prevent some obvious errors.
			_f_throw(ERR_PARAM1_INVALID);
		if (_f_callee_id == FID_ObjFromPtrAddRef)
			obj->AddRef();
		_f_return(obj);
	}
	else // FID_ObjPtr or FID_ObjPtrAddRef.
	{
		auto obj = ParamIndexToObject(0);
		if (!obj)
			_f_throw_type(_T("object"), *aParam[0]);
		if (_f_callee_id == FID_ObjPtrAddRef)
			obj->AddRef();
		_f_return((UINT_PTR)obj);
	}
}


//
// ObjSetBase/ObjGetBase - Change or return Object's base without invoking any meta-functions.
//

BIF_DECL(BIF_Base)
{
	Object *obj = dynamic_cast<Object*>(TokenToObject(*aParam[0]));
	if (_f_callee_id == FID_ObjSetBase)
	{
		if (!obj)
			_f_throw_type(_T("Object"), *aParam[0]);
		auto new_base = dynamic_cast<Object *>(TokenToObject(*aParam[1]));
		if (!obj->SetBase(new_base, aResultToken))
			return;
	}
	else // ObjGetBase
	{
		Object *obj_base;
		if (obj)
			obj_base = obj->Base();
		else
			obj_base = Object::ValueBase(*aParam[0]);
		if (obj_base)
		{
			obj_base->AddRef();
			_f_return(obj_base);
		}
		// Otherwise, it's something with no base, so return "".
		// Could be Object::sAnyPrototype, a ComObject or perhaps SYM_MISSING.
	}
	_f_return_empty;
}


bool Object::HasBase(ExprTokenType &aValue, IObject *aBase)
{
	if (auto obj = dynamic_cast<Object *>(TokenToObject(aValue)))
	{
		return obj->IsDerivedFrom(aBase);
	}
	if (auto value_base = Object::ValueBase(aValue))
	{
		return value_base == aBase || value_base->IsDerivedFrom(aBase);
	}
	// Something that doesn't fit in our type hierarchy, like a ComObject.
	// Returning false seems correct and more useful than throwing.
	// HasBase(ComObj, "".base.base) ; False, it's not a primitive value.
	// HasBase(ComObj, Object.Prototype) ; False, it's not one of our Objects.
	return false;
}


BIF_DECL(BIF_HasBase)
{
	auto that_base = ParamIndexToObject(1);
	if (!that_base)
		_f_throw_type(_T("object"), *aParam[1]);
	_f_return_b(Object::HasBase(*aParam[0], that_base));
}


Object *ParamToObjectOrBase(ExprTokenType &aToken, ResultToken &aResultToken)
{
	Object *obj;
	if (  (obj = dynamic_cast<Object *>(TokenToObject(aToken)))
		|| (obj = Object::ValueBase(aToken))  )
		return obj;
	aResultToken.Error(ERR_PARAM1_INVALID);
	return nullptr;
}


BIF_DECL(BIF_HasProp)
{
	auto obj = ParamToObjectOrBase(*aParam[0], aResultToken);
	if (!obj)
		return;
	_f_return_b(obj->HasProp(ParamIndexToString(1, _f_number_buf)));
}


BIF_DECL(BIF_HasMethod)
{
	auto obj = ParamToObjectOrBase(*aParam[0], aResultToken);
	if (!obj)
		return;
	_f_return_b(obj->HasMethod(ParamIndexToString(1, _f_number_buf)));
}


BIF_DECL(BIF_GetMethod)
{
	auto obj = ParamToObjectOrBase(*aParam[0], aResultToken);
	if (!obj)
		return;
	obj->GetMethod(aResultToken, ParamIndexToString(1, _f_number_buf));
}
