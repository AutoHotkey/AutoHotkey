﻿#include "stdafx.h" // pre-compiled headers
#include "defines.h"
#include "globaldata.h"
#include "script.h"

#include "script_object.h"


//
// BIF_ObjCreate - Object()
//

BIF_DECL(BIF_ObjCreate)
{
	IObject *obj = NULL;

	if (aParamCount == 1) // L33: POTENTIALLY UNSAFE - Cast IObject address to object reference.
	{
		if (obj = TokenToObject(*aParam[0]))
		{	// Allow &obj == Object(obj), but AddRef() for equivalence with ComObjActive(comobj).
			obj->AddRef();
			aResultToken.value_int64 = (__int64)obj;
			return; // symbol is already SYM_INTEGER.
		}
		obj = (IObject *)TokenToInt64(*aParam[0]);
		if (obj < (IObject *)1024) // Prevent some obvious errors.
			obj = NULL;
		else
			obj->AddRef();
	}
	else
		obj = Object::Create(aParam, aParamCount);

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


//
// BIF_ObjArray - Array(items*)
//

BIF_DECL(BIF_ObjArray)
{
	if (aResultToken.object = Object::CreateArray(aParam, aParamCount))
	{
		aResultToken.symbol = SYM_OBJECT;
		return;
	}
	aResultToken.symbol = SYM_STRING;
	aResultToken.marker = _T("");
}
	

//
// BIF_IsObject - IsObject(obj)
//

BIF_DECL(BIF_IsObject)
{
	int i;
	for (i = 0; i < aParamCount && TokenToObject(*aParam[i]); ++i);
	aResultToken.value_int64 = (__int64)(i == aParamCount); // TRUE if all are objects.
}
	

//
// BIF_ObjInvoke - Handles ObjGet/Set/Call() and get/set/call syntax.
//

BIF_DECL(BIF_ObjInvoke)
{
    int invoke_type;
    IObject *obj;
    ExprTokenType *obj_param;

	// Since ObjGet/ObjSet/ObjCall are not publicly accessible as functions, Func::mName
	// (passed via aResultToken.marker) contains the actual flag rather than a name.
	invoke_type = (int)(INT_PTR)aResultToken.marker;

	// Set default return value; ONLY AFTER THE ABOVE.
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
	else
		obj = NULL;
    
    if (obj)
	{
		bool param_is_var = obj_param->symbol == SYM_VAR;
		if (param_is_var)
			// Since the variable may be cleared as a side-effect of the invocation, call AddRef to ensure the object does not expire prematurely.
			// This is not necessary for SYM_OBJECT since that reference is already counted and cannot be released before we return.  Each object
			// could take care not to delete itself prematurely, but it seems more proper, more reliable and more maintainable to handle it here.
			obj->AddRef();
        aResult = obj->Invoke(aResultToken, *obj_param, invoke_type, aParam, aParamCount);
		if (param_is_var)
			obj->Release();
	}
	// Invoke meta-functions of g_MetaObject.
	else if (INVOKE_NOT_HANDLED == (aResult = g_MetaObject.Invoke(aResultToken, *obj_param, invoke_type | IF_META, aParam, aParamCount)))
	{
		// Since above did not handle it, check for attempts to access .base of non-object value (g_MetaObject itself).
		if (   invoke_type != IT_CALL // Exclude things like "".base().
			&& aParamCount > (invoke_type == IT_SET ? 2 : 0) // SET is supported only when an index is specified: "".base[x]:=y
			&& !_tcsicmp(TokenToString(*aParam[0]), _T("base"))   )
		{
			if (aParamCount > 1)	// "".base[x] or similar
			{
				// Re-invoke g_MetaObject without meta flag or "base" param.
				ExprTokenType base_token;
				base_token.symbol = SYM_OBJECT;
				base_token.object = &g_MetaObject;
				g_MetaObject.Invoke(aResultToken, base_token, invoke_type, aParam + 1, aParamCount - 1);
			}
			else					// "".base
			{
				// Return a reference to g_MetaObject.  No need to AddRef as g_MetaObject ignores it.
				aResultToken.symbol = SYM_OBJECT;
				aResultToken.object = &g_MetaObject;
			}
		}
		else
		{
			// Since it wasn't handled (not even by g_MetaObject), maybe warn at this point:
			if (obj_param->symbol == SYM_VAR)
				obj_param->var->MaybeWarnUninitialized();
		}
	}
	if (aResult == INVOKE_NOT_HANDLED)
		aResult = OK;
}
	

//
// BIF_ObjGetInPlace - Handles part of a compound assignment like x.y += z.
//

BIF_DECL(BIF_ObjGetInPlace)
{
	// Since the most common cases have two params, the "param count" param is omitted in
	// those cases. Otherwise we have one visible parameter, which indicates the number of
	// actual parameters below it on the stack.
	aParamCount = aParamCount ? (int)TokenToInt64(*aParam[0]) : 2; // x[<n-1 params>] : x.y
	BIF_ObjInvoke(aResult, aResultToken, aParam - aParamCount, aParamCount);
}


//
// BIF_ObjNew - Handles "new" as in "new Class()".
//

BIF_DECL(BIF_ObjNew)
{
	aResultToken.symbol = SYM_STRING;
	aResultToken.marker = _T("");

	ExprTokenType *class_token = aParam[0]; // Save this to be restored later.

	IObject *class_object = TokenToObject(*class_token);
	if (!class_object)
		return;

	Object *new_object = Object::Create();
	if (!new_object)
		return;

	new_object->SetBase(class_object);

	ExprTokenType name_token, this_token;
	this_token.symbol = SYM_OBJECT;
	this_token.object = new_object;
	name_token.symbol = SYM_STRING;
	aParam[0] = &name_token;

	ResultType result;
	LPTSTR buf = aResultToken.buf; // In case Invoke overwrites it via the union.

	Line *curr_line = g_script.mCurrLine;

	// __Init was added so that instance variables can be initialized in the correct order
	// (beginning at the root class and ending at class_object) before __New is called.
	// It shouldn't be explicitly defined by the user, but auto-generated in DefineClassVars().
	name_token.marker = _T("__Init");
	result = class_object->Invoke(aResultToken, this_token, IT_CALL | IF_METAOBJ, aParam, 1);
	if (result != INVOKE_NOT_HANDLED)
	{
		// It's possible that __Init is user-defined (despite recommendations in the
		// documentation) or built-in, so make sure the return value, if any, is freed:
		if (aResultToken.symbol == SYM_OBJECT)
			aResultToken.object->Release();
		if (aResultToken.mem_to_free)
		{
			free(aResultToken.mem_to_free);
			aResultToken.mem_to_free = NULL;
		}
		// Reset to defaults for __New, invoked below.
		aResultToken.symbol = SYM_STRING;
		aResultToken.marker = _T("");
		aResultToken.buf = buf;
		if (result == FAIL || result == EARLY_EXIT)
		{
			new_object->Release();
			aParam[0] = class_token; // Restore it to original caller-supplied value.
			aResult = result;
			return;
		}
	}

	g_script.mCurrLine = curr_line; // Prevent misleading error reports/Exception() stack trace.
	
	// __New may be defined by the script for custom initialization code.
	name_token.marker = Object::sMetaFuncName[4]; // __New
	result = class_object->Invoke(aResultToken, this_token, IT_CALL | IF_METAOBJ, aParam, aParamCount);
	if (result == EARLY_RETURN || !TokenIsEmptyString(aResultToken))
	{
		// __New() returned a value in aResultToken, so use it as our result.  aResultToken is checked
		// for the unlikely case that class_object is not an Object (perhaps it's a ComObject) or __New
		// points to a built-in function.  The older behaviour for those cases was to free the unexpected
		// return value, but this requires less code and might be useful somehow.
		new_object->Release();
	}
	else if (result == FAIL || result == EARLY_EXIT)
	{
		// An error was raised within __New() or while trying to call it, or Exit was called.
		new_object->Release();
		aResult = result;
	}
	else
	{
		// Either it wasn't handled (i.e. neither this class nor any of its super-classes define __New()),
		// or there was no explicit "return", so just return the new object.
		aResultToken.symbol = SYM_OBJECT;
		aResultToken.object = new_object;
	}
	aParam[0] = class_token;
}


//
// BIF_ObjIncDec - Handles pre/post-increment/decrement for object fields, such as ++x[y].
//

BIF_DECL(BIF_ObjIncDec)
{
	// Func::mName (which aResultToken.marker is set to) has been overloaded to pass
	// the type of increment/decrement to be performed on this object's field.
	SymbolType op = (SymbolType)(INT_PTR)aResultToken.marker;

	ExprTokenType temp_result, current_value, value_to_set;

	// Set the defaults expected by BIF_ObjInvoke:
	temp_result.symbol = SYM_INTEGER;
	temp_result.marker = (LPTSTR)IT_GET;
	temp_result.buf = aResultToken.buf;
	temp_result.mem_to_free = NULL;

	// Retrieve the current value.  Do it this way instead of calling Object::Invoke
	// so that if aParam[0] is not an object, g_MetaObject is correctly invoked.
	BIF_ObjInvoke(aResult, temp_result, aParam, aParamCount);

	if (aResult == FAIL || aResult == EARLY_EXIT)
		return;

	// Change SYM_STRING to SYM_OPERAND so below may treat it as a numeric string.
	if (temp_result.symbol == SYM_STRING)
	{
		temp_result.symbol = SYM_OPERAND;
		temp_result.buf = NULL; // Indicate that this SYM_OPERAND token LACKS a pre-converted binary integer.
	}

	switch (value_to_set.symbol = current_value.symbol = TokenIsPureNumeric(temp_result))
	{
	case PURE_INTEGER:
		value_to_set.value_int64 = (current_value.value_int64 = TokenToInt64(temp_result))
			+ ((op == SYM_POST_INCREMENT || op == SYM_PRE_INCREMENT) ? +1 : -1);
		break;

	case PURE_FLOAT:
		value_to_set.value_double = (current_value.value_double = TokenToDouble(temp_result))
			+ ((op == SYM_POST_INCREMENT || op == SYM_PRE_INCREMENT) ? +1 : -1);
		break;
	}

	// Free the object or string returned by BIF_ObjInvoke, if applicable.
	if (temp_result.symbol == SYM_OBJECT)
		temp_result.object->Release();
	if (temp_result.mem_to_free)
		free(temp_result.mem_to_free);

	if (current_value.symbol == PURE_NOT_NUMERIC)
	{
		// Value is non-numeric, so assign and return "".
		value_to_set.symbol = SYM_STRING;
		value_to_set.marker = _T("");
		//current_value.symbol = SYM_STRING; // Already done (SYM_STRING == PURE_NOT_NUMERIC).
		current_value.marker = _T("");
	}

	// Although it's likely our caller's param array has enough space to hold the extra
	// parameter, there's no way to know for sure whether it's safe, so we allocate our own:
	ExprTokenType **param = (ExprTokenType **)_alloca((aParamCount + 1) * sizeof(ExprTokenType *));
	memcpy(param, aParam, aParamCount * sizeof(ExprTokenType *)); // Copy caller's param pointers.
	param[aParamCount++] = &value_to_set; // Append new value as the last parameter.

	if (op == SYM_PRE_INCREMENT || op == SYM_PRE_DECREMENT)
	{
		aResultToken.marker = (LPTSTR)IT_SET;
		// Set the new value and pass the return value of the invocation back to our caller.
		// This should be consistent with something like x.y := x.y + 1.
		BIF_ObjInvoke(aResult, aResultToken, param, aParamCount);
	}
	else // SYM_POST_INCREMENT || SYM_POST_DECREMENT
	{
		// Must be re-initialized (and must use IT_SET instead of IT_GET):
		temp_result.symbol = SYM_INTEGER;
		temp_result.marker = (LPTSTR)IT_SET;
		temp_result.buf = aResultToken.buf;
		temp_result.mem_to_free = NULL;
		
		// Set the new value.
		BIF_ObjInvoke(aResult, temp_result, param, aParamCount);
		
		// Dispose of the result safely.
		if (temp_result.symbol == SYM_OBJECT)
			temp_result.object->Release();
		if (temp_result.mem_to_free)
			free(temp_result.mem_to_free);

		// Return the previous value.
		aResultToken.symbol = current_value.symbol;
		aResultToken.value_int64 = current_value.value_int64; // Union copy.
	}
}


//
// Functions for accessing built-in methods (even if obscured by a user-defined method).
//

#define BIF_METHOD(name) \
	BIF_DECL(BIF_Obj##name) { \
		if (!BIF_ObjMethod(FID_Obj##name, aResultToken, aParam, aParamCount)) \
			aResult = FAIL; \
	}

ResultType BIF_ObjMethod(int aID, ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	aResultToken.symbol = SYM_STRING;
	aResultToken.marker = _T("");

	Object *obj = dynamic_cast<Object*>(TokenToObject(*aParam[0]));
	if (!obj)
		return OK; // Return "".
	return obj->CallBuiltin(aID, aResultToken, aParam + 1, aParamCount - 1);
}

BIF_METHOD(Insert)
BIF_METHOD(InsertAt)
BIF_METHOD(Push)
BIF_METHOD(Pop)
BIF_METHOD(Delete)
BIF_METHOD(Remove)
BIF_METHOD(RemoveAt)
BIF_METHOD(GetCapacity)
BIF_METHOD(SetCapacity)
BIF_METHOD(GetAddress)
BIF_METHOD(Count)
BIF_METHOD(Length)
BIF_METHOD(MaxIndex)
BIF_METHOD(MinIndex)
BIF_METHOD(NewEnum)
BIF_METHOD(HasKey)
BIF_METHOD(Clone)
BIF_METHOD(Reverse)
BIF_METHOD(Swap)


//
// ObjAddRef/ObjRelease - used with pointers rather than object references.
//

BIF_DECL(BIF_ObjAddRefRelease)
{
	IObject *obj = (IObject *)TokenToInt64(*aParam[0]);
	if (obj < (IObject *)4096) // Rule out some obvious errors.
	{
		aResultToken.symbol = SYM_STRING;
		aResultToken.marker = _T("");
		return;
	}
	if (ctoupper(aResultToken.marker[3]) == 'A')
		aResultToken.value_int64 = obj->AddRef();
	else
		aResultToken.value_int64 = obj->Release();
}


//
// ObjBindMethod(Obj, Method, Params...)
//

BIF_DECL(BIF_ObjBindMethod)
{
	IObject *func, *bound_func;
	if (  !(func = TokenToObject(*aParam[0]))
		&& !(func = TokenToFunc(*aParam[0]))  )
	{
		aResult = g_script.ScriptError(ERR_PARAM1_INVALID);
		return;
	}
	if (  !(bound_func = BoundFunc::Bind(func, aParam + 1, aParamCount - 1, IT_CALL))  )
	{
		aResult = g_script.ScriptError(ERR_OUTOFMEM);
		return;
	}
	aResultToken.symbol = SYM_OBJECT;
	aResultToken.object = bound_func;
}


//
// ObjRawSet - set a value without invoking any meta-functions.
//

BIF_DECL(BIF_ObjRaw)
{
	Object *obj = dynamic_cast<Object*>(TokenToObject(*aParam[0]));
	if (!obj)
	{
		aResult = g_script.ScriptError(ERR_PARAM1_INVALID);
		return;
	}
	if (ctoupper(aResultToken.marker[6]) == 'S')
	{
		if (!obj->SetItem(*aParam[1], *aParam[2]))
		{
			aResult = g_script.ScriptError(ERR_OUTOFMEM);
			return;
		}
	}
	else
	{
		ExprTokenType value;
		if (obj->GetItem(value, *aParam[1]))
		{
			switch (value.symbol)
			{
			case SYM_OPERAND:
				aResultToken.symbol = SYM_STRING;
				aResult = TokenSetResult(aResultToken, value.marker);
				break;
			case SYM_OBJECT:
				aResultToken.symbol = SYM_OBJECT;
				aResultToken.object = value.object;
				aResultToken.object->AddRef();
				break;
			default:
				aResultToken.symbol = value.symbol;
				aResultToken.value_int64 = value.value_int64;
				break;
			}
			return;
		}
	}
	aResultToken.symbol = SYM_STRING;
	aResultToken.marker = _T("");
}


//
// ObjSetBase/ObjGetBase - Change or return Object's base without invoking any meta-functions.
//

BIF_DECL(BIF_ObjBase)
{
	Object *obj = dynamic_cast<Object*>(TokenToObject(*aParam[0]));
	if (!obj)
	{
		aResult = g_script.ScriptError(ERR_PARAM1_INVALID);
		return;
	}
	if (ctoupper(aResultToken.marker[3]) == 'S') // ObjSetBase
	{
		IObject *new_base = TokenToObject(*aParam[1]);
		if (!new_base && !TokenIsEmptyString(*aParam[1]))
		{
			aResult = g_script.ScriptError(ERR_PARAM2_INVALID);
			return;
		}
		obj->SetBase(new_base);
	}
	else // ObjGetBase
	{
		if (IObject *obj_base = obj->Base())
		{
			obj_base->AddRef();
			aResultToken.SetValue(obj_base);
			return;
		}
	}
	aResultToken.SetValue(_T(""));
}
