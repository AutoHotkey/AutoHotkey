#include "stdafx.h" // pre-compiled headers
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
		if (obj < (IObject *)65536) // Prevent some obvious errors.
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
		aResultToken.Error(aParamCount == 1 ? ERR_PARAM1_INVALID : ERR_OUTOFMEM);
	}
}


//
// BIF_ObjArray - Array(items*)
//

BIF_DECL(BIF_ObjArray)
{
	Object *obj = Object::Create();
	if (obj)
	{
		if (!aParamCount || obj->InsertAt(0, 1, aParam, aParamCount))
		{
			aResultToken.symbol = SYM_OBJECT;
			aResultToken.object = obj;
			return;
		}
		obj->Release();
	}
	aResultToken.Error(ERR_OUTOFMEM);
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
    
	ResultType result;
    if (obj)
	{
		bool param_is_var = obj_param->symbol == SYM_VAR;
		if (param_is_var)
			// Since the variable may be cleared as a side-effect of the invocation, call AddRef to ensure the object does not expire prematurely.
			// This is not necessary for SYM_OBJECT since that reference is already counted and cannot be released before we return.  Each object
			// could take care not to delete itself prematurely, but it seems more proper, more reliable and more maintainable to handle it here.
			obj->AddRef();
        result = obj->Invoke(aResultToken, *obj_param, invoke_type, aParam, aParamCount);
		if (param_is_var)
			obj->Release();
	}
	// Invoke meta-functions of g_MetaObject.
	else if (INVOKE_NOT_HANDLED == (result = g_MetaObject.Invoke(aResultToken, *obj_param, invoke_type | IF_META, aParam, aParamCount)))
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
				result = g_MetaObject.Invoke(aResultToken, base_token, invoke_type, aParam + 1, aParamCount - 1);
			}
			else					// "".base
			{
				// Return a reference to g_MetaObject.  No need to AddRef as g_MetaObject ignores it.
				aResultToken.symbol = SYM_OBJECT;
				aResultToken.object = &g_MetaObject;
				result = OK;
			}
		}
		else
		{
			// Since it wasn't handled (not even by g_MetaObject), maybe warn at this point:
			if (obj_param->symbol == SYM_VAR)
				obj_param->var->MaybeWarnUninitialized();
		}
	}
	if (result == INVOKE_NOT_HANDLED)
	{
		// Invocation not handled. Either there was no target object, or the object doesn't handle
		// this method/property.  For Object (associative arrays), only CALL should give this result.
		if (!obj)
			aResultToken.Error(ERR_NO_OBJECT);
		else
			aResultToken.Error(ERR_NO_MEMBER, aParamCount ? TokenToString(*aParam[0]) : _T(""));
	}
	else if (result == FAIL || result == EARLY_EXIT) // For maintainability: SetExitResult() might not have been called.
	{
		aResultToken.SetExitResult(result);
	}
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
	BIF_ObjInvoke(aResultToken, aParam - aParamCount, aParamCount);
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
	{
		aResultToken.Error(_T("Missing class object for \"new\" operator."));
		return;
	}

	Object *new_object = Object::Create();
	if (!new_object)
	{
		aResultToken.Error(ERR_OUTOFMEM);
		return;
	}

	new_object->SetBase(class_object);

	ExprTokenType name_token, this_token;
	this_token.symbol = SYM_OBJECT;
	this_token.object = new_object;
	name_token.symbol = SYM_STRING;
	aParam[0] = &name_token;

	ResultType result;
	LPTSTR buf = aResultToken.buf; // In case Invoke overwrites it via the union.

	// __Init was added so that instance variables can be initialized in the correct order
	// (beginning at the root class and ending at class_object) before __New is called.
	// It shouldn't be explicitly defined by the user, but auto-generated in DefineClassVars().
	name_token.marker = _T("__Init");
	result = class_object->Invoke(aResultToken, this_token, IT_CALL | IF_METAOBJ, aParam, 1);
	if (result != INVOKE_NOT_HANDLED)
	{
		if (result != OK)
		{
			new_object->Release();
			aParam[0] = class_token; // Restore it to original caller-supplied value.
			aResultToken.SetExitResult(result);
			return;
		}
		// See similar section below for comments.
		aResultToken.Free();
		// Reset to defaults for __New, invoked below.
		aResultToken.mem_to_free = NULL;
		aResultToken.symbol = SYM_STRING;
		aResultToken.marker = _T("");
		aResultToken.buf = buf;
	}
	
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
		// An error was raised within __New() or while trying to call it.
		new_object->Release();
		aResultToken.SetExitResult(result);
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

	ResultToken temp_result;
	// Set the defaults expected by BIF_ObjInvoke:
	temp_result.symbol = SYM_INTEGER;
	temp_result.marker = (LPTSTR)IT_GET;
	temp_result.buf = aResultToken.buf;
	temp_result.mem_to_free = NULL;

	// Retrieve the current value.  Do it this way instead of calling Object::Invoke
	// so that if aParam[0] is not an object, g_MetaObject is correctly invoked.
	BIF_ObjInvoke(temp_result, aParam, aParamCount);

	if (temp_result.Exited())
		return;

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
	}

	// Free the object or string returned by BIF_ObjInvoke, if applicable.
	temp_result.Free();

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
		BIF_ObjInvoke(aResultToken, param, aParamCount);
	}
	else // SYM_POST_INCREMENT || SYM_POST_DECREMENT
	{
		// Must be re-initialized (and must use IT_SET instead of IT_GET):
		temp_result.symbol = SYM_INTEGER;
		temp_result.marker = (LPTSTR)IT_SET;
		temp_result.buf = aResultToken.buf;
		temp_result.mem_to_free = NULL;
		
		// Set the new value.
		BIF_ObjInvoke(temp_result, param, aParamCount);
		
		// Dispose of the result safely.
		temp_result.Free();

		// Return the previous value.
		aResultToken.symbol = current_value.symbol;
		aResultToken.value_int64 = current_value.value_int64; // Union copy.
	}
}


//
// Functions for accessing built-in methods (even if obscured by a user-defined method).
//

BIF_DECL(BIF_ObjXXX)
{
	LPTSTR method_name = aResultToken.marker + 3;

	aResultToken.symbol = SYM_STRING;
	aResultToken.marker = _T("");
	
	Object *obj = dynamic_cast<Object*>(TokenToObject(*aParam[0]));
	if (obj)
	{
		if (!obj->CallBuiltin(method_name, aResultToken, aParam + 1, aParamCount - 1))
			aResultToken.SetExitResult(FAIL);
	}
	else
		aResultToken.Error(ERR_NO_OBJECT);
}

BIF_DECL(BIF_ObjNewEnum)
{
	aResultToken.symbol = SYM_STRING;
	aResultToken.marker = _T("");
	
	Object *obj = dynamic_cast<Object*>(TokenToObject(*aParam[0]));
	if (obj)
	{
		if (!obj->_NewEnum(aResultToken, NULL, 0)) // Parameters are ignored.
			aResultToken.SetExitResult(FAIL);
	}
	else
		aResultToken.Error(ERR_NO_OBJECT);
}


//
// ObjAddRef/ObjRelease - used with pointers rather than object references.
//

BIF_DECL(BIF_ObjAddRefRelease)
{
	IObject *obj = (IObject *)TokenToInt64(*aParam[0]);
	if (obj < (IObject *)65536) // Rule out some obvious errors.
	{
		aResultToken.Error(ERR_PARAM1_INVALID);
		return;
	}
	if (ctoupper(aResultToken.marker[3]) == 'A')
		aResultToken.value_int64 = obj->AddRef();
	else
		aResultToken.value_int64 = obj->Release();
}


//
// ObjRawSet - set a value without invoking any meta-functions.
//

BIF_DECL(BIF_ObjRawSet)
{
	Object *obj = dynamic_cast<Object*>(TokenToObject(*aParam[0]));
	if (!obj)
	{
		aResultToken.Error(ERR_PARAM1_INVALID);
		return;
	}
	if (!obj->SetItem(*aParam[1], *aParam[2]))
		aResultToken.Error(ERR_OUTOFMEM);
}