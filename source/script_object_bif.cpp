#include "stdafx.h" // pre-compiled headers
#include "defines.h"
#include "globaldata.h"
#include "script.h"

#include "script_object.h"


//
// BIF_ObjCreate - Object()
//

void BIF_ObjCreate(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
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

void BIF_ObjArray(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	Object *obj = Object::Create(NULL, 0);
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
	aResultToken.symbol = SYM_STRING;
	aResultToken.marker = _T("");
}
	

//
// BIF_IsObject - IsObject(obj)
//

void BIF_IsObject(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
// IsObject(obj) is currently equivalent to (obj && obj=""), but much more intuitive.
{
	int i;
	for (i = 0; i < aParamCount && TokenToObject(*aParam[i]); ++i);
	aResultToken.value_int64 = (__int64)(i == aParamCount); // TRUE if all are objects.
}
	

//
// BIF_ObjInvoke - Handles ObjGet/Set/Call() and get/set/call syntax.
//

void BIF_ObjInvoke(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
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
        obj->Invoke(aResultToken, *obj_param, invoke_type, aParam, aParamCount);
		if (param_is_var)
			obj->Release();
	}
	// Invoke meta-functions of g_MetaObject.
	else if (INVOKE_NOT_HANDLED == g_MetaObject.Invoke(aResultToken, *obj_param, invoke_type | IF_META, aParam, aParamCount))
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
}
	

//
// BIF_ObjGetInPlace - Handles part of a compound assignment like x.y += z.
//

void BIF_ObjGetInPlace(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
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

void BIF_ObjNew(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	aResultToken.symbol = SYM_STRING;
	aResultToken.marker = _T("");

	ExprTokenType *class_token = aParam[0]; // Save this to be restored later.

	IObject *class_object = TokenToObject(*class_token);
	if (!class_object)
		return;

	Object *new_object = Object::Create(NULL, 0);
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

	// __Init was added so that instance variables can be initialized in the correct order
	// (beginning at the root class and ending at class_object) before __New is called.
	// It shouldn't be explicitly defined by the user, but auto-generated in DefineClassVars().
	name_token.marker = _T("__Init");
	result = class_object->Invoke(aResultToken, this_token, IT_CALL | IF_METAOBJ, aParam, 1);
	if (result != INVOKE_NOT_HANDLED)
	{
		// See similar section below for comments.
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
		if (result == FAIL)
		{
			aParam[0] = class_token; // Restore it to original caller-supplied value.
			return;
		}
	}
	
	// __New may be defined by the script for custom initialization code.
	name_token.marker = Object::sMetaFuncName[4]; // __New
	result = class_object->Invoke(aResultToken, this_token, IT_CALL | IF_METAOBJ, aParam, aParamCount);
	if (result != EARLY_RETURN)
	{
		// Although it isn't likely to happen, if __New points at a built-in function or if mBase
		// (or an ancestor) is not an Object (i.e. it's a ComObject), aResultToken can be set even when
		// the result is not EARLY_RETURN.  So make sure to clean up any result we're not going to use.
		if (aResultToken.symbol == SYM_OBJECT)
			aResultToken.object->Release();
		if (aResultToken.mem_to_free)
		{
			// This can be done by our caller, but is done here for maintainability; i.e. because
			// some callers might expect mem_to_free to be NULL when the result isn't a string.
			free(aResultToken.mem_to_free);
			aResultToken.mem_to_free = NULL;
		}
		if (result == FAIL)
		{
			// Invocation failed, probably due to omitting a required parameter.
			aResultToken.symbol = SYM_STRING;
			aResultToken.marker = _T("");
		}
		else
		{
			// Either it wasn't handled (i.e. neither this class nor any of its super-classes define __New()),
			// or there was no explicit "return", so just return the new object.
			aResultToken.symbol = SYM_OBJECT;
			aResultToken.object = new_object;
		}
	}
	else
		new_object->Release();
	aParam[0] = class_token;
}


//
// BIF_ObjIncDec - Handles pre/post-increment/decrement for object fields, such as ++x[y].
//

void BIF_ObjIncDec(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
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
	BIF_ObjInvoke(temp_result, aParam, aParamCount);

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
	if (temp_result.symbol == SYM_OBJECT)
		temp_result.object->Release();
	if (temp_result.mem_to_free)
		free(temp_result.mem_to_free);

	if (current_value.symbol == PURE_NOT_NUMERIC)
	{
		// Value is non-numeric, so return "".
		aResultToken.symbol = SYM_STRING;
		aResultToken.marker = _T("");
		return;
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
void BIF_Obj##name(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount) \
{ \
	aResultToken.symbol = SYM_STRING; \
	aResultToken.marker = _T(""); \
	\
	Object *obj = dynamic_cast<Object*>(TokenToObject(*aParam[0])); \
	if (obj) \
		obj->_##name(aResultToken, aParam + 1, aParamCount - 1); \
}

BIF_METHOD(Insert)
BIF_METHOD(Remove)
BIF_METHOD(GetCapacity)
BIF_METHOD(SetCapacity)
BIF_METHOD(GetAddress)
BIF_METHOD(MaxIndex)
BIF_METHOD(MinIndex)
BIF_METHOD(NewEnum)
BIF_METHOD(HasKey)
BIF_METHOD(Clone)


//
// ObjAddRef/ObjRelease - used with pointers rather than object references.
//

void BIF_ObjAddRefRelease(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	IObject *obj = (IObject *)TokenToInt64(*aParam[0]);
	if (obj < (IObject *)4096) // Rule out some obvious errors.
	{
		aResultToken.symbol = SYM_STRING;
		aResultToken.marker = _T("");
		return;
	}
	if (aResultToken.marker[3] == 'A')
		aResultToken.value_int64 = obj->AddRef();
	else
		aResultToken.value_int64 = obj->Release();
}