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
	name_token.symbol = SYM_STRING;
	name_token.marker = Object::sMetaFuncName[4]; // __New
	this_token.symbol = SYM_OBJECT;
	this_token.object = new_object;
	aParam[0] = &name_token;
	if (class_object->Invoke(aResultToken, this_token, IT_CALL | IF_META, aParam, aParamCount) == INVOKE_NOT_HANDLED)
	{
		// Since it wasn't handled, neither this class nor any of its super-classes define __New().
		aResultToken.symbol = SYM_OBJECT;
		aResultToken.object = new_object;
	}
	aParam[0] = class_token;
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