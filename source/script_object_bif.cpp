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
	++aParam;
	--aParamCount;
    
    if (obj = TokenToObject(*obj_param))
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
	else if (INVOKE_NOT_HANDLED == g_MetaObject.Invoke(aResultToken, *obj_param, invoke_type | IF_META, aParam, aParamCount)
		// If above did not handle it, check for attempts to access .base of non-object value (g_MetaObject itself).
		// CALL is unsupported (for simplicity); SET is supported only in "multi-dimensional" mode: "".base[x]:=y
		&& invoke_type != IT_CALL && (invoke_type == IT_SET ? aParamCount > 2 : aParamCount)
		&& !_tcsicmp(TokenToString(*aParam[0]), _T("base")))
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
}