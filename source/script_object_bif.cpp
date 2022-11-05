#include "stdafx.h" // pre-compiled headers
#include "defines.h"
#include "globaldata.h"
#include "script.h"

#include "script_object.h"
#include "script_func_impl.h"


extern BuiltInFunc *OpFunc_GetProp, *OpFunc_GetItem, *OpFunc_SetProp, *OpFunc_SetItem;

//
// Object()
//

BIF_DECL(Op_Object)
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

BIF_DECL(Op_Array)
{
	if (auto arr = Array::Create(aParam, aParamCount))
		_f_return(arr);
	_f_throw_oom;
}
	

//
// IsObject
//

BIF_DECL(BIF_IsObject)
{
	_f_return_b(TokenToObject(*aParam[0]) != nullptr);
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
		_f_throw_param(0);
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
	if (  !(func = ParamIndexToObject(0))  )
		_f_throw_param(0);
	LPCTSTR name = nullptr;
	if (aParamCount > 1)
	{
		switch (TypeOfToken(*aParam[1]))
		{
		case SYM_MISSING: break;
		case SYM_OBJECT: _f_throw_param(1, _T("String"));
		default: name = TokenToString(*aParam[1], _f_number_buf);
		}
		aParam += 2;
		aParamCount -= 2;
	}
	else
		aParamCount = 0;
	bound_func = BoundFunc::Bind(func, IT_CALL, name, aParam, aParamCount);
	if (!bound_func)
		_f_throw_oom;
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
			_f_throw_param(0);
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
	IObject *iobj = ParamIndexToObject(0);
	if (_f_callee_id == FID_ObjSetBase)
	{
		Object *obj = dynamic_cast<Object*>(iobj);
		if (!obj)
			_f_throw_type(_T("Object"), *aParam[0]);
		auto new_base = dynamic_cast<Object *>(TokenToObject(*aParam[1]));
		if (!new_base)
			_f_throw_type(_T("Object"), *aParam[1]);
		if (!obj->SetBase(new_base, aResultToken))
			return;
	}
	else // ObjGetBase
	{
		Object *obj_base;
		if (iobj)
			obj_base = iobj->Base();
		else
			obj_base = Object::ValueBase(*aParam[0]);
		if (obj_base)
		{
			obj_base->AddRef();
			_f_return(obj_base);
		}
		// Otherwise, could be Object::sAnyPrototype or SYM_MISSING (via a variadic call).
	}
	_f_return_empty;
}


bool Object::HasBase(ExprTokenType &aValue, IObject *aBase)
{
	Object *value_base;
	if (auto obj = TokenToObject(aValue))
		value_base = obj->Base();
	else
		value_base = Object::ValueBase(aValue);
	if (value_base)
		return value_base == aBase || value_base->IsDerivedFrom(aBase);
	// Otherwise, could be Object::sAnyPrototype or SYM_MISSING (via a variadic call).
	return false;
}


BIF_DECL(BIF_HasBase)
{
	auto that_base = ParamIndexToObject(1);
	if (!that_base)
		_f_throw_type(_T("object"), *aParam[1]);
	_f_return_b(Object::HasBase(*aParam[0], that_base));
}


Object *ParamToObjectOrBase(ExprTokenType &aToken)
{
	if (IObject *iobj = TokenToObject(aToken))
	{
		if (iobj->IsOfType(Object::sPrototype))
			return (Object *)iobj;
		return iobj->Base();
	}
	return Object::ValueBase(aToken);
}


BIF_DECL(BIF_HasProp)
{
	auto obj = ParamToObjectOrBase(*aParam[0]);
	if (obj == Object::sComObjectPrototype)
		_f_throw_param(0);
	_f_return_b(obj->HasProp(ParamIndexToString(1, _f_number_buf)));
}


BIF_DECL(BIF_GetMethod)
{
	auto obj = ParamToObjectOrBase(*aParam[0]);
	if (obj == Object::sComObjectPrototype)
		_f_throw_param(0);
	auto method_name = ParamIndexToOptionalStringDef(1, nullptr, _f_number_buf);
	auto method = method_name ? obj->GetMethod(method_name) : obj; // Validate obj itself as a function if method name is omitted.
	if (method)
	{
		int param_count = -1; // Default to no parameter count validation.
		if (!ParamIndexIsOmitted(2))
		{
			Throw_if_Param_NaN(2);
			param_count = ParamIndexToInt(2);
		}
		if (param_count != -1 && method_name)
			++param_count; // So caller does not need to include the implicit `this` parameter.
		switch (ValidateFunctor(method, param_count, aResultToken, nullptr, _f_callee_id == FID_GetMethod))
		{
		case FAIL: return; // A property call threw an exception, or validation failed for GetMethod.
		case CONDITION_FALSE: method = nullptr; // Validation failed for HasMethod.
		}
	}
	if (_f_callee_id == FID_HasMethod)
		_f_return_b(method != nullptr);
	if (!method) // No method for GetMethod to return: throw MethodError().
		_f__ret(aResultToken.UnknownMemberError(*aParam[0], IT_CALL, method_name));
	method->AddRef();
	_f_return(method);
}
