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
	if (  !(func = ParamIndexToObject(0))  )
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
