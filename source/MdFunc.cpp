
#include "stdafx.h"
#include "MdFunc.h"
#include "abi.h"
#include "script_func_impl.h"
#include "window.h"


// internal to this file
struct MdFuncEntry
{
	LPCTSTR name;
	void *function;
	MdType rettype;
	MdType argtype[23];
};


#define md_mode decl
#include "lib\functions.h"
#undef md_mode

MdFuncEntry sMdFunc[]
{
	#define md_mode data
	#include "lib\functions.h"
	#undef md_mode
};


Func *Script::GetBuiltInMdFunc(LPTSTR aFuncName)
{
#ifdef _DEBUG
	static bool sChecked = false;
	if (!sChecked)
	{
		sChecked = true;
		for (int i = 1; i < _countof(sMdFunc); ++i)
			if (_tcsicmp(sMdFunc[i-1].name, sMdFunc[i].name) >= 0)
				MsgBox(_T("DEBUG: sMdFunc out of order."), 0, sMdFunc[i].name);
	}
#endif
	int left, right, mid, result;
	for (left = 0, right = _countof(sMdFunc) - 1; left <= right;)
	{
		mid = (left + right) / 2;
		auto &f = sMdFunc[mid];
		result = _tcsicmp(aFuncName, f.name);
		if (result > 0)
			left = mid + 1;
		else if (result < 0)
			right = mid - 1;
		else // Match found.
		{
			int ac;
			for (ac = 0; ac < _countof(f.argtype) && f.argtype[ac] != MdType::Void; ++ac);
			return new MdFunc(f.name, f.function, f.rettype, f.argtype, ac);
		}
	}
	return nullptr;
}


#pragma region PerformDynaCall

extern "C" UINT64 DynaCall(size_t aArgCount, UINT_PTR *aArg, void *aFunction, DWORD aFlag);
extern "C" float GetFloatRetval();
extern "C" double GetDoubleRetval();

#pragma endregion


MdFunc::MdFunc(LPCTSTR aName, void *aMcFunc, MdType aRetType, MdType *aArg, UINT aArgSize, Object *aPrototype)
	: NativeFunc(aName)
	, mMcFunc {aMcFunc}
	, mArgType {aArg}
	, mRetType {aRetType}
	, mMaxResultTokens {0}
	, mArgSlots {0}
	, mPrototype {aPrototype}
	, mThisCall {aPrototype != nullptr}
{
	// #if _DEBUG, ensure aArg is effectively terminated for the inner loop below.
	ASSERT(!aArgSize || !MdType_IsMod(aArg[aArgSize - 1]));

#ifdef ENABLE_MD_THISCALL
	if (aArgSize > 1 && *aArg == MdType::ThisCall)
	{
		mThisCall = true;
		mArgType++;
	}
#endif

	int ac = 0, pc = 0;
	if (aPrototype)
		mMinParams = pc = ac = 1;
	for (UINT i = 0; i < aArgSize; ++i)
	{
		bool opt = false, retval = false;
		MdType out = MdType::Void;
		for (; MdType_IsMod(aArg[i]); ++i)
		{
			ASSERT(i < aArgSize);
			if (aArg[i] == MdType::Optional)
				opt = true;
			else if (MdType_IsOut(aArg[i]))
				out = aArg[i];
			else if (aArg[i] == MdType::RetVal)
				retval = true;
		}
#ifndef _WIN64
		if (MdType_Is64bit(aArg[i]) && out == MdType::Void && !opt) // out and opt parameters are excluded because they are passed by address.
			++ac;
#endif
		++ac;
		if (aArg[i] == MdType::Params)
		{
			ASSERT(out == MdType::Void && !retval && !opt);
			mIsVariadic = true;
		}
		else if (!retval && !MdType_IsBits(aArg[i]))
		{
			ASSERT(!mIsVariadic);
			++pc;
			if (!opt && pc - 1 == mMinParams)
				mMinParams = pc;
			if (aArg[i] == MdType::Variant && out != MdType::Void)
				++mMaxResultTokens;
#ifdef ENABLE_IMPLICIT_TOSTRING
			else if (aArg[i] == MdType::String)
				++mMaxResultTokens;
#endif
		}
	}
	mParamCount = pc;
	mArgSlots = ac;
}


bool MdFunc::Call(ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	if (!Func::Call(aResultToken, aParam, aParamCount))
		return false;

	if (aParamCount < mMinParams)
	{
		aResultToken.Error(ERR_TOO_FEW_PARAMS, mName);
		return false;
	}

	DEBUGGER_STACK_PUSH(this) // See comments in BuiltInFunc::Call.

	// rtp stores the results of ToString() calls if needed.
	ResultToken *rtp = mMaxResultTokens == 0 ? nullptr
		: (ResultToken *)_alloca(mMaxResultTokens * sizeof(ResultToken));
	int rt_count = 0;
	
	UINT_PTR *args = (UINT_PTR *)_alloca(mArgSlots * sizeof(UINT_PTR));
	int first_param_index = 0;

	ResultType result = OK;

	if (mPrototype) // This implies thiscall and an initial 'this' parameter, excluded from mArgType.
	{
		auto obj = ParamIndexToObject(0);
		// This handling here is similar to that in BuiltInMethod::Call():
		if (!obj || !obj->IsOfType(mPrototype))
		{
			LPCTSTR expected_type;
			ExprTokenType value;
			if (mPrototype->GetOwnProp(value, _T("__Class")) && value.symbol == SYM_STRING)
				expected_type = value.marker;
			else
				expected_type = _T("?"); // Script may have tampered with the prototype.
			result = aResultToken.TypeError(expected_type, *aParam[0]);
			goto end;
		}
		// This is reliant on (IObject*)obj being the same address as (WhateverObject*)obj
		// (which might not be the case for classes with multiple virtual base classes):
		args[0] = (UINT_PTR)obj;
		first_param_index = 1;
	}

	MdType retval_arg_type = MdType::Void;
	int retval_index = -1;
	int output_var_count = 0;
	auto atp = mArgType;
	for (int ai = first_param_index, pi = ai; ai < mArgSlots; ++ai, ++atp)
	{
		bool opt = false;
		MdType out = MdType::Void;
		for (; MdType_IsMod(*atp); ++atp)
		{
			if (*atp == MdType::Optional)
				opt = true;
			else if (MdType_IsOut(*atp))
				out = *atp;
			else if (*atp == MdType::RetVal)
				retval_index = ai;
		}
		ASSERT(retval_index != ai || out != MdType::Void && !opt);
		auto arg_type = *atp;
		auto &arg_value = args[ai];
		
		if (arg_type == MdType::Params)
		{
			auto p = (VariantParams *)_alloca(sizeof(VariantParams));
			p->count = aParamCount - pi;
			p->value = aParam + pi;
			arg_value = (UINT_PTR)p;
			continue; // Not break, since there might be a retval parameter after it.
		}
		
		if (out != MdType::Void && !(opt && ParamIndexIsOmitted(pi)))
		{
			void *av_buf;
			// Different 'out' modifiers could be supported here to control allocation behaviour,
			// but for built-in functions we'll just use StrRet for String, pointer for all others.
			if (arg_type == MdType::String)
			{
				if (retval_index == ai)
					av_buf = aResultToken.buf;
				else
					av_buf = _alloca(_TSIZE(StrRet::CallerBufSize));
				*(LPTSTR)av_buf = '\0';
 				av_buf = new (_alloca(sizeof(StrRet))) StrRet((LPTSTR)av_buf);
				//MessageBox(NULL, (LPCWSTR)av_buf, NULL, 0);
			}
			else if (arg_type == MdType::Variant)
			{
				// Rarely used, so no specialized abstraction yet.
				if (retval_index == ai)
					av_buf = &aResultToken;
				else
				{
					ResultToken &rt = rtp[rt_count++];
					rt.InitResult(talloca(_f_retval_buf_size));
					av_buf = &rt;
				}
			}
			else
			{
				//ASSERT(arg_type != MdType::Variant);
				av_buf = _alloca(8); // This buffer will receive the actual output value.
				*(__int64*)av_buf = 0;
			}
			arg_value = (UINT_PTR)av_buf; // Pass the address of the buffer or StrRet.
		}
		
		if (retval_index == ai) // Not included within aParam.
		{
			//if (arg_type == MdType::Variant)
			//	arg_value = (UINT_PTR)&aResultToken;
			retval_arg_type = arg_type;
			//retval_out_type = out;
			continue;
		}

#ifdef ENABLE_MD_BITS
		if (MdType_IsBits(arg_type))
		{
			// arg_type represents a constant value to put directly into args.
			arg_value = MdType_BitsValue(arg_type);
			continue;
		}
#endif

		if (ParamIndexIsOmitted(pi))
		{
			if (!opt)
			{
				result = aResultToken.ParamError(pi, nullptr);
				goto end;
			}
			// Pass nullptr for this optional parameter to indicate that it has been omitted.
			// MdType_Is64bit(arg_type) isn't relevant in this case since opt == true.
			arg_value = 0;
			pi++;
			continue;
		}
		auto &param = *aParam[pi++];

		if (out != MdType::Void) // Out or some variant, and not retval (which was already handled).
		{
			if (!TokenToOutputVar(param))
			{
				result = aResultToken.ParamError(pi, &param, _T("VarRef"));
				goto end;
			}
			++output_var_count;
			// arg_value was already set above.
			continue;
		}

		if (arg_type == MdType::String)
		{
			LPTSTR buf = nullptr;
			switch (param.symbol)
			{
			case SYM_VAR:
				if (!param.var->HasObject())
					break; // No buffer needed even if it's pure numeric.
			case SYM_INTEGER:
			case SYM_FLOAT:
			case SYM_OBJECT:
				buf = (LPTSTR)_alloca(MAX_NUMBER_SIZE * sizeof(TCHAR));
			}
			ExprTokenType *t;
			if (auto obj = TokenToObject(param))
			{
#ifdef ENABLE_IMPLICIT_TOSTRING
				ResultToken &rt = rtp[rt_count++];
				rt.InitResult(buf);
				ObjectToString(rt, param, obj);
				if (rt.Exited())
				{
					result = FAIL;
					goto end;
				}
				if (rt.symbol == SYM_OBJECT)
				{
					result = aResultToken.TypeError(_T("String"), rt);
					goto end;
				}
				t = &rt;
#else
				result = aResultToken.ParamError(pi, &param, _T("String"));
				goto end;
#endif
			}
			else
				t = &param;
			arg_value = (UINT_PTR)TokenToString(*t, buf);
		}
		else if (arg_type == MdType::Object)
		{
			arg_value = (DWORD_PTR)TokenToObject(param);
			if (!arg_value)
			{
				result = aResultToken.ParamError(pi, &param, _T("Object"));
				goto end;
			}
		}
		else if (arg_type == MdType::Variant)
		{
			arg_value = (DWORD_PTR)&param;
		}
		else
		{
			ExprTokenType nt;
			if (arg_type == MdType::Bool32)
			{
				nt.SetValue(TokenToBOOL(param));
			}
			else
			{
				ASSERT(MdType_IsNum(arg_type));
				if (!TokenToDoubleOrInt64(param, nt))
				{
					result = aResultToken.ParamError(pi - 1, &param, _T("Number"));
					goto end;
				}
			}
			// If necessary, convert integer <-> float within the value union.
			switch (arg_type)
			{
			case MdType::Float64:
				if (nt.symbol == SYM_INTEGER)
					nt.value_double = (double)nt.value_int64;
				break;
			case MdType::Float32:
				if (nt.symbol == SYM_INTEGER)
					*((float*)&nt.value_int64) = (float)nt.value_int64;
				else
					*((float*)&nt.value_int64) = (float)nt.value_double;
				break;
			default:
				if (nt.symbol == SYM_FLOAT)
					nt.value_int64 = (__int64)nt.value_double;
				break;
			}
			void *target = &arg_value;
			if (opt) // Optional values are represented by a pointer to a value.
				arg_value = (UINT_PTR)(target = _alloca(8));
#ifndef _WIN64
			if (MdType_Is64bit(arg_type))
			{
				*(__int64*)target = nt.value_int64;
				if (!opt) // See above.
					++ai; // Consume an additional arg slot.
			}
			else
				*(UINT*)target = (UINT)nt.value_int64;
#else
			*(__int64*)target = nt.value_int64;
#endif
		}
	}

	union {
		UINT64 rup;
		__int64 ri64;
		int ri32;
		FResult res;
	};
	// Make the call
	rup = DynaCall(mArgSlots, args, mMcFunc, mThisCall);

	// Convert the return value
	bool aborted = false;
	switch (mRetType)
	{
	// Unused return types are commented out or omitted to reduce code size, and disabled
	// at compile-time by not providing a valid md_retval<T>::t via template definition.
	// Place the most common type (FResult) first in case this compiles to an if-else ladder:
	case MdType::FResult:
		if (FAILED(res))
		{
			FResultToError(aResultToken, aParam, aParamCount, res, first_param_index);
			aborted = true;
		}
		else
			aborted = (res == FR_ABORTED);
		break;
	case MdType::ResultType: aResultToken.SetResult((ResultType)rup); break;
	case MdType::Int32: aResultToken.SetValue(ri32); break;
	case MdType::UInt64:
	case MdType::Int64: aResultToken.SetValue(ri64); break;
	case MdType::UInt32: aResultToken.SetValue((UINT)rup); break;
	//case MdType::Float64: aResultToken.SetValue(GetDoubleRetval()); break;
	//case MdType::String: aResultToken.SetValue((LPTSTR)rup); break; // Strictly statically-allocated strings.
	//case MdType::NzIntWin32:
	//	if (!(BOOL)rup)
	//		aResultToken.Win32Error();
	//	break;
	case MdType::Bool32: aResultToken.SetValue(ri32 ? TRUE : FALSE); break;
	}
	if (retval_index != -1)
	{
		if (retval_arg_type == MdType::String)
		{
			auto strret = (StrRet*)args[retval_index];
			if (strret->UsedMalloc())
				aResultToken.AcceptMem(const_cast<LPTSTR>(strret->Value()), strret->Length());
			else if (strret->Value())
				aResultToken.SetValue(const_cast<LPTSTR>(strret->Value()), strret->Length());
			//else leave aResultToken set to its default value, "".
		}
		else if (retval_arg_type != MdType::Variant) // Variant type passes aResultToken directly.
			TypedPtrToToken(retval_arg_type, (void*)args[retval_index], aResultToken);
	}
	if (aborted)
	{
		aResultToken.Free(); // In case memory was allocated or an object was returned, despite the return value indicating failure.
		aResultToken.mem_to_free = nullptr; // Because Free() doesn't clear it.
		aResultToken.SetValue(_T(""), 0);
	}

	// Copy output parameters
	atp = mArgType;
	for (int ai = first_param_index, pi = ai; output_var_count; ++atp, ++ai, ++pi)
	{
		if (ai == retval_index)
			++ai; // This args slot doesn't correspond to an aParam slot.
		ASSERT(pi < aParamCount); // Implied by how output_var_count was calculated.
		MdType out = MdType::Void;
		for (; MdType_IsMod(*atp); ++atp)
			if (MdType_IsOut(*atp))
				out = *atp;
		if (out == MdType::Void || ParamIndexIsOmitted(pi))
			continue;
		--output_var_count;
		auto var = ParamIndexToOutputVar(pi);
		ASSERT(var); // Implied by validation during processing of parameter inputs.
		auto arg_value = args[ai];
		if (*atp == MdType::String)
		{
			auto strret = (StrRet*)arg_value;
			if (!strret->Value())
				var->Assign();
			else if (strret->UsedMalloc())
				var->AcceptNewMem(const_cast<LPTSTR>(strret->Value()), strret->Length());
			else
				var->AssignString(strret->Value(), strret->Length());
		}
		else if (*atp == MdType::Variant)
		{
			ResultToken &value = *(ResultToken*)arg_value;
			if (value.mem_to_free)
			{
				ASSERT(value.symbol == SYM_STRING && value.marker == value.mem_to_free);
				var->AcceptNewMem(value.marker, value.marker_length);
				value.mem_to_free = nullptr;
			}
			else
				var->Assign(value);
			// ResultTokens are allocated from rtp[], and are freed below.
			//value.Free();
		}
		else
		{
			ExprTokenType value;
			TypedPtrToToken(*atp, (void*)arg_value, value);
			if (value.symbol == SYM_OBJECT)
				var->AssignSkipAddRef(value.object);
			else
				var->Assign(value);
		}
		// Now that any memory or object allocated by the function has been assigned:
		if (aborted)
		{
			// Although 0 or "" is a fairly conventional default, it might not be safe.
			// For error-detection and to avoid unexpected behaviour, "unset" the var.
			var->Uninitialize();
		}
	}

end:
	DEBUGGER_STACK_POP()
	// Free any temporary results of ToString() calls.
	for (int i = 0; i < rt_count; ++i)
		rtp[i].Free();
	return result;
}


// Shallow-copy a value from aPtr to aToken.
void TypedPtrToToken(MdType aType, void *aPtr, ExprTokenType &aToken)
{
	switch (aType)
	{
	case MdType::Bool32:
	case MdType::Int32: aToken.SetValue(*(int*)aPtr); break;
	case MdType::UInt32: aToken.SetValue(*(UINT*)aPtr); break;
	case MdType::UInt64:
	case MdType::Int64: aToken.SetValue(*(__int64*)aPtr); break;
	case MdType::Float64: aToken.SetValue(*(double*)aPtr); break;
	case MdType::Float32: aToken.SetValue(*(float*)aPtr); break;
	case MdType::Int8: aToken.SetValue(*(INT8*)aPtr); break;
	case MdType::UInt8: aToken.SetValue(*(UINT8*)aPtr); break;
	case MdType::Int16: aToken.SetValue(*(INT16*)aPtr); break;
	case MdType::UInt16: aToken.SetValue(*(UINT16*)aPtr); break;
	case MdType::Object:
		if (auto obj = *(IObject**)aPtr)
			aToken.SetValue(obj);
		else
			aToken.SetValue(_T(""), 0);
		break;
	}
}


bool MdFunc::ArgIsOutputVar(int aIndex)
{
	auto atp = mArgType;
	for (int ai = 0; ai < mArgSlots; ++ai, ++atp, --aIndex)
	{
		MdType out = MdType::Void;
		for (; MdType_IsMod(*atp); ++atp)
		{
			if (MdType_IsOut(*atp))
				out = *atp;
			else if (*atp == MdType::RetVal)
				++aIndex;
		}
		if (aIndex == 0)
			return out != MdType::Void;
#ifndef _WIN64
		if (MdType_Is64bit(*atp))
			++ai;
#endif
	}
	return false;
}


bool MdFunc::ArgIsOptional(int aIndex)
{
	auto atp = mArgType;
	for (int ai = 0; ai < mArgSlots; ++ai, ++atp, --aIndex)
	{
		bool opt = false;
		for (; MdType_IsMod(*atp); ++atp)
		{
			if (*atp == MdType::Optional)
				opt = true;
			else if (*atp == MdType::RetVal)
				++aIndex;
		}
		if (aIndex == 0)
			return opt;
#ifndef _WIN64
		if (MdType_Is64bit(*atp))
			++ai;
#endif
	}
	return false;
}


Object *Object::DefineMetadataMembers(Object *obj, LPCTSTR aClassName, ObjectMemberMd aMember[], int aMemberCount)
{
	if (aMemberCount)
		obj->mFlags |= NativeClassPrototype;

	TCHAR full_name[MAX_VAR_NAME_LENGTH + 1];
	TCHAR *name = full_name + _stprintf(full_name, _T("%s.Prototype."), aClassName);

	for (int i = 0; i < aMemberCount; ++i)
	{
		auto &member = aMember[i];

		int ac;
		for (ac = 0; ac < _countof(member.argtype) && member.argtype[ac] != MdType::Void; ++ac);

		_tcscpy(name, member.name);
		auto op_name = _tcschr(name, '\0');
		switch (member.invokeType)
		{
		case IT_GET: _tcscpy(op_name, _T(".Get")); break;
		case IT_SET: _tcscpy(op_name, _T(".Set")); break;
		}

		auto func = new MdFunc(SimpleHeap::Alloc(full_name), member.method, MdType::FResult, member.argtype, ac, obj);

		if (member.invokeType == IT_CALL)
		{
			obj->DefineMethod(name, func);
		}
		else
		{
			auto prop = obj->DefineProperty(const_cast<LPTSTR>(member.name));
			if (member.invokeType == IT_GET)
			{
				prop->MinParams = func->mMinParams - 1;
				if (!func->mIsVariadic)
					prop->MaxParams = func->mParamCount - 1;
				prop->SetGetter(func);
			}
			else
			{
				// There should be a getter for every setter; rely on the getter to set Min/MaxParams:
				//prop->MinParams = func->mMinParams - 2;
				//prop->MaxParams = func->mParamCount - 2;
				prop->SetSetter(func);
			}
		}
		func->Release();
	}
	return obj;
}
