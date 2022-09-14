
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


#if 0 // Currently unused
// For use reinterpreting member function pointers (not standard C++).
template<typename T> constexpr void* cast_into_voidp(T in)
{
	union { T in; void *out; } u { in };
	return u.out;
}
#endif


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


MdFunc::MdFunc(LPCTSTR aName, void *aMcFunc, MdType aRetType, MdType *aArg, UINT aArgSize)
	: NativeFunc(aName)
	, mMcFunc {aMcFunc}
	, mArgType {aArg}
	, mRetType {aRetType}
	, mMaxResultTokens {0}
	, mArgSlots {0}
	, mThisCall {false}
{
	// #if _DEBUG, ensure aArg is effectively terminated for the inner loop below.
	ASSERT(!aArgSize || !MdType_IsMod(aArg[aArgSize - 1]));

	if (aArgSize > 1 && *aArg == MdType::ThisCall)
	{
		mThisCall = true;
		mArgType++;
	}

	int ac = 0, pc = 0;
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
		if (MdType_Is64bit(aArg[i]))
			++ac;
#endif
		++ac;
		if (!retval && !MdType_IsBits(aArg[i]))
		{
			++pc;
			if (!opt)
				mMinParams = pc;
			if (aArg[i] == MdType::String || aArg[i] == MdType::Variant && out != MdType::Void)
				++mMaxResultTokens;
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

	ResultType result = OK;

	MdType retval_arg_type = MdType::Void;
	int retval_index = -1;
	int output_var_count = 0;
	auto atp = mArgType;
	for (int ai = 0, pi = 0; ai < mArgSlots; ++ai, ++atp)
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

		if (MdType_IsBits(arg_type))
		{
			// arg_type represents a constant value to put directly into args.
			arg_value = MdType_BitsValue(arg_type);
			continue;
		}

		if (ParamIndexIsOmitted(pi))
		{
			if (!opt)
			{
				result = aResultToken.Error(ERR_PARAM_REQUIRED);
				goto end;
			}
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
			if (arg_type == MdType::Float64)
			{
				if (nt.symbol == SYM_INTEGER)
					nt.value_double = (double)nt.value_int64;
			}
			else
			{
				if (nt.symbol == SYM_FLOAT)
					nt.value_int64 = (__int64)nt.value_double;
			}
			void *target = &arg_value;
			if (opt) // Optional values are represented by a pointer to a value.
				arg_value = (UINT_PTR)(target = _alloca(8));
#ifndef _WIN64
			if (MdType_Is64bit(arg_type))
			{
				*(__int64*)target = nt.value_int64;
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
	switch (mRetType)
	{
	case MdType::Int32: aResultToken.SetValue(ri32); break;
	case MdType::Int64: aResultToken.SetValue(ri64); break;
	case MdType::Float64: aResultToken.SetValue(GetDoubleRetval()); break;
	case MdType::String: aResultToken.SetValue((LPTSTR)rup); break; // Strictly statically-allocated strings.
	case MdType::ResultType: aResultToken.SetResult((ResultType)rup); break;
	case MdType::FResult:
		if (FAILED(res))
			FResultToError(aResultToken, aParam, aParamCount, res);
		else if (res == FR_ABORTED)
			retval_index = -1; // Leave aResultToken set to its default value.
		break;
	case MdType::NzIntWin32:
		if (!(BOOL)rup)
			aResultToken.Win32Error();
		break;
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

	// Copy output parameters
	atp = mArgType;
	for (int ai = 0, pi = 0; output_var_count; ++atp, ++ai, ++pi)
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
	case MdType::Int64: aToken.SetValue(*(__int64*)aPtr); break;
	case MdType::Float64: aToken.SetValue(*(double*)aPtr); break;
	case MdType::Object:
		if (aPtr)
			aToken.SetValue(*(IObject**)aPtr);
		else
			aToken.SetValue(_T(""), 0);
		break;
	case MdType::String: // String*, not String
		if (aPtr)
			aToken.SetValue(*(LPTSTR*)aPtr);
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