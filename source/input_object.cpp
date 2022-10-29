#include "stdafx.h" // pre-compiled headers
#include "defines.h"
#include "globaldata.h"
#include "script.h"
#include "application.h"

#include "script_object.h"
#include "script_func_impl.h"
#include "input_object.h"


ObjectMemberMd InputObject::sMembers[] =
{
	#define BOOL_OPTION(OPT)  md_property(InputObject, OPT, Bool32)
	#define ONX_OPTION(X) \
		md_property_get(InputObject, On##X, Object), \
		md_property_set(InputObject, On##X, Variant)
	
	md_member(InputObject, __New, CALL, (In_Opt, String, Options), (In_Opt, String, EndKeys), (In_Opt, String, MatchList)),
	md_member(InputObject, KeyOpt, CALL, (In, String, Keys), (In, String, KeyOptions)),
	md_member(InputObject, Start, CALL, md_arg_none),
	md_member(InputObject, Stop, CALL, md_arg_none),
	md_member(InputObject, Wait, CALL, (In_Opt, Float64, MaxTime), (Ret, String, RetVal)),
	
	BOOL_OPTION(BackspaceIsUndo),
	BOOL_OPTION(CaseSensitive),
	BOOL_OPTION(FindAnywhere),
	BOOL_OPTION(NotifyNonText),
	BOOL_OPTION(VisibleNonText),
	BOOL_OPTION(VisibleText),
	
	md_property_get(InputObject, EndKey, String),
	md_property_get(InputObject, EndMods, String),
	md_property_get(InputObject, EndReason, String),
	md_property_get(InputObject, InProgress, Bool32),
	md_property_get(InputObject, Input, String),
	md_property_get(InputObject, Match, String),
	
	md_property(InputObject, MinSendLevel, Int32),
	
	ONX_OPTION(Char),
	ONX_OPTION(End),
	ONX_OPTION(KeyDown),
	ONX_OPTION(KeyUp),
	
	md_property(InputObject, Timeout, Float64)
	#undef BOOL_OPTION
	#undef ONX_OPTION
};
int InputObject::sMemberCount = _countof(sMembers);

Object *InputObject::sPrototype;

InputObject::InputObject()
{
	input.ScriptObject = this;
	SetBase(sPrototype);
}


Object *InputObject::Create()
{
	return new InputObject();
}


FResult InputObject::__New(optl<StrArg> aOptions, optl<StrArg> aEndKeys, optl<StrArg> aMatchList)
{
	return input.Setup(aOptions.value_or_empty(), aEndKeys.value_or_empty(), aMatchList.value_or_empty()) ? OK : FR_FAIL;
}


FResult InputObject::Start()
{
	if (!input.InProgress())
	{
		input.Buffer[input.BufferLength = 0] = '\0';
		InputStart(input);
	}
	return OK;
}


FResult InputObject::Wait(optl<double> aMaxTime, StrRet &aRetVal)
{
	auto wait_ms = aMaxTime.has_value() ? (UINT)(aMaxTime.value() * 1000) : UINT_MAX;
	auto tick_start = GetTickCount();
	while (input.InProgress() && (GetTickCount() - tick_start) < wait_ms)
		MsgSleep();
	return get_EndReason(aRetVal);
}


FResult InputObject::Stop()
{
	if (input.InProgress())
		input.Stop();
	return OK;
}


FResult InputObject::get_EndReason(StrRet &aRetVal)
{
	aRetVal.SetStatic(input.GetEndReason(NULL, 0));
	return OK;
}

FResult InputObject::get_EndKey(StrRet &aRetVal)
{
	if (input.Status == INPUT_TERMINATED_BY_ENDKEY)
	{
		auto buf = aRetVal.CallerBuf();
		input.GetEndReason(buf, aRetVal.CallerBufSize);
		aRetVal.SetTemp(buf);
	}
	else
		aRetVal.SetEmpty();
	return OK;
}

FResult InputObject::get_EndMods(StrRet &aRetVal)
{
	auto buf = aRetVal.CallerBuf(), cp = buf;
	const auto mod_string = MODLR_STRING;
	for (int i = 0; i < 8; ++i)
		if (input.EndingMods & (1 << i))
		{
			*cp++ = mod_string[i * 2];
			*cp++ = mod_string[i * 2 + 1];
		}
	*cp = '\0';
	aRetVal.SetTemp(buf, cp - buf);
	return OK;
}

FResult InputObject::get_Input(StrRet &aRetVal)
{
	aRetVal.SetTemp(input.Buffer, input.BufferLength);
	return OK;
}

FResult InputObject::get_Match(StrRet &aRetVal)
{
	if (input.Status == INPUT_TERMINATED_BY_MATCH && input.EndingMatchIndex < input.MatchCount)
		aRetVal.SetTemp(input.match[input.EndingMatchIndex]);
	else
		aRetVal.SetEmpty();
	return OK;
}


FResult InputObject::get_On(IObject *&aRetVal, IObject *&aOn)
{
	if (aRetVal = aOn)
		aRetVal->AddRef();
	return OK;
}


FResult InputObject::set_On(ExprTokenType &aValue, IObject *&aOn, int aValidParamCount)
{
	auto obj = TokenToObject(aValue);
	if (obj)
	{
		auto fr = ValidateFunctor(obj, aValidParamCount);
		if (fr != OK)
			return fr;
		obj->AddRef();
	}
	else if (!TokenIsEmptyString(aValue))
		return FTypeError(_T("object"), aValue);
	auto prev = aOn;
	aOn = obj;
	if (prev)
		prev->Release();
	return OK;
}


FResult InputObject::set_MinSendLevel(int aValue)
{
	if (aValue < 0 || aValue > 101)
		return FR_E_ARG(0);
	input.MinSendLevel = aValue;
	return OK;
}


FResult InputObject::set_Timeout(double aValue)
{
	input.Timeout = (int)(aValue * 1000);
	if (input.InProgress() && input.Timeout > 0)
		input.SetTimeoutTimer();
	return OK;
}


FResult InputObject::KeyOpt(StrArg keys, StrArg options)
{
	bool adding = true;
	UCHAR flag, add_flags = 0, remove_flags = 0;
	for (auto cp = options; *cp; ++cp)
	{
		switch (ctoupper(*cp))
		{
		case '+': adding = true; continue;
		case '-': adding = false; continue;
		case ' ': case '\t': continue;
		case 'E': flag = END_KEY_ENABLED; break;
		case 'I': flag = INPUT_KEY_IGNORE_TEXT; break;
		case 'N': flag = INPUT_KEY_NOTIFY; break;
		case 'S':
			flag = INPUT_KEY_SUPPRESS;
			if (adding)
				remove_flags |= INPUT_KEY_VISIBLE;
			break;
		case 'V':
			flag = INPUT_KEY_VISIBLE;
			if (adding)
				remove_flags |= INPUT_KEY_SUPPRESS;
			break;
		case 'Z': // Zero (reset)
			add_flags = 0;
			remove_flags = INPUT_KEY_OPTION_MASK;
			continue;
		default: return FValueError(ERR_INVALID_OPTION, cp);
		}
		if (adding)
			add_flags |= flag; // Add takes precedence over remove, so remove_flags isn't changed.
		else
		{
			remove_flags |= flag;
			add_flags &= ~flag; // Override any previous add.
		}
	}
	
	if (!_tcsicmp(keys, _T("{All}")))
	{
		// Could optimize by using memset() when remove_flags == 0xFF, but that doesn't seem
		// worthwhile since this mode is already faster than SetKeyFlags() with a single key.
		for (int i = 0; i < _countof(input.KeyVK); ++i)
			input.KeyVK[i] = (input.KeyVK[i] & ~remove_flags) | add_flags;
		for (int i = 0; i < _countof(input.KeySC); ++i)
			input.KeySC[i] = (input.KeySC[i] & ~remove_flags) | add_flags;
		return OK;
	}

	return input.SetKeyFlags(keys, false, remove_flags, add_flags) ? OK : FR_FAIL;
}
