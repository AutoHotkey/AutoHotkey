#include "stdafx.h" // pre-compiled headers
#include "defines.h"
#include "globaldata.h"
#include "script.h"
#include "application.h"

#include "script_object.h"
#include "script_func_impl.h"
#include "input_object.h"


ObjectMember InputObject::sMembers[] =
{
	Object_Method1(KeyOpt, 2, 2),
	Object_Method(Start, 0, 0),
	Object_Method(Wait, 0, 1),
	Object_Method(Stop, 0, 0),

	Object_Member(BackspaceIsUndo, BoolOpt, P_BackspaceIsUndo, IT_SET),
	Object_Member(CaseSensitive, BoolOpt, P_CaseSensitive, IT_SET),
	Object_Property_get(EndKey),
	Object_Property_get(EndMods),
	Object_Property_get(EndReason),
	Object_Member(FindAnywhere, BoolOpt, P_FindAnywhere, IT_SET),
	Object_Property_get(InProgress),
	Object_Property_get(Input),
	Object_Property_get(Match),
	Object_Property_get_set(MinSendLevel),
	Object_Member(NotifyNonText, BoolOpt, P_NotifyNonText, IT_SET),
	Object_Member(OnChar, OnX, P_OnChar, IT_SET),
	Object_Member(OnEnd, OnX, P_OnEnd, IT_SET),
	Object_Member(OnKeyDown, OnX, P_OnKeyDown, IT_SET),
	Object_Member(OnKeyUp, OnX, P_OnKeyUp, IT_SET),
	Object_Property_get_set(Timeout),
	Object_Member(VisibleNonText, BoolOpt, P_VisibleNonText, IT_SET),
	Object_Member(VisibleText, BoolOpt, P_VisibleText, IT_SET),
};
int InputObject::sMemberCount = _countof(sMembers);

Object *InputObject::sPrototype;

InputObject::InputObject()
{
	input.ScriptObject = this;
	SetBase(sPrototype);
}


BIF_DECL(BIF_InputHook)
{
	auto *input_handle = new InputObject();
	
	_f_param_string_opt(aOptions, 0);
	_f_param_string_opt(aEndKeys, 1);
	_f_param_string_opt(aMatchList, 2);
	
	if (!input_handle->Setup(aOptions, aEndKeys, aMatchList, _tcslen(aMatchList)))
	{
		input_handle->Release();
		_f_return_FAIL;
	}

	aResultToken.symbol = SYM_OBJECT;
	aResultToken.object = input_handle;
}


ResultType InputObject::Setup(LPTSTR aOptions, LPTSTR aEndKeys, LPTSTR aMatchList, size_t aMatchList_length)
{
	return input.Setup(aOptions, aEndKeys, aMatchList, aMatchList_length);
}


ResultType InputObject::Invoke(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	switch (aID)
	{
	case M_Start:
		if (input.InProgress())
			return OK;
		input.Buffer[input.BufferLength = 0] = '\0';
		return InputStart(input);

	case M_Wait:
		DWORD wait_ms, tick_start;
		wait_ms = ParamIndexIsOmitted(0) ? UINT_MAX : (UINT)(ParamIndexToDouble(0) * 1000);
		tick_start = GetTickCount();
		while (input.InProgress() && (GetTickCount() - tick_start) < wait_ms)
			MsgSleep();
		// Return EndReason:
	case P_EndReason:
		_o_return(input.GetEndReason(NULL, 0));

	case M_Stop:
		if (input.InProgress())
			input.Stop();
		return OK;

	case P_Input:
		_o_return(input.Buffer, input.BufferLength);

	case P_InProgress:
		_o_return(input.InProgress());

	case P_EndKey:
		aResultToken.symbol = SYM_STRING;
		if (input.Status == INPUT_TERMINATED_BY_ENDKEY)
			input.GetEndReason(aResultToken.marker = _f_retval_buf, _f_retval_buf_size);
		else
			aResultToken.marker = _T("");
		return OK;

	case P_EndMods:
	{
		aResultToken.symbol = SYM_STRING;
		TCHAR *cp = aResultToken.marker = aResultToken.buf;
		const auto mod_string = MODLR_STRING;
		for (int i = 0; i < 8; ++i)
			if (input.EndingMods & (1 << i))
			{
				*cp++ = mod_string[i * 2];
				*cp++ = mod_string[i * 2 + 1];
			}
		*cp = '\0';
		return OK;
	}

	case P_Match:
		aResultToken.symbol = SYM_STRING;
		if (input.Status == INPUT_TERMINATED_BY_MATCH && input.EndingMatchIndex < input.MatchCount)
			aResultToken.marker = input.match[input.EndingMatchIndex];
		else
			aResultToken.marker = _T("");
		return OK;

	case P_MinSendLevel:
		if (IS_INVOKE_SET)
		{
			input.MinSendLevel = (SendLevelType)ParamIndexToInt64(0);
			return OK;
		}
		_o_return(input.MinSendLevel);

	case P_Timeout:
		if (IS_INVOKE_SET)
		{
			input.Timeout = (int)(ParamIndexToDouble(0) * 1000);
			if (input.InProgress() && input.Timeout > 0)
				input.SetTimeoutTimer();
			return OK;
		}
		_o_return(input.Timeout / 1000.0);
	}
	return INVOKE_NOT_HANDLED;
}


ResultType InputObject::BoolOpt(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	bool *bool_option = nullptr;
	switch (aID)
	{
	case P_BackspaceIsUndo: bool_option = &input.BackspaceIsUndo; break;
	case P_CaseSensitive: bool_option = &input.CaseSensitive; break;
	case P_FindAnywhere: bool_option = &input.FindAnywhere; break;
	case P_VisibleText: bool_option = &input.VisibleText; break;
	case P_VisibleNonText: bool_option = &input.VisibleNonText; break;
	case P_NotifyNonText: bool_option = &input.NotifyNonText; break;
	}
	if (IS_INVOKE_SET)
		*bool_option = ParamIndexToBOOL(0);
	_o_return(*bool_option);
}


ResultType InputObject::OnX(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	IObject **pon;
	switch (aID)
	{
	case P_OnKeyDown: pon = &onKeyDown; break;
	case P_OnKeyUp: pon = &onKeyUp; break;
	case P_OnChar: pon = &onChar; break;
	default: pon = &onEnd; break;
	}
	if (IS_INVOKE_SET)
	{
		IObject *obj = ParamIndexToObject(0);
		if (obj)
			obj->AddRef();
		else if (!TokenIsEmptyString(*aParam[0]))
			_o_throw_type(_T("object"), *aParam[0]);
		if (*pon)
			(*pon)->Release();
		*pon = obj;
	}
	if (*pon)
	{
		(*pon)->AddRef();
		aResultToken.SetValue(*pon);
	}
	return OK;
}


ResultType InputObject::KeyOpt(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	_f_param_string(keys, 0);
	_f_param_string(options, 1);

	bool adding = true;
	UCHAR flag, add_flags = 0, remove_flags = 0;
	for (LPTSTR cp = options; *cp; ++cp)
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
		default: _o_throw_value(ERR_INVALID_OPTION, cp);
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

	return input.SetKeyFlags(keys, false, remove_flags, add_flags);
}
