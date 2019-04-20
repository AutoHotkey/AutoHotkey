#include "stdafx.h" // pre-compiled headers
#include "defines.h"
#include "globaldata.h"
#include "script.h"
#include "application.h"

#include "script_object.h"
#include "script_func_impl.h"
#include "input_object.h"


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


ResultType InputObject::Invoke(ExprTokenType &aResultToken, ExprTokenType &aThisToken, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	if (!aParamCount)
		_o_throw(ERR_INVALID_USAGE);

	LPTSTR name = TokenToString(*aParam[0]);
	
	if (IS_INVOKE_CALL)
	{
		if (!_tcsicmp(name, _T("Start")))
		{
			if (input.InProgress())
				return OK;
			input.Buffer[input.BufferLength = 0] = '\0';
			return InputStart(input);
		}
		else if (!_tcsicmp(name, _T("Wait")))
		{
			UINT wait_ms = ParamIndexIsOmitted(1) ? UINT_MAX : (UINT)(ParamIndexToDouble(1) * 1000);
			DWORD tick_start = GetTickCount();
			while (input.InProgress() && (GetTickCount() - tick_start) < wait_ms)
				MsgSleep();
			aResultToken.symbol = SYM_STRING;
			aResultToken.marker = input.GetEndReason(NULL, 0, false);
			return OK;
		}
		else if (!_tcsicmp(name, _T("Stop")))
		{
			if (input.InProgress())
				input.Stop();
			return OK;
		}
		else if (!_tcsicmp(name, _T("KeyOpt")))
		{
			return KeyOpt(aResultToken, aParam + 1, aParamCount - 1);
		}
		return INVOKE_NOT_HANDLED;
	}

	if (aParamCount != (IS_INVOKE_SET ? 2 : 1))
		_o_throw(ERR_INVALID_USAGE);

	if (IS_INVOKE_GET)
	{
		if (!_tcsicmp(name, _T("Input")))
		{
			aResultToken.symbol = SYM_STRING;
			return TokenSetResult(aResultToken, input.Buffer, input.BufferLength);
		}
		else if (!_tcsicmp(name, _T("InProgress")))
		{
			aResultToken.symbol = SYM_INTEGER;
			aResultToken.value_int64 = input.InProgress();
			return OK;
		}
		else if (!_tcsicmp(name, _T("EndReason")))
		{
			aResultToken.symbol = SYM_STRING;
			aResultToken.marker = input.GetEndReason(NULL, 0, false);
			return OK;
		}
		else if (!_tcsicmp(name, _T("EndKey")))
		{
			aResultToken.symbol = SYM_STRING;
			if (input.Status == INPUT_TERMINATED_BY_ENDKEY)
				input.GetEndReason(aResultToken.marker = _f_retval_buf, _f_retval_buf_size, false);
			else
				aResultToken.marker = _T("");
			return OK;
		}
		else if (!_tcsicmp(name, _T("Match")))
		{
			aResultToken.symbol = SYM_STRING;
			if (input.Status == INPUT_TERMINATED_BY_MATCH && input.EndingMatchIndex < input.MatchCount)
				aResultToken.marker = input.match[input.EndingMatchIndex];
			else
				aResultToken.marker = _T("");
			return OK;
		}
	}
	if (!_tcsnicmp(name, _T("On"), 2))
	{
		IObject **pon;
		if (!_tcsicmp(name + 2, _T("End")))
			pon = &onEnd;
		else if (!_tcsicmp(name + 2, _T("KeyDown")))
			pon = &onKeyDown;
		else if (!_tcsicmp(name + 2, _T("Char")))
			pon = &onChar;
		else
			return INVOKE_NOT_HANDLED;

		if (IS_INVOKE_SET)
		{
			IObject *obj = ParamIndexToObject(1);
			if (obj)
				obj->AddRef();
			else if (!TokenIsEmptyString(*aParam[1]))
				_o_throw(ERR_INVALID_VALUE);
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
	// OPTIONS
	bool *bool_option = NULL;
	if (!_tcsicmp(name, _T("MinSendLevel")))
	{
		if (IS_INVOKE_SET)
			input.MinSendLevel = (SendLevelType)ParamIndexToInt64(1);
		aResultToken.SetValue(input.MinSendLevel);
		return OK;
	}
	else if (!_tcsicmp(name, _T("Timeout")))
	{
		if (IS_INVOKE_SET)
		{
			input.Timeout = (int)(ParamIndexToDouble(1) * 1000);
			if (input.InProgress() && input.Timeout > 0)
				input.SetTimeoutTimer();
		}
		aResultToken.SetValue(input.Timeout / 1000.0);
		return OK;
	}
	else if (!_tcsicmp(name, _T("BackspaceIsUndo"))) bool_option = &input.BackspaceIsUndo;
	else if (!_tcsicmp(name, _T("CaseSensitive"))) bool_option = &input.CaseSensitive;
	else if (!_tcsicmp(name, _T("FindAnywhere"))) bool_option = &input.FindAnywhere;
	else if (!_tcsicmp(name, _T("VisibleText"))) bool_option = &input.VisibleText;
	else if (!_tcsicmp(name, _T("VisibleNonText"))) bool_option = &input.VisibleNonText;
	else if (!_tcsicmp(name, _T("NotifyNonText"))) bool_option = &input.NotifyNonText;
	if (bool_option)
	{
		if (IS_INVOKE_SET)
			*bool_option = ParamIndexToBOOL(1);
		aResultToken.SetValue(*bool_option);
		return OK;
	}
	return INVOKE_NOT_HANDLED;
}


ResultType InputObject::KeyOpt(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	if (aParamCount < 2)
		_o_throw(ERR_TOO_FEW_PARAMS);

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
		case 'V':
			add_flags |= INPUT_KEY_VISIBILITY_OVERRIDE; // So -V can override the default visibility.
			flag = INPUT_KEY_VISIBLE;
			break;
		case 'Z': // Zero (reset)
			add_flags = 0;
			remove_flags = INPUT_KEY_OPTION_MASK;
			continue;
		default: _o_throw(ERR_INVALID_OPTION, cp);
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
