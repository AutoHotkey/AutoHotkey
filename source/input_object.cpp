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
	if (!_tcsicmp(name, _T("OnEnd")))
	{
		if (IS_INVOKE_SET)
		{
			IObject *obj = ParamIndexToObject(1);
			if (obj)
				obj->AddRef();
			else if (!TokenIsEmptyString(*aParam[1]))
				_o_throw(ERR_INVALID_VALUE);
			if (onEnd)
				onEnd->Release();
			onEnd = obj;
		}
		if (onEnd)
		{
			onEnd->AddRef();
			aResultToken.SetValue(onEnd);
		}
		return OK;
	}
	// OPTIONS
	else if (!_tcsicmp(name, _T("BackspaceIsUndo")))
	{
		if (IS_INVOKE_SET)
			input.BackspaceIsUndo = ParamIndexToBOOL(1);
		aResultToken.SetValue(input.BackspaceIsUndo);
		return OK;
	}
	else if (!_tcsicmp(name, _T("CaseSensitive")))
	{
		if (IS_INVOKE_SET)
			input.CaseSensitive = ParamIndexToBOOL(1);
		aResultToken.SetValue(input.CaseSensitive);
		return OK;
	}
	else if (!_tcsicmp(name, _T("FindAnywhere")))
	{
		if (IS_INVOKE_SET)
			input.FindAnywhere = ParamIndexToBOOL(1);
		aResultToken.SetValue(input.FindAnywhere);
		return OK;
	}
	else if (!_tcsicmp(name, _T("MinSendLevel")))
	{
		if (IS_INVOKE_SET)
			input.MinSendLevel = (SendLevelType)ParamIndexToInt64(1);
		aResultToken.SetValue(input.MinSendLevel);
		return OK;
	}
	else if (!_tcsicmp(name, _T("Timeout")))
	{
		if (IS_INVOKE_SET)
			input.Timeout = ParamIndexToInt(1);
		aResultToken.SetValue(input.Timeout);
		return OK;
	}
	return INVOKE_NOT_HANDLED;
}