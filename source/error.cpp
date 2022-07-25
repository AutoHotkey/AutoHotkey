/*
AutoHotkey

Copyright 2003-2009 Chris Mallett (support@autohotkey.com)

This program is free software; you can redistribute it and/or
modify it under the terms of the GNU General Public License
as published by the Free Software Foundation; either version 2
of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.
*/

#include "stdafx.h"
#include "script.h"
#include "globaldata.h"
#include "window.h"
#include "TextIO.h"


Line *Line::PreparseError(LPTSTR aErrorText, LPTSTR aExtraInfo)
// Returns a different type of result for use with the Pre-parsing methods.
{
	// Make all preparsing errors critical because the runtime reliability
	// of the program relies upon the fact that the aren't any kind of
	// problems in the script (otherwise, unexpected behavior may result).
	// Update: It's okay to return FAIL in this case.  CRITICAL_ERROR should
	// be avoided whenever OK and FAIL are sufficient by themselves, because
	// otherwise, callers can't use the NOT operator to detect if a function
	// failed (since FAIL is value zero, but CRITICAL_ERROR is non-zero):
	LineError(aErrorText, FAIL, aExtraInfo);
	return NULL; // Always return NULL because the callers use it as their return value.
}


#ifdef CONFIG_DEBUGGER
LPCTSTR Debugger::WhatThrew()
{
	// We want 'What' to indicate the function/sub/operation that *threw* the exception.
	// For BIFs, throwing is always explicit.  For a UDF, 'What' should only name it if
	// it explicitly constructed the Exception object.  This provides an easy way for
	// OnError and Catch to categorise errors.  No information is lost because File/Line
	// can already be used locate the function/sub that was running.
	// So only return a name when a BIF is raising an error:
	if (mStack.mTop < mStack.mBottom || mStack.mTop->type != DbgStack::SE_BIF)
		return _T("");
	return mStack.mTop->func->mName;
}
#endif


IObject *Line::CreateRuntimeException(LPCTSTR aErrorText, LPCTSTR aExtraInfo, Object *aPrototype)
{
	// Build the parameters for Object::Create()
	ExprTokenType aParams[3]; int aParamCount = 2;
	ExprTokenType* aParam[3] { aParams + 0, aParams + 1, aParams + 2 };
	aParams[0].SetValue(const_cast<LPTSTR>(aErrorText));
#ifdef CONFIG_DEBUGGER
	aParams[1].SetValue(const_cast<LPTSTR>(g_Debugger.WhatThrew()));
#else
	// Without the debugger stack, there's no good way to determine what's throwing. It could be:
	//g_act[mActionType].Name; // A command implemented as an Action (g_act).
	//g->CurrentFunc->mName; // A user-defined function (perhaps when mActionType == ACT_THROW).
	//???; // A built-in function implemented as a Func (g_BIF).
	aParams[1].SetValue(_T(""), 0);
#endif
	if (aExtraInfo && *aExtraInfo)
		aParams[aParamCount++].SetValue(const_cast<LPTSTR>(aExtraInfo));

	auto obj = Object::Create();
	if (!obj)
		return nullptr;
	if (!aPrototype)
		aPrototype = ErrorPrototype::Error;
	obj->SetBase(aPrototype);
	FuncResult rt;
	g_script.mCurrLine = this;
	if (!obj->Construct(rt, aParam, aParamCount))
		return nullptr;
	return obj;
}


ResultType Line::ThrowRuntimeException(LPCTSTR aErrorText, LPCTSTR aExtraInfo)
{
	return g_script.ThrowRuntimeException(aErrorText, aExtraInfo, this, FAIL);
}

ResultType Script::ThrowRuntimeException(LPCTSTR aErrorText, LPCTSTR aExtraInfo
	, Line *aLine, ResultType aErrorType, Object *aPrototype)
{
	// ThrownToken should only be non-NULL while control is being passed up the
	// stack, which implies no script code can be executing.
	ASSERT(!g->ThrownToken);

	if (!aLine || !aLine->mLineNumber) // The mLineNumber check is a workaround for BIF_PerformAction.
		aLine = mCurrLine;

	ResultToken *token;
	if (   !(token = new ResultToken)
		|| !(token->object = aLine->CreateRuntimeException(aErrorText, aExtraInfo, aPrototype))   )
	{
		// Out of memory. It's likely that we were called for this very reason.
		// Since we don't even have enough memory to allocate an exception object,
		// just show an error message and exit the thread. Don't call LineError(),
		// since that would recurse into this function.
		if (token)
			delete token;
		if (!g->ThrownToken)
		{
			MsgBox(ERR_OUTOFMEM ERR_ABORT);
			return FAIL;
		}
		//else: Thrown by Error constructor?
	}
	else
	{
		token->symbol = SYM_OBJECT;
		token->mem_to_free = NULL;

		g->ThrownToken = token;
	}
	if (!(g->ExcptMode & EXCPTMODE_CATCH))
		return UnhandledException(aLine, aErrorType);

	// Returning FAIL causes each caller to also return FAIL, until either the
	// thread has fully exited or the recursion layer handling ACT_TRY is reached:
	return FAIL;
}

ResultType Script::ThrowRuntimeException(LPCTSTR aErrorText, LPCTSTR aExtraInfo)
{
	return ThrowRuntimeException(aErrorText, aExtraInfo, mCurrLine, FAIL);
}


ResultType Script::Win32Error(DWORD aError, ResultType aErrorType)
{
	TCHAR number_string[_MAX_ULTOSTR_BASE10_COUNT];
	// Convert aError to string to pass it through RuntimeError, but it will ultimately
	// be converted to the error number and proper message by OSError.Prototype.__New.
	_ultot(aError, number_string, 10);
	return RuntimeError(number_string, _T(""), aErrorType, nullptr, ErrorPrototype::OS);
}


ResultType Line::ThrowIfTrue(bool aError)
{
	return aError ? ThrowRuntimeException(ERR_FAILED) : OK;
}

ResultType Line::ThrowIntIfNonzero(int aErrorValue)
{
	if (!aErrorValue)
		return OK;
	TCHAR buf[12];
	return ThrowRuntimeException(ERR_FAILED, _itot(aErrorValue, buf, 10));
}

// Logic from the above functions is duplicated in the below functions rather than calling
// g_script.mCurrLine->Throw() to squeeze out a little extra performance for
// "success" cases.

ResultType Script::ThrowIfTrue(bool aError)
{
	return aError ? ThrowRuntimeException(ERR_FAILED) : OK;
}

ResultType Script::ThrowIntIfNonzero(int aErrorValue)
{
	if (!aErrorValue)
		return OK;
	TCHAR buf[12];
	return ThrowRuntimeException(ERR_FAILED, _itot(aErrorValue, buf, 10));
}


ResultType Line::SetLastErrorMaybeThrow(bool aError, DWORD aLastError)
{
	g->LastError = aLastError; // Set this unconditionally.
	return aError ? g_script.Win32Error(aLastError) : OK;
}

void ResultToken::SetLastErrorMaybeThrow(bool aError, DWORD aLastError)
{
	g->LastError = aLastError; // Set this unconditionally.
	if (aError)
		Win32Error(aLastError);
}

void ResultToken::SetLastErrorCloseAndMaybeThrow(HANDLE aHandle, bool aError, DWORD aLastError)
{
	g->LastError = aLastError;
	CloseHandle(aHandle);
	if (aError)
		Win32Error(aLastError);
}


void Script::SetErrorStdOut(LPTSTR aParam)
{
	mErrorStdOutCP = Line::ConvertFileEncoding(aParam);
	// Seems best not to print errors to stderr if the encoding was invalid.  Current behaviour
	// for an encoding of -1 would be to print only the ASCII characters and drop the rest, but
	// if our caller is expecting UTF-16, it won't be readable.
	mErrorStdOut = mErrorStdOutCP != -1;
	// If invalid, no error is shown here because this function might be called early, before
	// Line::sSourceFile[0] is given its value.  Instead, errors appearing as dialogs should
	// be a sufficient clue that the /ErrorStdOut= value was invalid.
}

void Script::PrintErrorStdOut(LPCTSTR aErrorText, int aLength, LPCTSTR aFile)
{
#ifdef CONFIG_DEBUGGER
	if (g_Debugger.OutputStdOut(aErrorText))
		return;
#endif
	TextFile tf;
	tf.Open(aFile, TextStream::APPEND, mErrorStdOutCP);
	tf.Write(aErrorText, aLength);
	tf.Close();
}

// For backward compatibility, this actually prints to stderr, not stdout.
void Script::PrintErrorStdOut(LPCTSTR aErrorText, LPCTSTR aExtraInfo, FileIndexType aFileIndex, LineNumberType aLineNumber)
{
	TCHAR buf[LINE_SIZE * 2];
#define STD_ERROR_FORMAT _T("%s (%d) : ==> %s\n")
	int n = sntprintf(buf, _countof(buf), STD_ERROR_FORMAT, Line::sSourceFile[aFileIndex], aLineNumber, aErrorText);
	if (*aExtraInfo)
		n += sntprintf(buf + n, _countof(buf) - n, _T("     Specifically: %s\n"), aExtraInfo);
	PrintErrorStdOut(buf, n, _T("**"));
}

ResultType Line::LineError(LPCTSTR aErrorText, ResultType aErrorType, LPCTSTR aExtraInfo)
{
	ASSERT(aErrorText);
	if (!aExtraInfo)
		aExtraInfo = _T("");

	if (g_script.mIsReadyToExecute)
	{
		return g_script.RuntimeError(aErrorText, aExtraInfo, aErrorType, this);
	}

	if (g_script.mErrorStdOut && aErrorType != WARN)
	{
		// JdeB said:
		// Just tested it in Textpad, Crimson and Scite. they all recognise the output and jump
		// to the Line containing the error when you double click the error line in the output
		// window (like it works in C++).  Had to change the format of the line to: 
		// printf("%s (%d) : ==> %s: \n%s \n%s\n",szInclude, nAutScriptLine, szText, szScriptLine, szOutput2 );
		// MY: Full filename is required, even if it's the main file, because some editors (EditPlus)
		// seem to rely on that to determine which file and line number to jump to when the user double-clicks
		// the error message in the output window.
		// v1.0.47: Added a space before the colon as originally intended.  Toralf said, "With this minor
		// change the error lexer of Scite recognizes this line as a Microsoft error message and it can be
		// used to jump to that line."
		g_script.PrintErrorStdOut(aErrorText, aExtraInfo, mFileIndex, mLineNumber);
		return FAIL;
	}

	return g_script.ShowError(aErrorText, aErrorType, aExtraInfo, this);
}

ResultType Script::RuntimeError(LPCTSTR aErrorText, LPCTSTR aExtraInfo, ResultType aErrorType, Line *aLine, Object *aPrototype)
{
	ASSERT(aErrorText);
	if (!aExtraInfo)
		aExtraInfo = _T("");
	
	if (g->ExcptMode == EXCPTMODE_LINE_WORKAROUND && mCurrLine)
		aLine = mCurrLine;
	
	if ((g->ExcptMode || mOnError.Count() || aPrototype && aPrototype->HasOwnProps()) && aErrorType != WARN)
		return ThrowRuntimeException(aErrorText, aExtraInfo, aLine, aErrorType, aPrototype);

	return ShowError(aErrorText, aErrorType, aExtraInfo, aLine);
}

ResultType Script::ShowError(LPCTSTR aErrorText, ResultType aErrorType, LPCTSTR aExtraInfo, Line *aLine)
{
	if (!aErrorText)
		aErrorText = _T("");
	if (!aExtraInfo)
		aExtraInfo = _T("");
	if (!aLine)
		aLine = mCurrLine;

	TCHAR buf[MSGBOX_TEXT_SIZE];
	FormatError(buf, _countof(buf), aErrorType, aErrorText, aExtraInfo, aLine);
		
	// It's currently unclear why this would ever be needed, so it's disabled:
	//mCurrLine = aLine;  // This needs to be set in some cases where the caller didn't.
		
#ifdef CONFIG_DEBUGGER
	g_Debugger.OutputStdErr(buf);
#endif
	if (MsgBox(buf, MB_TOPMOST | (aErrorType == FAIL_OR_OK ? MB_YESNO|MB_DEFBUTTON2 : 0)) == IDYES)
		return OK;

	if (aErrorType == CRITICAL_ERROR && mIsReadyToExecute)
		// Pass EXIT_DESTROY to ensure the program always exits, regardless of OnExit.
		ExitApp(EXIT_CRITICAL);

	// Since above didn't exit, the caller isn't CriticalError(), which ignores
	// the return value.  Other callers always want FAIL at this point.
	return FAIL;
}



int Script::FormatError(LPTSTR aBuf, int aBufSize, ResultType aErrorType, LPCTSTR aErrorText, LPCTSTR aExtraInfo, Line *aLine)
{
	TCHAR source_file[MAX_PATH * 2];
	if (aLine && aLine->mFileIndex)
		sntprintf(source_file, _countof(source_file), _T(" in #include file \"%s\""), Line::sSourceFile[aLine->mFileIndex]);
	else
		*source_file = '\0'; // Don't bother cluttering the display if it's the main script file.

	LPTSTR aBuf_orig = aBuf;
	// Error message:
	aBuf += sntprintf(aBuf, aBufSize, _T("%s%s:%s %-1.500s\n\n")  // Keep it to a sane size in case it's huge.
		, aErrorType == WARN ? _T("Warning") : (aErrorType == CRITICAL_ERROR ? _T("Critical Error") : _T("Error"))
		, source_file, *source_file ? _T("\n    ") : _T(" "), aErrorText);
	// Specifically:
	if (*aExtraInfo)
		// Use format specifier to make sure really huge strings that get passed our
		// way, such as a var containing clipboard text, are kept to a reasonable size:
		aBuf += sntprintf(aBuf, BUF_SPACE_REMAINING, _T("Specifically: %-1.100s%s\n\n")
			, aExtraInfo, _tcslen(aExtraInfo) > 100 ? _T("...") : _T(""));
	// Relevant lines of code:
	if (aLine)
		aBuf = aLine->VicinityToText(aBuf, BUF_SPACE_REMAINING);
	// What now?:
	LPCTSTR footer;
	switch (aErrorType)
	{
	case WARN: footer = ERR_WARNING_FOOTER; break;
	case FAIL_OR_OK: footer = ERR_CONTINUE_THREAD_Q; break;
	case CRITICAL_ERROR: footer = UNSTABLE_WILL_EXIT; break;
	default: footer = (g->ExcptMode & EXCPTMODE_DELETE) ? ERR_ABORT_DELETE
		: mIsReadyToExecute ? ERR_ABORT_NO_SPACES
		: mIsRestart ? OLD_STILL_IN_EFFECT
		: WILL_EXIT;
	}
	aBuf += sntprintf(aBuf, BUF_SPACE_REMAINING, _T("\n%s"), footer);
	
	return (int)(aBuf - aBuf_orig);
}



ResultType Script::ScriptError(LPCTSTR aErrorText, LPCTSTR aExtraInfo) //, ResultType aErrorType)
// Even though this is a Script method, including it here since it shares
// a common theme with the other error-displaying functions:
{
	if (mCurrLine)
		// If a line is available, do LineError instead since it's more specific.
		// If an error occurs before the script is ready to run, assume it's always critical
		// in the sense that the program will exit rather than run the script.
		// Update: It's okay to return FAIL in this case.  CRITICAL_ERROR should
		// be avoided whenever OK and FAIL are sufficient by themselves, because
		// otherwise, callers can't use the NOT operator to detect if a function
		// failed (since FAIL is value zero, but CRITICAL_ERROR is non-zero):
		return mCurrLine->LineError(aErrorText, FAIL, aExtraInfo);
	// Otherwise: The fact that mCurrLine is NULL means that the line currently being loaded
	// has not yet been successfully added to the linked list.  Such errors will always result
	// in the program exiting.
	if (!aErrorText)
		aErrorText = _T("Unk"); // Placeholder since it shouldn't be NULL.
	if (!aExtraInfo) // In case the caller explicitly called it with NULL.
		aExtraInfo = _T("");

	if (g_script.mErrorStdOut && !g_script.mIsReadyToExecute) // i.e. runtime errors are always displayed via dialog.
	{
		// See LineError() for details.
		PrintErrorStdOut(aErrorText, aExtraInfo, mCurrFileIndex, mCombinedLineNumber);
	}
	else
	{
		TCHAR buf[MSGBOX_TEXT_SIZE], *cp = buf;
		int buf_space_remaining = (int)_countof(buf);

		if (mCombinedLineNumber || mCurrFileIndex)
		{
			cp += sntprintf(cp, buf_space_remaining, _T("Error at line %u"), mCombinedLineNumber); // Don't call it "critical" because it's usually a syntax error.
			buf_space_remaining = (int)(_countof(buf) - (cp - buf));
		}

		if (mCurrFileIndex)
		{
			cp += sntprintf(cp, buf_space_remaining, _T(" in #include file \"%s\""), Line::sSourceFile[mCurrFileIndex]);
			buf_space_remaining = (int)(_countof(buf) - (cp - buf));
		}
		//else don't bother cluttering the display if it's the main script file.

		if (mCombinedLineNumber || mCurrFileIndex)
		{
			cp += sntprintf(cp, buf_space_remaining, _T(".\n\n"));
			buf_space_remaining = (int)(_countof(buf) - (cp - buf));
		}

		if (*aExtraInfo)
		{
			cp += sntprintf(cp, buf_space_remaining, _T("Line Text: %-1.100s%s\nError: ")  // i.e. the word "Error" is omitted as being too noisy when there's no ExtraInfo to put into the dialog.
				, aExtraInfo // aExtraInfo defaults to "" so this is safe.
				, _tcslen(aExtraInfo) > 100 ? _T("...") : _T(""));
			buf_space_remaining = (int)(_countof(buf) - (cp - buf));
		}
		sntprintf(cp, buf_space_remaining, _T("%s\n\n%s"), aErrorText, mIsRestart ? OLD_STILL_IN_EFFECT : WILL_EXIT);

		//ShowInEditor();
#ifdef CONFIG_DEBUGGER
		if (!g_Debugger.OutputStdErr(buf))
#endif
		MsgBox(buf);
	}
	return FAIL; // See above for why it's better to return FAIL than CRITICAL_ERROR.
}



LPCTSTR VarKindForErrorMessage(Var *aVar)
{
	switch (aVar->Type())
	{
	case VAR_VIRTUAL: return _T("built-in variable");
	case VAR_CONSTANT: return aVar->Object()->Type();
	default: return Var::DeclarationType(aVar->Scope());
	}
}

ResultType Script::ConflictingDeclarationError(LPCTSTR aDeclType, Var *aExisting)
{
	TCHAR buf[127];
	sntprintf(buf, _countof(buf), _T("This %s declaration conflicts with an existing %s.")
		, aDeclType, VarKindForErrorMessage(aExisting));
	return ScriptError(buf, aExisting->mName);
}


ResultType Line::ValidateVarUsage(Var *aVar, int aUsage)
{
	if (   VARREF_IS_WRITE(aUsage)
		&& (aUsage == VARREF_REF
			? aVar->Type() != VAR_NORMAL // Aliasing VAR_VIRTUAL is currently unsupported.
			: aVar->IsReadOnly())   )
		return VarIsReadOnlyError(aVar, aUsage);
	return OK;
}

ResultType Script::VarIsReadOnlyError(Var *aVar, int aErrorType)
{
	TCHAR buf[127];
	sntprintf(buf, _countof(buf), _T("This %s cannot %s.")
		, VarKindForErrorMessage(aVar)
		, aErrorType == VARREF_LVALUE ? _T("be assigned a value")
		: aErrorType == VARREF_REF ? _T("have its reference taken")
		: _T("be used as an output variable"));
	return ScriptError(buf, aVar->mName);
}

ResultType Line::VarIsReadOnlyError(Var *aVar, int aErrorType)
{
	g_script.mCurrLine = this;
	return g_script.VarIsReadOnlyError(aVar, aErrorType);
}


ResultType Line::LineUnexpectedError()
{
	TCHAR buf[127];
	sntprintf(buf, _countof(buf), _T("Unexpected \"%s\""), g_act[mActionType].Name);
	return LineError(buf);
}



ResultType Script::CriticalError(LPCTSTR aErrorText, LPCTSTR aExtraInfo)
{
	g->ExcptMode = EXCPTMODE_NONE; // Do not throw an exception.
	if (mCurrLine)
		mCurrLine->LineError(aErrorText, CRITICAL_ERROR, aExtraInfo);
	// mCurrLine should always be non-NULL during runtime, and CRITICAL_ERROR should
	// cause LineError() to exit even if an OnExit routine is present, so this is here
	// mainly for maintainability.
	TerminateApp(EXIT_CRITICAL, 0);
	return FAIL; // Never executed.
}



__declspec(noinline)
ResultType ResultToken::Error(LPCTSTR aErrorText)
{
	// Defining this overload separately rather than making aErrorInfo optional reduces code size
	// by not requiring the compiler to 'push' the second parameter's default value at each call site.
	return Error(aErrorText, _T(""));
}

__declspec(noinline)
ResultType ResultToken::Error(LPCTSTR aErrorText, LPCTSTR aExtraInfo)
{
	return Error(aErrorText, aExtraInfo, nullptr);
}

__declspec(noinline)
ResultType ResultToken::Error(LPCTSTR aErrorText, Object *aPrototype)
{
	return Error(aErrorText, nullptr, aPrototype);
}

__declspec(noinline)
ResultType ResultToken::Error(LPCTSTR aErrorText, LPCTSTR aExtraInfo, Object *aPrototype)
{
	// These two assertions should always pass, since anything else would imply returning a value,
	// not throwing an error.  If they don't, the memory/object might not be freed since the caller
	// isn't expecting a value, or they might be freed twice (if the callee already freed it).
	//ASSERT(!mem_to_free); // At least one caller frees it after calling this function.
	ASSERT(symbol != SYM_OBJECT);
	if (g_script.RuntimeError(aErrorText, aExtraInfo, FAIL_OR_OK, nullptr, aPrototype) == FAIL)
		return SetExitResult(FAIL);
	SetValue(_T(""), 0);
	// Caller may rely on FAIL to unwind stack, but this->result is still OK.
	return FAIL;
}

__declspec(noinline)
ResultType ResultToken::MemoryError()
{
	return Error(ERR_OUTOFMEM, nullptr, ErrorPrototype::Memory);
}

ResultType MemoryError()
{
	return g_script.RuntimeError(ERR_OUTOFMEM, nullptr, FAIL, nullptr, ErrorPrototype::Memory);
}

void SimpleHeap::CriticalFail()
{
	g_script.CriticalError(ERR_OUTOFMEM);
}

__declspec(noinline)
ResultType ResultToken::ValueError(LPCTSTR aErrorText)
{
	return Error(aErrorText, nullptr, ErrorPrototype::Value);
}

__declspec(noinline)
ResultType ResultToken::ValueError(LPCTSTR aErrorText, LPCTSTR aExtraInfo)
{
	return Error(aErrorText, aExtraInfo, ErrorPrototype::Value);
}

__declspec(noinline)
ResultType ValueError(LPCTSTR aErrorText, LPCTSTR aExtraInfo, ResultType aErrorType)
{
	if (!g_script.mIsReadyToExecute)
		return g_script.ScriptError(aErrorText, aExtraInfo);
	return g_script.RuntimeError(aErrorText, aExtraInfo, aErrorType, nullptr, ErrorPrototype::Value);
}

__declspec(noinline)
ResultType ResultToken::UnknownMemberError(ExprTokenType &aObject, int aFlags, LPCTSTR aMember)
{
	TCHAR msg[512];
	if (!aMember)
		aMember = (aFlags & IT_CALL) ? _T("Call") : _T("__Item");
	sntprintf(msg, _countof(msg), _T("This value of type \"%s\" has no %s named \"%s\".")
		, TokenTypeString(aObject), (aFlags & IT_CALL) ? _T("method") : _T("property"), aMember);
	return Error(msg, nullptr, (aFlags & IT_CALL) ? ErrorPrototype::Method : ErrorPrototype::Property);
}

__declspec(noinline)
ResultType ResultToken::Win32Error(DWORD aError)
{
	if (g_script.Win32Error(aError) == FAIL)
		return SetExitResult(FAIL);
	SetValue(_T(""), 0);
	return FAIL;
}

void TokenTypeAndValue(ExprTokenType &aToken, LPTSTR &aType, LPTSTR &aValue, TCHAR *aNBuf)
{
	if (aToken.symbol == SYM_VAR && aToken.var->IsUninitializedNormalVar())
		aType = _T("unset variable"), aValue = aToken.var->mName;
	else if (TokenIsEmptyString(aToken))
		aType = _T("empty string"), aValue = _T("");
	else
		aType = TokenTypeString(aToken), aValue = TokenToString(aToken, aNBuf);
}

__declspec(noinline)
ResultType ResultToken::TypeError(LPCTSTR aExpectedType, ExprTokenType &aActualValue)
{
	TCHAR number_buf[MAX_NUMBER_SIZE];
	LPTSTR actual_type, value_as_string;
	TokenTypeAndValue(aActualValue, actual_type, value_as_string, number_buf);
	return TypeError(aExpectedType, actual_type, value_as_string);
}

ResultType ResultToken::TypeError(LPCTSTR aExpectedType, LPCTSTR aActualType, LPTSTR aExtraInfo)
{
	auto an = [](LPCTSTR thing) {
		return _tcschr(_T("aeiou"), ctolower(*thing)) ? _T("n") : _T("");
	};
	TCHAR msg[512];
	sntprintf(msg, _countof(msg), _T("Expected a%s %s but got a%s %s.")
		, an(aExpectedType), aExpectedType, an(aActualType), aActualType);
	return Error(msg, aExtraInfo, ErrorPrototype::Type);
}

__declspec(noinline)
ResultType ResultToken::ParamError(int aIndex, ExprTokenType *aParam)
{
	return ParamError(aIndex, aParam, nullptr, nullptr);
}

__declspec(noinline)
ResultType ResultToken::ParamError(int aIndex, ExprTokenType *aParam, LPCTSTR aExpectedType)
{
	return ParamError(aIndex, aParam, aExpectedType, nullptr);
}

__declspec(noinline)
ResultType ResultToken::ParamError(int aIndex, ExprTokenType *aParam, LPCTSTR aExpectedType, LPCTSTR aFunction)
{
	auto an = [](LPCTSTR thing) {
		return _tcschr(_T("aeiou"), ctolower(*thing)) ? _T("n") : _T("");
	};
	TCHAR msg[512];
	TCHAR number_buf[MAX_NUMBER_SIZE];
	LPTSTR actual_type, value_as_string;
	TokenTypeAndValue(*aParam, actual_type, value_as_string, number_buf);
	if (!*value_as_string && !aExpectedType)
		value_as_string = actual_type;
#ifdef CONFIG_DEBUGGER
	if (!aFunction)
		aFunction = g_Debugger.WhatThrew();
	if (aExpectedType)
		sntprintf(msg, _countof(msg), _T("Parameter #%i of %s requires a%s %s, but received a%s %s.")
			, aIndex + 1, aFunction, an(aExpectedType), aExpectedType, an(actual_type), actual_type);
	else
		sntprintf(msg, _countof(msg), _T("Parameter #%i of %s is invalid."), aIndex + 1, g_Debugger.WhatThrew());
#else
	if (aExpectedType)
		sntprintf(msg, _countof(msg), _T("Parameter #%i requires a%s %s, but received a%s %s.")
			, aIndex + 1, an(aExpectedType), aExpectedType, an(actual_type), actual_type);
	else
		sntprintf(msg, _countof(msg), _T("Parameter #%i invalid."), aIndex + 1);
#endif
	return Error(msg, value_as_string, ErrorPrototype::Type);
}



ResultType Script::UnhandledException(Line* aLine, ResultType aErrorType)
{
	LPCTSTR message = _T(""), extra = _T("");
	TCHAR extra_buf[MAX_NUMBER_SIZE], message_buf[MAX_NUMBER_SIZE];

	global_struct &g = *::g;
	
	ResultToken *token = g.ThrownToken;
	// Clear ThrownToken to allow any applicable callbacks to execute correctly.
	// This includes OnError callbacks explicitly called below, but also COM events
	// and CallbackCreate callbacks that execute while MsgBox() is waiting.
	g.ThrownToken = NULL;

	// OnError: Allow the script to handle it via a global callback.
	static bool sOnErrorRunning = false;
	if (mOnError.Count() && !sOnErrorRunning)
	{
		__int64 retval;
		sOnErrorRunning = true;
		ExprTokenType param[2];
		param[0].CopyValueFrom(*token);
		param[1].SetValue(aErrorType == CRITICAL_ERROR ? _T("ExitApp")
			: aErrorType == FAIL_OR_OK ? _T("Return") : _T("Exit"));
		mOnError.Call(param, 2, INT_MAX, &retval);
		sOnErrorRunning = false;
		if (g.ThrownToken) // An exception was thrown by the callback.
		{
			// UnhandledException() has already been called recursively for g.ThrownToken,
			// so don't show a second error message.  This allows `throw param1` to mean
			// "abort all OnError callbacks and show default message now".
			FreeExceptionToken(token);
			return FAIL;
		}
		if (retval < 0 && aErrorType == FAIL_OR_OK)
		{
			FreeExceptionToken(token);
			return OK; // Ignore error and continue.
		}
		// Some callers rely on g.ThrownToken!=NULL to unwind the stack, so it is restored
		// rather than freeing it immediately.  If the exception object has __Delete, it
		// will be called after the stack unwinds.
		if (retval)
		{
			g.ThrownToken = token;
			return FAIL; // Exit thread.
		}
	}

	if (Object *ex = dynamic_cast<Object *>(TokenToObject(*token)))
	{
		// For simplicity and safety, we call into the Object directly rather than via Invoke().
		ExprTokenType t;
		if (ex->GetOwnProp(t, _T("Message")))
			message = TokenToString(t, message_buf);
		if (ex->GetOwnProp(t, _T("Extra")))
			extra = TokenToString(t, extra_buf);
		if (ex->GetOwnProp(t, _T("Line")))
		{
			LineNumberType line_no = (LineNumberType)TokenToInt64(t);
			if (ex->GetOwnProp(t, _T("File")))
			{
				LPCTSTR file = TokenToString(t);
				// Locate the line by number and file index, then display that line instead
				// of the caller supplied one since it's probably more relevant.
				int file_index;
				for (file_index = 0; file_index < Line::sSourceFileCount; ++file_index)
					if (!_tcsicmp(file, Line::sSourceFile[file_index]))
						break;
				if (!aLine || aLine->mFileIndex != file_index || aLine->mLineNumber != line_no) // Keep aLine if it matches, in case of multiple Lines with the same number.
				{
					Line *line;
					for (line = mFirstLine;
						line && (line->mLineNumber != line_no || line->mFileIndex != file_index
							|| !line->mArgc && line->mNextLine && line->mNextLine->mLineNumber == line_no); // Skip any same-line block-begin/end, try, else or finally.
						line = line->mNextLine);
					if (line)
						aLine = line;
				}
			}
		}
	}
	else
	{
		// Assume it's a string or number.
		extra = TokenToString(*token, message_buf);
	}

	// If message is empty (or a string or number was thrown), display a default message for clarity.
	if (!*message)
		message = _T("Unhandled exception.");

	if (ShowError(message, aErrorType, extra, aLine) == OK)
	{
		FreeExceptionToken(token);
		return OK;
	}
	g.ThrownToken = token;
	return FAIL;
}



void Script::FreeExceptionToken(ResultToken*& aToken)
{
	// Release any potential content the token may hold
	aToken->Free();
	// Free the token itself.
	delete aToken;
	// Clear caller's variable.
	aToken = NULL;
}



bool Line::CatchThis(ExprTokenType &aThrown) // ACT_CATCH
{
	auto args = (CatchStatementArgs *)mAttribute;
	if (!args || !args->prototype_count)
		return Object::HasBase(aThrown, ErrorPrototype::Error);
	for (int i = 0; i < args->prototype_count; ++i)
		if (Object::HasBase(aThrown, args->prototype[i]))
			return true;
	return false;
}



void Script::ScriptWarning(WarnMode warnMode, LPCTSTR aWarningText, LPCTSTR aExtraInfo, Line *line)
{
	if (warnMode == WARNMODE_OFF)
		return;

	if (!line) line = mCurrLine;
	int fileIndex = line ? line->mFileIndex : mCurrFileIndex;
	FileIndexType lineNumber = line ? line->mLineNumber : mCombinedLineNumber;

	TCHAR buf[MSGBOX_TEXT_SIZE], *cp = buf;
	int buf_space_remaining = (int)_countof(buf);
	
	#define STD_WARNING_FORMAT _T("%s (%d) : ==> Warning: %s\n")
	cp += sntprintf(cp, buf_space_remaining, STD_WARNING_FORMAT, Line::sSourceFile[fileIndex], lineNumber, aWarningText);
	buf_space_remaining = (int)(_countof(buf) - (cp - buf));

	if (*aExtraInfo)
	{
		cp += sntprintf(cp, buf_space_remaining, _T("     Specifically: %s\n"), aExtraInfo);
		buf_space_remaining = (int)(_countof(buf) - (cp - buf));
	}

	if (warnMode == WARNMODE_STDOUT)
		PrintErrorStdOut(buf);
	else
#ifdef CONFIG_DEBUGGER
	if (!g_Debugger.OutputStdErr(buf))
#endif
		OutputDebugString(buf);

	// In MsgBox mode, MsgBox is in addition to OutputDebug
	if (warnMode == WARNMODE_MSGBOX)
	{
		g_script.RuntimeError(aWarningText, aExtraInfo, WARN, line);
	}
}



void Script::WarnUnassignedVar(Var *var, Line *aLine)
{
	auto warnMode = g_Warn_VarUnset;
	if (!warnMode)
		return;

	// Currently only the first reference to each var generates a warning even when using
	// StdOut/OutputDebug, since MarkAlreadyWarned() is used to suppress warnings for any
	// var which is checked with IsSet().
	//if (warnMode == WARNMODE_MSGBOX)
	{
		// The following check uses a flag separate to IsAssignedSomewhere() because setting
		// that one for the purpose of preventing multiple MsgBoxes would cause other callers
		// of IsAssignedSomewhere() to get the wrong result if a MsgBox has been shown.
		if (var->HasAlreadyWarned())
			return;
		var->MarkAlreadyWarned();
	}

	bool isUndeclaredLocal = (var->Scope() & (VAR_LOCAL | VAR_DECLARED)) == VAR_LOCAL;
	LPCTSTR sameNameAsGlobal = isUndeclaredLocal && FindGlobalVar(var->mName) ? _T("  (same name as a global)") : _T("");
	TCHAR buf[DIALOG_TITLE_SIZE];
	sntprintf(buf, _countof(buf), _T("%s %s%s"), Var::DeclarationType(var->Scope()), var->mName, sameNameAsGlobal);
	ScriptWarning(warnMode, WARNING_ALWAYS_UNSET_VARIABLE, buf, aLine);
}



ResultType Script::VarUnsetError(Var *var)
{
	bool isUndeclaredLocal = (var->Scope() & (VAR_LOCAL | VAR_DECLARED)) == VAR_LOCAL;
	TCHAR buf[DIALOG_TITLE_SIZE];
	if (*var->mName) // Avoid showing "Specifically: global " for temporary VarRefs of unspecified scope, such as those used by Array::FromEnumerable or the debugger.
	{
		LPCTSTR sameNameAsGlobal = isUndeclaredLocal && FindGlobalVar(var->mName) ? _T("  (same name as a global)") : _T("");
		sntprintf(buf, _countof(buf), _T("%s %s%s"), Var::DeclarationType(var->Scope()), var->mName, sameNameAsGlobal);
	}
	else *buf = '\0';
	return RuntimeError(ERR_VAR_UNSET, buf, FAIL_OR_OK, nullptr, ErrorPrototype::Unset);
}



void Script::WarnLocalSameAsGlobal(LPCTSTR aVarName)
// Relies on the following pre-conditions:
//  1) It is an implicit (not declared) variable.
//  2) Caller has verified that a global variable exists with the same name.
//  3) g_Warn_LocalSameAsGlobal is on (true).
//  4) g->CurrentFunc is the function which contains this variable.
{
	auto func_name = g->CurrentFunc ? g->CurrentFunc->mName : _T("");
	TCHAR buf[DIALOG_TITLE_SIZE];
	sntprintf(buf, _countof(buf), _T("%s  (in function %s)"), aVarName, func_name);
	ScriptWarning(g_Warn_LocalSameAsGlobal, WARNING_LOCAL_SAME_AS_GLOBAL, buf);
}
