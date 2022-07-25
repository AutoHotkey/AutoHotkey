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

#include "stdafx.h" // pre-compiled headers
#include "script.h"
#include "globaldata.h" // for a lot of things
#include "util.h" // for strlcpy() etc.
#include "window.h" // for a lot of things
#include "application.h" // for MsgSleep()
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

	TCHAR buf[MSGBOX_TEXT_SIZE];
	FormatError(buf, _countof(buf), aErrorType, message, extra, aLine);

	if (MsgBox(buf, aErrorType == FAIL_OR_OK ? MB_YESNO|MB_DEFBUTTON2 : 0) == IDYES)
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



ResultType Script::PreprocessLocalVars(FuncList &aFuncs)
{
	for (int i = 0; i < aFuncs.mCount; ++i)
	{
		ASSERT(!aFuncs.mItem[i]->IsBuiltIn());
		auto &func = *(UserFunc *)aFuncs.mItem[i];
		if (!PreprocessLocalVars(func))
			return FAIL;
		// Nested functions will be preparsed next, due to the fact that they immediately
		// follow the outer function in aFuncs.
	}
	return OK;
}



ResultType Script::PreprocessLocalVars(UserFunc &aFunc)
{
	Var **vars = aFunc.mVars.mItem;
	int var_count = aFunc.mVars.mCount;
	int closure_count = 0, non_closure_count = 0;
	for (int v = 0; v < var_count; ++v)
	{
		Var &var = *vars[v];
		if (var.Type() == VAR_CONSTANT)
		{
			// Currently only nested functions create local constants, but this could be
			// a local constant or an upvar alias for a non-local constant.
			ASSERT(var.HasObject() && dynamic_cast<UserFunc *>(var.Object()));
			auto &func = *(UserFunc *)var.Object();
			if (!func.mUpVarCount)
			{
				// This function doesn't need to be a closure, so convert it to static.
				int insert_pos;
				aFunc.mStaticVars.Find(var.mName, &insert_pos);
				aFunc.mStaticVars.Insert(&var, insert_pos);
				var.Scope() |= VAR_LOCAL_STATIC;
				++non_closure_count;
				if (var.Scope() & VAR_DOWNVAR)
				{
					var.Scope() &= ~VAR_DOWNVAR;
					--aFunc.mDownVarCount;
				}
			}
			else if (!var.IsAlias()) // At this stage only upvars are aliases.
			{
				// This is a closure of aFunc and not an alias for an outer function's
				// nested function constant.
				++closure_count;
			}
		}
	}
	if (non_closure_count) // One or more static vars are to be removed from mVars.
	{
		int dst = 0;
		for (int src = 0; src < var_count; ++src)
			if (!vars[src]->IsStatic())
				vars[dst++] = vars[src];
		aFunc.mVars.mCount = var_count = dst;
	}
	if (closure_count)
	{
		aFunc.mClosure = SimpleHeap::Alloc<ClosureInfo>(closure_count);
		for (int v = 0; v < var_count; ++v)
		{
			Var &var = *vars[v];
			if (!var.IsAlias() && var.Type() == VAR_CONSTANT)
			{
				ASSERT(var.IsObject() && dynamic_cast<UserFunc *>(var.Object()));
				auto &ci = aFunc.mClosure[aFunc.mClosureCount++];
				ci.var = &var;
				ci.func = (UserFunc *)var.Object();
			}
		}
		ASSERT(aFunc.mClosureCount == closure_count);
	}
	if (aFunc.mUpVarCount)
	{
		// Upvars are vars local to an outer function which are referenced by aFunc.
		// They might be local to aFunc.mOuterFunc or one further out.  In the latter
		// case, aFunc.mOuterFunc contains a downvar which is also one of its upvars.
		aFunc.mUpVar = SimpleHeap::Alloc<Var*>(aFunc.mUpVarCount);
		aFunc.mUpVarIndex = SimpleHeap::Alloc<int>(aFunc.mUpVarCount);

		auto &outer = *aFunc.mOuterFunc; // Always non-null when aFunc.mUpVarCount > 0.
		int upvar_count = 0;
		for (int v = 0; v < var_count; ++v)
		{
			Var &var = *vars[v];
			if (!var.IsAlias()) // At this stage only upvars are aliases.
				continue;

			Var *downvar = var.GetAliasFor(); // FindUpVar() set var to be an alias of the corresponding downvar.
			int d;
			for (d = 0; ; ++d)
			{
				ASSERT(d < outer.mDownVarCount);
				if (outer.mDownVar[d] == downvar)
					break;
			}

			ASSERT(upvar_count < aFunc.mUpVarCount);
			aFunc.mUpVar[upvar_count] = &var;
			aFunc.mUpVarIndex[upvar_count] = d;
			++upvar_count;
		}
		ASSERT(upvar_count == aFunc.mUpVarCount);
	}
	if (aFunc.mDownVarCount)
	{
		// Downvars are vars local to aFunc which are referenced by a nested function.
		aFunc.mDownVar = SimpleHeap::Alloc<Var*>(aFunc.mDownVarCount);
		
		int downvar_count = 0;
		for (int v = 0; v < var_count; ++v)
		{
			if (vars[v]->Scope() & VAR_DOWNVAR)
			{
				ASSERT(downvar_count < aFunc.mDownVarCount);
				aFunc.mDownVar[downvar_count++] = vars[v];
			}
		}
		ASSERT(downvar_count == aFunc.mDownVarCount);
	}
	return OK;
}



ResultType Script::PreparseVarRefs()
{
	// AddLine() and ExpressionToPostfix() currently leave validation of output variables
	// and lvalues to this function, to reduce code size.  Search for "VarIsReadOnlyError".
	for (Line *line = mFirstLine; line; line = line->mNextLine)
	{
		switch (line->mActionType)
		{ // Establish context for FindOrAddVar:
		case ACT_BLOCK_BEGIN: if (line->mAttribute) g->CurrentFunc = (UserFunc *)line->mAttribute; break;
		case ACT_BLOCK_END: if (line->mAttribute) g->CurrentFunc = g->CurrentFunc->mOuterFunc; break;
		}
		
		mCurrLine = line; // For error-reporting.

		for (int a = 0; a < line->mArgc; ++a)
		{
			ArgStruct &arg = line->mArg[a];
			if (!arg.is_expression)
				continue;
			for (ExprTokenType *token = arg.postfix; token->symbol != SYM_INVALID; ++token)
			{
				if (token->symbol != SYM_VAR // Not a var.
					|| VARREF_IS_WRITE(token->var_usage)) // Already resolved by ExpressionToPostfix.
					continue;
				if (token->var_deref->type == DT_FUNCREF)
				{
					token->var = token->var_deref->var;
					continue;
				}
				if (  !(token->var = FindOrAddVar(token->var_deref->marker, token->var_deref->length, FINDVAR_FOR_READ))  )
					return FAIL;
				if (token->var->IsAlias()) // Upvar.
					continue;
				switch (token->var->Type())
				{
				case VAR_CONSTANT:
					// Resolve constants to their values, except in cases like IsSet(SomeClass), or when
					// the constant is a closure (which may change each time the outer function is called).
					if (!token->var->IsLocal() && VARREF_IS_READ(token->var_usage) && !token->var->IsUninitialized())
						token->var->ToToken(*token);
					continue;
				case VAR_VIRTUAL:
					if (VARREF_IS_READ(token->var_usage))
						++arg.max_alloc; // Reserve a to_free[] slot for it in ExpandExpression().
					break;
				default:
					// Suppress any VarUnset warnings for IsSet(var) so that it can be used to determine if an
					// optionally-set global variable has been set, such as a class in an optional #Include.
					// MarkAssignedSomewhere() isn't used for this because PreprocessLocalVars() hasn't been
					// called yet, and it relies on the attribute being accurate.
					if (token->var_usage == VARREF_ISSET)
						token->var->MarkAlreadyWarned();
					// It's too early to show VarUnset warnings, since not all VARREF_ISSET references have been marked.
					// The effect of IsSet(v) for suppressing the warning shouldn't be positional since it's feasible for
					// a check in one function to guard evaluation of a VARREF_READ in some other function.
				}
			}
			if (arg.type == ARG_TYPE_INPUT_VAR)
			{
				if (arg.postfix->symbol != SYM_VAR || arg.postfix->var->Type() != VAR_NORMAL)
				{
					// Can't be ARG_TYPE_INPUT_VAR after all, as VAR_VIRTUAL and VAR_CONSTANT require ExpandExpression
					// (unless it's a constant which was converted to SYM_OBJECT above).
					arg.type = ARG_TYPE_NORMAL;
				}
				else
				{
					// This arg can be optimized by avoiding ExpandExpression.
					arg.deref = (DerefType *)arg.postfix->var;
					arg.is_expression = false;
				}
			}
		}
	}
	return OK;
}



LPTSTR ListVarsHelper(LPTSTR aBuf, int aBufSize, LPTSTR aBuf_orig, VarList &aVars)
{
	for (int i = 0; i < aVars.mCount; ++i)
		if (aVars.mItem[i]->Type() == VAR_NORMAL) // Don't bother showing VAR_CONSTANT; ToText() doesn't support VAR_VIRTUAL.
			aBuf = aVars.mItem[i]->ToText(aBuf, BUF_SPACE_REMAINING, true);
	return aBuf;
}

LPTSTR Script::ListVars(LPTSTR aBuf, int aBufSize) // aBufSize should be an int to preserve negatives from caller (caller relies on this).
// aBufSize is an int so that any negative values passed in from caller are not lost.
// Translates this script's list of variables into text equivalent, putting the result
// into aBuf and returning the position in aBuf of its new string terminator.
{
	LPTSTR aBuf_orig = aBuf;
	auto current_func = g->CurrentFunc;
	if (current_func)
	{
		// This definition might help compiler string pooling by ensuring it stays the same for all usages:
		#define LIST_VARS_UNDERLINE _T("\r\n--------------------------------------------------\r\n")
		if (current_func->mVars.mCount || !current_func->mStaticVars.mCount)
		{
			aBuf += sntprintf(aBuf, BUF_SPACE_REMAINING, _T("Local Variables for %s()%s")
				, current_func->mName, LIST_VARS_UNDERLINE);
			aBuf = ListVarsHelper(aBuf, aBufSize, aBuf_orig, current_func->mVars);
		}
		do {
			if (current_func->mStaticVars.mCount)
			{
				aBuf += sntprintf(aBuf, BUF_SPACE_REMAINING, _T("%sStatic Variables for %s()%s")
					, (aBuf > aBuf_orig) ? _T("\r\n\r\n") : _T(""), current_func->mName, LIST_VARS_UNDERLINE);
				aBuf = ListVarsHelper(aBuf, aBufSize, aBuf_orig, current_func->mStaticVars);
			}
		} while (current_func = current_func->mOuterFunc);
	}
	aBuf += sntprintf(aBuf, BUF_SPACE_REMAINING, _T("%sGlobal Variables (alphabetical)%s")
		, (aBuf > aBuf_orig) ? _T("\r\n\r\n") : _T(""), LIST_VARS_UNDERLINE);
	aBuf = ListVarsHelper(aBuf, aBufSize, aBuf_orig, mVars);
	return aBuf;
}



LPTSTR Script::ListKeyHistory(LPTSTR aBuf, int aBufSize) // aBufSize should be an int to preserve negatives from caller (caller relies on this).
// aBufSize is an int so that any negative values passed in from caller are not lost.
// Translates this key history into text equivalent, putting the result
// into aBuf and returning the position in aBuf of its new string terminator.
{
	LPTSTR aBuf_orig = aBuf; // Needed for the BUF_SPACE_REMAINING macro.
	// I was initially concerned that GetWindowText() can hang if the target window is
	// hung.  But at least on newer OS's, this doesn't seem to be a problem: MSDN says
	// "If the window does not have a caption, the return value is a null string. This
	// behavior is by design. It allows applications to call GetWindowText without hanging
	// if the process that owns the target window is hung. However, if the target window
	// is hung and it belongs to the calling application, GetWindowText will hang the
	// calling application."
	HWND target_window = GetForegroundWindow();
	TCHAR win_title[100];
	if (target_window)
		GetWindowText(target_window, win_title, _countof(win_title));
	else
		*win_title = '\0';

	TCHAR timer_list[128];
	*timer_list = '\0';
	for (ScriptTimer *timer = mFirstTimer; timer != NULL; timer = timer->mNextTimer)
		if (timer->mEnabled)
			sntprintfcat(timer_list, _countof(timer_list) - 3, _T("%s "), timer->mCallback->Name()); // Allow room for "..."
	if (*timer_list)
	{
		size_t length = _tcslen(timer_list);
		if (length > (_countof(timer_list) - 5))
			tcslcpy(timer_list + length, _T("..."), _countof(timer_list) - length);
		else if (timer_list[length - 1] == ' ')
			timer_list[--length] = '\0';  // Remove the last space if there was room enough for it to have been added.
	}

	TCHAR LRtext[256];
	aBuf += sntprintf(aBuf, aBufSize,
		_T("Window: %s")
		//"\r\nBlocks: %u"
		_T("\r\nKeybd hook: %s")
		_T("\r\nMouse hook: %s")
		_T("\r\nEnabled Timers: %u of %u (%s)")
		//"\r\nInterruptible?: %s"
		_T("\r\nInterrupted threads: %d%s")
		_T("\r\nPaused threads: %d of %d (%d layers)")
		_T("\r\nModifiers (GetKeyState() now) = %s")
		_T("\r\n")
		, win_title
		//, SimpleHeap::GetBlockCount()
		, g_KeybdHook == NULL ? _T("no") : _T("yes")
		, g_MouseHook == NULL ? _T("no") : _T("yes")
		, mTimerEnabledCount, mTimerCount, timer_list
		//, INTERRUPTIBLE ? "yes" : "no"
		, g_nThreads > 1 ? g_nThreads - 1 : 0
		, g_nThreads > 1 ? _T(" (preempted: they will resume when the current thread finishes)") : _T("")
		, g_nPausedThreads - (g_array[0].IsPaused && !mAutoExecSectionIsRunning)  // Historically thread #0 isn't counted as a paused thread unless the auto-exec section is running but paused.
		, g_nThreads, g_nLayersNeedingTimer
		, ModifiersLRToText(GetModifierLRState(true), LRtext));
	GetHookStatus(aBuf, BUF_SPACE_REMAINING);
	aBuf += _tcslen(aBuf); // Adjust for what GetHookStatus() wrote to the buffer.
	return aBuf + sntprintf(aBuf, BUF_SPACE_REMAINING, g_KeyHistory ? _T("\r\nPress [F5] to refresh.")
		: _T("\r\nKey History has been disabled via KeyHistory(0)."));
}



ResultType Script::ActionExec(LPTSTR aAction, LPTSTR aParams, LPTSTR aWorkingDir, bool aDisplayErrors
	, LPTSTR aRunShowMode, HANDLE *aProcess, bool aUpdateLastError, bool aUseRunAs, Var *aOutputVar)
// Caller should specify NULL for aParams if it wants us to attempt to parse out params from
// within aAction.  Caller may specify empty string ("") instead to specify no params at all.
// Remember that aAction and aParams can both be NULL, so don't dereference without checking first.
// Note: For the Run & RunWait commands, aParams should always be NULL.  Params are parsed out of
// the aActionString at runtime, here, rather than at load-time because Run & RunWait might contain
// dereferenced variable(s), which can only be resolved at runtime.
{
	HANDLE hprocess_local;
	HANDLE &hprocess = aProcess ? *aProcess : hprocess_local; // To simplify other things.
	hprocess = NULL; // Init output param if the caller gave us memory to store it.  Even if caller didn't, other things below may rely on this being initialized.
	if (aOutputVar) // Same
		aOutputVar->Assign();

	// Launching nothing is always a success:
	if (!aAction || !*aAction) return OK;

	if (aWorkingDir)
	{
		if (*aWorkingDir)
		{
			// aWorkingDir is validated for two main reasons:
			//  1) To make the cause of failure much clearer when both CreateProcess and ShellExecuteEx fail.
			//  2) To prevent ShellExecuteEx from wrongly executing a file:
			//     a) with aAction relative to the wrong working directory; and/or
			//     b) with the child process having the wrong working directory.
			auto attrib = GetFileAttributes(aWorkingDir);
			if (attrib == INVALID_FILE_ATTRIBUTES || !(attrib & FILE_ATTRIBUTE_DIRECTORY))
				return aDisplayErrors ? RuntimeError(ERR_PARAM2_INVALID, aWorkingDir) : FAIL;
		}
		else
			// Set it to NULL because CreateProcess() won't work if it's the empty string.
			aWorkingDir = NULL;
	}

	#define IS_VERB(str) (   !_tcsicmp(str, _T("find")) || !_tcsicmp(str, _T("explore")) || !_tcsicmp(str, _T("open"))\
		|| !_tcsicmp(str, _T("Edit")) || !_tcsicmp(str, _T("print")) || !_tcsicmp(str, _T("properties"))   )

	// Set default items to be run by ShellExecute().  These are also used by the error
	// reporting at the end, which is why they're initialized even if CreateProcess() works
	// and there's no need to use ShellExecute():
	LPTSTR shell_verb = NULL;
	LPTSTR shell_action = aAction;
	LPTSTR shell_params = NULL;
	
	///////////////////////////////////////////////////////////////////////////////////
	// This next section is done prior to CreateProcess() because when aParams is NULL,
	// we need to find out whether aAction contains a system verb.
	///////////////////////////////////////////////////////////////////////////////////
	if (aParams) // Caller specified the params (even an empty string counts, for this purpose).
	{
		if (IS_VERB(shell_action))
		{
			shell_verb = shell_action;
			shell_action = aParams;
		}
		else
			shell_params = aParams;
	}
	else // Caller wants us to try to parse params out of aAction.
	{
		// Find out the "first phrase" in the string to support the special "find" and "explore" operations.
		LPTSTR phrase;
		size_t phrase_len;
		// Set phrase_end to be the location of the first whitespace char, if one exists:
		LPTSTR phrase_end = StrChrAny(shell_action, _T(" \t")); // Find space or tab.
		if (phrase_end) // i.e. there is a second phrase.
		{
			phrase_len = phrase_end - shell_action;
			// Create a null-terminated copy of the phrase for comparison.
			phrase = tmemcpy(talloca(phrase_len + 1), shell_action, phrase_len);
			phrase[phrase_len] = '\0';
			// Firstly, treat anything following '*' as a verb, to support custom verbs like *Compile.
			if (*phrase == '*')
				shell_verb = phrase + 1;
			// Secondly, check for common system verbs like "find" and "edit".
			else if (IS_VERB(phrase))
				shell_verb = phrase;
			if (shell_verb)
				// Exclude the verb and its trailing space or tab from further consideration.
				shell_action += phrase_len + 1;
			// Otherwise it's not a verb, and may be re-parsed later.
		}
		// shell_action will be split into action and params at a later stage if ShellExecuteEx is to be used.
	}

	// This is distinct from hprocess being non-NULL because the two aren't always the
	// same.  For example, if the user does "Run, find D:\" or "RunWait, www.yahoo.com",
	// no new process handle will be available even though the launch was successful:
	bool success = false; // Separate from last_error for maintainability.
	DWORD last_error = 0;

	bool use_runas = aUseRunAs && (!mRunAsUser.IsEmpty() || !mRunAsPass.IsEmpty() || !mRunAsDomain.IsEmpty());
	if (use_runas && shell_verb)
	{
		if (aDisplayErrors)
			return RuntimeError(_T("System verbs unsupported with RunAs."));
		return FAIL;
	}
	
	size_t action_length = _tcslen(shell_action); // shell_action == aAction if shell_verb == NULL.
	if (action_length >= LINE_SIZE) // Max length supported by CreateProcess() is 32 KB. But there hasn't been any demand to go above 16 KB, so seems little need to support it (plus it reduces risk of stack overflow).
	{
        if (aDisplayErrors)
			return RuntimeError(_T("String too long.")); // Short msg since so rare.
		return FAIL;
	}

	// If the caller originally gave us NULL for aParams, always try CreateProcess() before
	// trying ShellExecute().  This is because ShellExecute() is usually a lot slower.
	// The only exception is if the action appears to be a verb such as open, edit, or find.
	// In that case, we'll also skip the CreateProcess() attempt and do only the ShellExecute().
	// If the user really meant to launch find.bat or find.exe, for example, he should add
	// the extension (e.g. .exe) to differentiate "find" from "find.exe":
	if (!shell_verb)
	{
		STARTUPINFO si = {0}; // Zero fill to be safer.
		si.cb = sizeof(si);
		// The following are left at the default of NULL/0 set higher above:
		//si.lpReserved = si.lpDesktop = si.lpTitle = NULL;
		//si.lpReserved2 = NULL;
		si.dwFlags = STARTF_USESHOWWINDOW;  // This tells it to use the value of wShowWindow below.
		si.wShowWindow = (aRunShowMode && *aRunShowMode) ? Line::ConvertRunMode(aRunShowMode) : SW_SHOWNORMAL;
		PROCESS_INFORMATION pi = {0};

		// Since CreateProcess() requires that the 2nd param be modifiable, ensure that it is
		// (even if this is ANSI and not Unicode; it's just safer):
		LPTSTR command_line;
		if (aParams && *aParams)
		{
			command_line = talloca(action_length + _tcslen(aParams) + 10); // +10 to allow room for quotes, space, terminator, and any extra chars that might get added in the future.
			_stprintf(command_line, _T("\"%s\" %s"), aAction, aParams);
		}
		else // We're running the original action from caller.
		{
			command_line = talloca(action_length + 1);
        	_tcscpy(command_line, aAction); // CreateProcessW() requires modifiable string.  Although non-W version is used now, it feels safer to make it modifiable anyway.
		}

		if (use_runas)
		{
			if (!DoRunAs(command_line, aWorkingDir, aDisplayErrors, si.wShowWindow  // wShowWindow (min/max/hide).
				, aOutputVar, pi, success, hprocess, last_error)) // These are output parameters it will set for us.
				return FAIL; // It already displayed the error, if appropriate.
		}
		else
		{
			// MSDN: "If [lpCurrentDirectory] is NULL, the new process is created with the same
			// current drive and directory as the calling process." (i.e. since caller may have
			// specified a NULL aWorkingDir).  Also, we pass NULL in for the first param so that
			// it will behave the following way (hopefully under all OSes): "the first white-space delimited
			// token of the command line specifies the module name. If you are using a long file name that
			// contains a space, use quoted strings to indicate where the file name ends and the arguments
			// begin (see the explanation for the lpApplicationName parameter). If the file name does not
			// contain an extension, .exe is appended. Therefore, if the file name extension is .com,
			// this parameter must include the .com extension. If the file name ends in a period (.) with
			// no extension, or if the file name contains a path, .exe is not appended. If the file name does
			// not contain a directory path, the system searches for the executable file in the following
			// sequence...".
			// Provide the app name (first param) if possible, for greater expected reliability.
			// UPDATE: Don't provide the module name because if it's enclosed in double quotes,
			// CreateProcess() will fail, at least under XP:
			//if (CreateProcess(aParams && *aParams ? aAction : NULL
			if (CreateProcess(NULL, command_line, NULL, NULL, FALSE, 0, NULL, aWorkingDir, &si, &pi))
			{
				success = true;
				if (pi.hThread)
					CloseHandle(pi.hThread); // Required to avoid memory leak.
				hprocess = pi.hProcess;
				if (aOutputVar)
					aOutputVar->Assign(pi.dwProcessId);
			}
			else
				last_error = GetLastError();
		}
	}

	// Since CreateProcessWithLogonW() was either not attempted or did not work, it's probably
	// best to display an error rather than trying to run it without the RunAs settings.
	// This policy encourages users to have RunAs in effect only when necessary:
	if (!success && !use_runas) // Either the above wasn't attempted, or the attempt failed.  So try ShellExecute().
	{
		SHELLEXECUTEINFO sei = {0};
		// sei.hwnd is left NULL to avoid potential side-effects with having a hidden window be the parent.
		// However, doing so may result in the launched app appearing on a different monitor than the
		// script's main window appears on (for multimonitor systems).  This seems fairly inconsequential
		// since scripted workarounds are possible.
		sei.cbSize = sizeof(sei);
		// Below: "indicate that the hProcess member receives the process handle" and not to display error dialog:
		sei.fMask = SEE_MASK_NOCLOSEPROCESS | SEE_MASK_FLAG_NO_UI;
		sei.lpDirectory = aWorkingDir; // OK if NULL or blank; that will cause current dir to be used.
		sei.nShow = (aRunShowMode && *aRunShowMode) ? Line::ConvertRunMode(aRunShowMode) : SW_SHOWNORMAL;
		if (shell_verb)
		{
			sei.lpVerb = shell_verb;
			if (!_tcsicmp(shell_verb, _T("properties")))
				sei.fMask |= SEE_MASK_INVOKEIDLIST;  // Need to use this for the "properties" verb to work reliably.
		}
		if (!shell_params) // i.e. above hasn't determined the params yet.
		{
// Rather than just consider the first phrase to be the executable and the rest to be the param, we check it
// for a proper extension so that the user can launch a document name containing spaces, without having to
// enclose it in double quotes.  UPDATE: Want to be able to support executable filespecs without requiring them
// to be enclosed in double quotes.  Therefore, search the entire string, rather than just first_phrase, for
// the left-most occurrence of a valid executable extension.  This should be fine since the user can still
// pass in EXEs and such as params as long as the first executable is fully qualified with its real extension
// so that we can tell that it's the action and not one of the params.  UPDATE: Since any file type may
// potentially accept parameters (.lnk or .ahk files for instance), the first space-terminated substring which
// is either an existing file or ends in one of .exe,.bat,.com,.cmd,.hta is considered the executable and the
// rest is considered the param.  Remaining shortcomings of this method include:
//   -  It doesn't handle an extensionless executable such as "notepad test.txt"
//   -  It doesn't handle custom file types (scripts etc.) which don't exist in the working directory but can
//      still be executed due to %PATH% and %PATHEXT% even when our caller doesn't supply an absolute path.
// These limitations seem acceptable since the caller can allow even those cases to work by simply wrapping
// the action in quote marks.
			// Make a copy so that we can modify it (i.e. split it into action & params).
			// Using talloca ensures it will stick around until the function exits:
			LPTSTR parse_buf = talloca(action_length + 1);
			_tcscpy(parse_buf, shell_action);
			LPTSTR action_extension, action_end;
			// Let quotation marks be used to remove all ambiguity:
			if (*parse_buf == '"' && (action_end = _tcschr(parse_buf + 1, '"')))
			{
				shell_action = parse_buf + 1;
				*action_end = '\0';
				if (action_end[1])
				{
					shell_params = action_end + 1;
					// Omit the space which should follow, but only one, in case spaces
					// are meaningful to the target program.
					if (*shell_params == ' ')
						++shell_params;
				}
				// Otherwise, there's only the action in quotation marks and no params.
			}
			else
			{
				if (aWorkingDir) // Set current directory temporarily in case the action is a relative path:
					SetCurrentDirectory(aWorkingDir);
				// For each space which possibly delimits the action and params:
				for (action_end = parse_buf + 1; action_end = _tcschr(action_end, ' '); ++action_end)
				{
					// Find the beginning of the substring or file extension; if \ is encountered, this might be
					// an extensionless filename, but it probably wouldn't be meaningful to pass params to such a
					// file since it can't be associated with anything, so skip to the next space in that case.
					for ( action_extension = action_end - 1;
						  action_extension > parse_buf && !_tcschr(_T("\\/."), *action_extension);
						  --action_extension );
					if (*action_extension == '.') // Potential file extension; even if action_extension == parse_buf since ".ext" on its own is a valid filename.
					{
						*action_end = '\0'; // Temporarily terminate.
						// If action_extension is a common executable extension, don't call GetFileAttributes() since
						// the file might actually be in a location listed in %PATH% or the App Paths registry key:
						if ( (action_end-action_extension == 4 && tcscasestr(_T(".exe.bat.com.cmd.hta"), action_extension))
						// Otherwise the file might still be something capable of accepting params, like a script,
						// so check if what we have is the name of an existing file:
							|| !(GetFileAttributes(parse_buf) & FILE_ATTRIBUTE_DIRECTORY) ) // i.e. THE FILE EXISTS and is not a directory. This works because (INVALID_FILE_ATTRIBUTES & FILE_ATTRIBUTE_DIRECTORY) is non-zero.
						{	
							shell_action = parse_buf;
							shell_params = action_end + 1;
							break;
						}
						// What we have so far isn't an obvious executable file type or the path of an existing
						// file, so assume it isn't a valid action.  Unterminate and continue the loop:
						*action_end = ' ';
					}
				}
				if (aWorkingDir)
					SetCurrentDirectory(g_WorkingDir); // Restore to proper value.
			}
		}
		//else aParams!=NULL, so the extra parsing in the block above isn't necessary.

		// Not done because it may have been set to shell_verb above:
		//sei.lpVerb = NULL;
		sei.lpFile = shell_action;
		sei.lpParameters = shell_params; // NULL if no parameters were present.
		// Above was fixed v1.0.42.06 to be NULL rather than the empty string to prevent passing an
		// extra space at the end of a parameter list (this might happen only when launching a shortcut
		// [.lnk file]).  MSDN states: "If the lpFile member specifies a document file, lpParameters should
		// be NULL."  This implies that NULL is a suitable value for lpParameters in cases where you don't
		// want to pass any parameters at all.
		
		if (ShellExecuteEx(&sei)) // Relies on short-circuit boolean order.
		{
			if (hprocess = sei.hProcess)
			{
				// A new process was created, so get its ID if possible.
				if (aOutputVar)
					aOutputVar->Assign(GetProcessId(hprocess));
			}
			// Even if there's no process handle, it's considered a success because some
			// system verbs and file associations do not create a new process, by design.
			success = true;
		}
		else
			last_error = GetLastError();
	}

	if (!success) // The above attempt(s) to launch failed.
	{
		if (aUpdateLastError)
			g->LastError = last_error;

		if (aDisplayErrors)
		{
			TCHAR error_text[2048], verb_text[128], system_error_text[512];
			GetWin32ErrorText(system_error_text, _countof(system_error_text), last_error);
			if (shell_verb)
				sntprintf(verb_text, _countof(verb_text), _T("\nVerb: <%s>"), shell_verb);
			else // Don't bother showing it if it's just "open".
				*verb_text = '\0';
			if (!shell_params)
				shell_params = _T(""); // Expected to be non-NULL below.
			// Use format specifier to make sure it doesn't get too big for the error
			// function to display:
			sntprintf(error_text, _countof(error_text)
				, _T("%s\nAction: <%-0.400s%s>")
				_T("%s")
				_T("\nParams: <%-0.400s%s>")
				, use_runas ? _T("Launch Error (possibly related to RunAs):") : _T("Failed attempt to launch program or document:")
				, shell_action, _tcslen(shell_action) > 400 ? _T("...") : _T("")
				, verb_text
				, shell_params, _tcslen(shell_params) > 400 ? _T("...") : _T("")
				);
			return RuntimeError(error_text, system_error_text);
		}
		return FAIL;
	}

	// Otherwise, success:
	if (aUpdateLastError)
		g->LastError = 0; // Force zero to indicate success, which seems more maintainable and reliable than calling GetLastError() right here.

	// If aProcess isn't NULL, the caller wanted the process handle left open and so it must eventually call
	// CloseHandle().  Otherwise, we should close the process if it's non-NULL (it can be NULL in the case of
	// launching things like "find D:\" or "www.yahoo.com").
	if (!aProcess && hprocess)
		CloseHandle(hprocess); // Required to avoid memory leak.
	return OK;
}
