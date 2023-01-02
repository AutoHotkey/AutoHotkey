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
#include "abi.h"
#include <richedit.h>


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

	if (!aLine)
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

		return aLine->SetThrownToken(*g, token, aErrorType);
	}

	// Returning FAIL causes each caller to also return FAIL, until either the
	// thread has fully exited or the recursion layer handling ACT_TRY is reached:
	return FAIL;
}

ResultType Script::ThrowRuntimeException(LPCTSTR aErrorText, LPCTSTR aExtraInfo)
{
	return ThrowRuntimeException(aErrorText, aExtraInfo, mCurrLine, FAIL);
}


ResultType Line::SetThrownToken(global_struct &g, ResultToken *aToken, ResultType aErrorType)
{
#ifdef CONFIG_DEBUGGER
	if (g_Debugger.IsConnected())
		if (g_Debugger.PreThrow(aToken) && !(g.ExcptMode & EXCPTMODE_CATCH))
			// The debugger has entered (and left) a break state, so the client has had a
			// chance to inspect the exception and report it.  There's nothing in the DBGp
			// spec about what to do next, probably since PHP would just log the error.
			// In our case, it seems more useful to suppress the dialog than to show it.
			return FAIL;
#endif
	g.ThrownToken = aToken;
	if (!(g.ExcptMode & EXCPTMODE_CATCH))
		return g_script.UnhandledException(this, aErrorType); // Usually returns FAIL; may return OK if aErrorType == FAIL_OR_OK.
	return FAIL;
}


ResultType Script::Win32Error(DWORD aError, ResultType aErrorType)
{
	TCHAR number_string[_MAX_ULTOSTR_BASE10_COUNT];
	// Convert aError to string to pass it through RuntimeError, but it will ultimately
	// be converted to the error number and proper message by OSError.Prototype.__New.
	_ultot(aError, number_string, 10);
	return RuntimeError(number_string, _T(""), aErrorType, nullptr, ErrorPrototype::OS);
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

int FormatStdErr(LPTSTR aBuf, int aBufSize, LPCTSTR aErrorText, LPCTSTR aExtraInfo, FileIndexType aFileIndex, LineNumberType aLineNumber, bool aWarn = false)
{
	#define STD_ERROR_FORMAT _T("%s (%d) : ==> %s%s\n")
	int n = sntprintf(aBuf, aBufSize, STD_ERROR_FORMAT, Line::sSourceFile[aFileIndex], aLineNumber
		, aWarn ? _T("Warning: ") : _T(""), aErrorText);
	if (*aExtraInfo)
		n += sntprintf(aBuf + n, aBufSize - n, _T("     Specifically: %s\n"), aExtraInfo);
	return n;
}

// For backward compatibility, this actually prints to stderr, not stdout.
void Script::PrintErrorStdOut(LPCTSTR aErrorText, LPCTSTR aExtraInfo, FileIndexType aFileIndex, LineNumberType aLineNumber)
{
	TCHAR buf[LINE_SIZE * 2];
	auto n = FormatStdErr(buf, _countof(buf), aErrorText, aExtraInfo, aFileIndex, aLineNumber);
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
	
#ifdef CONFIG_DLL
	if (LibNotifyProblem(aErrorText, aExtraInfo, this))
		return aErrorType;
#endif
	
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

	if ((g->ExcptMode || mOnError.Count()
#ifdef CONFIG_DEBUGGER
		|| g_Debugger.BreakOnExceptionIsEnabled()
#endif
		|| aPrototype) && aErrorType != WARN)
		return ThrowRuntimeException(aErrorText, aExtraInfo, aLine, aErrorType, aPrototype);
	
#ifdef CONFIG_DLL
	if (LibNotifyProblem(aErrorText, aExtraInfo, aLine))
		return aErrorType;
#endif
	
	return ShowError(aErrorText, aErrorType, aExtraInfo, aLine);
}

FResult FError(LPCTSTR aErrorText, LPCTSTR aExtraInfo, Object *aPrototype)
{
	return g_script.RuntimeError(aErrorText, aExtraInfo, FAIL_OR_OK, nullptr, aPrototype) ? FR_ABORTED : FR_FAIL;
}


struct ErrorBoxParam
{
	LPCTSTR text;
	ResultType type;
	LPCTSTR info;
	Line *line;
	IObject *obj;
#ifdef CONFIG_DEBUGGER
	int stack_index;
#endif
};


#ifdef CONFIG_DEBUGGER
void InsertCallStack(HWND re, ErrorBoxParam &error)
{
	TCHAR buf[SCRIPT_STACK_BUF_SIZE], *stack = _T("");
	if (error.obj && error.obj->IsOfType(Object::sPrototype))
	{
		auto obj = static_cast<Object*>(error.obj);
		ExprTokenType stk;
		if (obj->GetOwnProp(stk, _T("Stack")) && stk.symbol == SYM_STRING)
			stack = stk.marker;
	}
	else if (error.stack_index >= 0)
	{
		GetScriptStack(stack = buf, _countof(buf), g_Debugger.mStack.mBottom + error.stack_index);
	}

	CHARFORMAT cfBold;
	cfBold.cbSize = sizeof(cfBold);
	cfBold.dwMask = CFM_BOLD | CFM_LINK;
	cfBold.dwEffects = CFE_BOLD;
	SendMessage(re, EM_SETCHARFORMAT, SCF_SELECTION, (LPARAM)&cfBold);
	SendMessage(re, EM_REPLACESEL, FALSE, (LPARAM)_T("Call stack:\n"));
	cfBold.dwEffects = 0;
	SendMessage(re, EM_SETCHARFORMAT, SCF_SELECTION, (LPARAM)&cfBold);

	if (!*stack)
		return;

	// Prevent insertion of a blank line at the end (a bit pedantic, I know):
	auto stack_end = _tcschr(stack, '\0');
	if (stack_end[-1] == '\n')
		*--stack_end = '\0';
	if (stack_end > stack && stack_end[-1] == '\r')
		*--stack_end = '\0';
	
	CHARFORMAT cfLink;
	cfLink.cbSize = sizeof(cfLink);
	cfLink.dwMask = CFM_LINK | CFM_BOLD;
	cfLink.dwEffects = CFE_LINK;
	//cfLink.crTextColor = 0xbb4d00; // Has no effect on Windows 7 or 11 (even with CFM_COLOR).

	CHARRANGE cr;
	SendMessage(re, EM_EXGETSEL, 0, (LPARAM)&cr); // This will become the start position of the stack text.
	auto start_pos = cr.cpMin;
	SendMessage(re, EM_REPLACESEL, FALSE, (LPARAM)stack);
		
	for (auto cp = stack; ; )
	{
		if (auto ext = _tcsstr(cp, _T(".ahk (")))
		{
			// Apply CFE_LINK effect (and possibly colour) to the full path.
			cr.cpMax = cr.cpMin + int(ext - cp) + 4;
			SendMessage(re, EM_EXSETSEL, 0, (LPARAM)&cr);
			SendMessage(re, EM_SETCHARFORMAT, SCF_SELECTION, (LPARAM)&cfLink);
		}
		auto cpn = _tcschr(cp, '\n');
		if (!cpn)
			break;
		cr.cpMin += int(cpn - cp);
		if (cpn == cp || cpn[-1] != '\r') // Count the \n only if \r wasn't already counted, since it seems RichEdit uses just \r internally.
			++cr.cpMin;
		cp = cpn + 1;
	}

	cr.cpMin = cr.cpMax = start_pos - 1; // Remove selection.
	SendMessage(re, EM_EXSETSEL, 0, (LPARAM)&cr);
}
#endif


void InitErrorBox(HWND hwnd, ErrorBoxParam &error)
{
	TCHAR buf[1024];

	SetWindowText(hwnd, g_script.DefaultDialogTitle());

	SetWindowLongPtr(hwnd, DWLP_USER, (LONG_PTR)&error);

	HWND re = GetDlgItem(hwnd, IDC_ERR_EDIT);

	RECT rc, rcOffset {0,0,7,7};
	SendMessage(re, EM_GETRECT, 0, (LPARAM)&rc);
	MapDialogRect(hwnd, &rcOffset);
	rc.left += rcOffset.right;
	rc.top += rcOffset.bottom;
	rc.right -= rcOffset.right;
	rc.bottom -= rcOffset.bottom;
	SendMessage(re, EM_SETRECTNP, 0, (LPARAM)&rc);

	PARAFORMAT pf;
	pf.cbSize = sizeof(pf);
	pf.dwMask = PFM_TABSTOPS;
	pf.cTabCount = 1;
	pf.rgxTabs[0] = 300;
	SendMessage(re, EM_SETPARAFORMAT, 0, (LPARAM)&pf);

	SETTEXTEX t { ST_SELECTION | ST_UNICODE, CP_UTF16 };
	SETTEXTEX t_rtf { ST_SELECTION | ST_DEFAULT, CP_UTF8 };

	CHARFORMAT2 cf;
	cf.cbSize = sizeof(cf);

	cf.dwMask = CFM_SIZE;
	cf.yHeight = 9*20;
	SendMessage(re, EM_SETCHARFORMAT, SCF_DEFAULT, (LPARAM)&cf);
	
	cf.dwMask = CFM_SIZE | CFM_COLOR;
	cf.yHeight = 10*20;
	cf.crTextColor = 0x3399;
	cf.dwEffects = 0;
	SendMessage(re, EM_SETCHARFORMAT, SCF_SELECTION, (LPARAM)&cf);
	
	sntprintf(buf, _countof(buf), _T("%s: %.500s\n\n")
		, error.type == CRITICAL_ERROR ? _T("Critical Error")
			: error.type == WARN ? _T("Warning") : _T("Error")
		, error.text);
	SendMessage(re, EM_REPLACESEL, FALSE, (LPARAM)buf);
	
	if (error.info && *error.info)
	{
		UINT suffix = _tcslen(error.info) > 80 ? 8230 : 0;
		if (error.line)
			sntprintf(buf, _countof(buf), _T("Specifically: %.80s%s\n\n"), error.info, &suffix);
		else
			sntprintf(buf, _countof(buf), _T("Text:\t%.80s%s\n"), error.info, &suffix);
		SendMessage(re, EM_REPLACESEL, FALSE, (LPARAM)buf);
	}

	if (error.line)
	{
		#define LINES_ABOVE_AND_BELOW 2

		// Determine the range of lines to be shown:
		Line *line_start = error.line, *line_end = error.line;
		if (g_AllowMainWindow)
		{
			for (int i = 0
				; i < LINES_ABOVE_AND_BELOW && line_start->mPrevLine != NULL
				; ++i, line_start = line_start->mPrevLine);
			for (int i = 0
				; i < LINES_ABOVE_AND_BELOW && line_end->mNextLine != NULL
				; ++i, line_end = line_end->mNextLine);
		}
		//else show only a single line, to conceal the script's source code.

		int last_file = 0; // Init to zero so path is omitted if it is the main file.
		for (auto line = line_start; ; line = line->mNextLine)
		{
			if (last_file != line->mFileIndex)
			{
				last_file = line->mFileIndex;
				sntprintf(buf, _countof(buf), _T("\t---- %s\n"), Line::sSourceFile[line->mFileIndex]);
				SendMessage(re, EM_REPLACESEL, FALSE, (LPARAM)buf);
			}
			int lead = 0;
			if (line == error.line)
			{
				cf.dwMask = CFM_COLOR | CFM_BACKCOLOR;
				cf.crTextColor = 0; // Use explicit black to ensure visibility if a high contrast theme is enabled.
				cf.crBackColor = 0x60ffff;
				SendMessage(re, EM_SETCHARFORMAT, SCF_SELECTION, (LPARAM)&cf);
				buf[lead++] = 9654; // ▶
			}
			buf[lead++] = '\t';
			buf[lead] = '\0';
			SendMessage(re, EM_REPLACESEL, FALSE, (LPARAM)buf);

			line->ToText(buf, _countof(buf), true);
			SendMessage(re, EM_REPLACESEL, FALSE, (LPARAM)buf);
			if (line == line_end)
				break;
		}
	}
	else
	{
		sntprintf(buf, _countof(buf), _T("Line:\t%d\nFile:\t"), g_script.CurrentLine());
		SendMessage(re, EM_REPLACESEL, FALSE, (LPARAM)buf);
		cf.dwMask = CFM_LINK;
		cf.dwEffects = CFE_LINK; // Mark it as a link.
		SendMessage(re, EM_SETCHARFORMAT, SCF_SELECTION, (LPARAM)&cf);
		SendMessage(re, EM_REPLACESEL, FALSE, (LPARAM)g_script.CurrentFile());
		SendMessage(re, EM_REPLACESEL, FALSE, (LPARAM)_T("\n\n"));
	}

	LPCTSTR footer;
	switch (error.type)
	{
	case WARN: footer = ERR_WARNING_FOOTER; break;
	case FAIL_OR_OK: footer = nullptr; break;
	case CRITICAL_ERROR: footer = UNSTABLE_WILL_EXIT; break;
	default: footer = (g->ExcptMode & EXCPTMODE_DELETE) ? ERR_ABORT_DELETE
		: g_script.mIsReadyToExecute ? ERR_ABORT_NO_SPACES
		: g_script.mIsRestart ? OLD_STILL_IN_EFFECT
		: WILL_EXIT;
	}
	if (footer)
	{
		if (error.line)
			SendMessage(re, EM_REPLACESEL, FALSE, (LPARAM)_T("\n"));
		SendMessage(re, EM_REPLACESEL, FALSE, (LPARAM)footer);
	}

#ifdef CONFIG_DEBUGGER
	ExprTokenType tk;
	if (   error.stack_index >= 0
		|| error.obj && error.obj->IsOfType(Object::sPrototype)
			&& static_cast<Object*>(error.obj)->GetOwnProp(tk, _T("Stack"))
			&& !TokenIsEmptyString(tk)   )
	{
		// Stack trace appears to be available, so add a link to show it.
		CHARRANGE cr;
		for (int i = footer ? 2 : 1; i; --i)
			SendMessage(re, EM_REPLACESEL, FALSE, (LPARAM)_T("\n"));
		SendMessage(re, EM_EXGETSEL, 0, (LPARAM)&cr);
#define SHOW_CALL_STACK_TEXT _T("Show call stack »")
		SendMessage(re, EM_REPLACESEL, FALSE, (LPARAM)SHOW_CALL_STACK_TEXT);
		cr.cpMax = -1; // Select to end.
		SendMessage(re, EM_EXSETSEL, 0, (LPARAM)&cr);
		cf.dwMask = CFM_LINK;
		cf.dwEffects = CFE_LINK; // Mark it as a link.
		SendMessage(re, EM_SETCHARFORMAT, SCF_SELECTION, (LPARAM)&cf);
		cr.cpMin = -1; // Deselect (move selection anchor to insertion point).
		SendMessage(re, EM_EXSETSEL, 0, (LPARAM)&cr);
	}
#endif

	SendMessage(re, EM_SETEVENTMASK, 0, ENM_REQUESTRESIZE | ENM_LINK);
	SendMessage(re, EM_REQUESTRESIZE, 0, 0);

#ifndef AUTOHOTKEYSC
	if (error.line && error.line->mFileIndex ? *Line::sSourceFile[error.line->mFileIndex] == '*'
		: g_script.mKind != Script::ScriptKindFile)
		// Source "file" is an embedded resource or stdin, so can't be edited.
		EnableWindow(GetDlgItem(hwnd, ID_FILE_EDITSCRIPT), FALSE);
#endif

	if (error.type != FAIL_OR_OK)
	{
		HWND hide = GetDlgItem(hwnd, error.type == WARN ? IDCANCEL : IDCONTINUE);
		if (error.type == WARN)
		{
			// Hide "Abort" since it it's not applicable to warnings except as an alias of ExitApp,
			// shift "Continue" to the right for aesthetic purposes and make it the default button
			// (otherwise the left-most button would become the default).
			RECT rc;
			GetClientRect(hide, &rc);
			MapWindowPoints(hide, hwnd, (LPPOINT)&rc, 1);
			HWND keep = GetDlgItem(hwnd, IDCONTINUE);
			MoveWindow(keep, rc.left, rc.top, rc.right, rc.bottom, FALSE);
			DefDlgProc(hwnd, DM_SETDEFID, IDCONTINUE, 0);
			SetFocus(keep);
		}
		ShowWindow(hide, SW_HIDE);
	}
}


INT_PTR CALLBACK ErrorBoxProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	switch (msg)
	{
	case WM_COMMAND:
		switch (LOWORD(wParam))
		{
		case IDCONTINUE:
		case IDCANCEL:
			EndDialog(hwnd, wParam);
			return TRUE;
		case ID_FILE_EDITSCRIPT:
		{
			auto &error = *(ErrorBoxParam*)GetWindowLongPtr(hwnd, DWLP_USER);
			if (error.line)
			{
				g_script.Edit(Line::sSourceFile[error.line->mFileIndex]);
				return TRUE;
			}
		}
		default:
			if (LOWORD(wParam) >= ID_FILE_RELOADSCRIPT)
			{
				// Call the handler directly since g_hWnd might be NULL if this is a warning dialog.
				HandleMenuItem(NULL, LOWORD(wParam), NULL);
				return TRUE;
			}
		}
		break;
	case WM_NOTIFY:
		if (wParam == IDC_ERR_EDIT)
		{
			HWND re = ((NMHDR*)lParam)->hwndFrom;
			switch (((NMHDR*)lParam)->code)
			{
			case EN_REQUESTRESIZE: // Received when the RichEdit's content grows beyond its capacity to display all at once.
			{
				RECT &rcNew = ((REQRESIZE*)lParam)->rc;
				RECT rcOld, rcInner;
				GetWindowRect(re, &rcOld);
				SendMessage(re, EM_GETRECT, 0, (LPARAM)&rcInner);
				// Stack traces can get quite "tall" if the paths/lines are mostly short,
				// so impose a rough limit to ensure the dialog remains usable.
				int rough_limit = GetSystemMetrics(SM_CYSCREEN) * 3 / 4;
				if (rcNew.bottom > rough_limit)
					rcNew.bottom = rough_limit;
				int delta = rcNew.bottom - (rcInner.bottom - rcInner.top);
				if (rcNew.bottom == rough_limit)
					SendMessage(re, EM_SHOWSCROLLBAR, SB_VERT, TRUE);
				// Enable horizontal scroll bars if necessary.
				if (rcNew.right > (rcInner.right - rcInner.left)
					&& !(GetWindowLong(re, GWL_STYLE) & WS_HSCROLL))
				{
					SendMessage(re, EM_SHOWSCROLLBAR, SB_HORZ, TRUE);
					delta += GetSystemMetrics(SM_CYHSCROLL);
				}
				// Move the buttons (and the RichEdit, temporarily).
				ScrollWindow(hwnd, 0, delta, NULL, NULL);
				// Resize the RichEdit, while also moving it back to the origin.
				MoveWindow(re, 0, 0, rcOld.right - rcOld.left, rcOld.bottom - rcOld.top + delta, TRUE);
				// Adjust the dialog's height and vertical position.
				GetWindowRect(hwnd, &rcOld);
				MoveWindow(hwnd, rcOld.left, rcOld.top - (delta / 2), rcOld.right - rcOld.left, rcOld.bottom - rcOld.top + delta, TRUE);
				break;
			}
			case EN_LINK: // Received when the user clicks or moves the mouse over text with the CFE_LINK effect.
				if (((ENLINK*)lParam)->msg == WM_LBUTTONUP)
				{
					TEXTRANGE tr { ((ENLINK*)lParam)->chrg };
					SendMessage(re, EM_EXSETSEL, 0, (LPARAM)&tr.chrg);
					tr.lpstrText = (LPTSTR)talloca(tr.chrg.cpMax - tr.chrg.cpMin + 1);
					*tr.lpstrText = '\0';
					SendMessage(re, EM_GETTEXTRANGE, 0, (LPARAM)&tr);
					PostMessage(hwnd, WM_NEXTDLGCTL, TRUE, FALSE); // Make it less edit-like by shifting focus away.
#ifdef CONFIG_DEBUGGER
					if (!_tcscmp(tr.lpstrText, SHOW_CALL_STACK_TEXT))
					{
						auto &error = *(ErrorBoxParam*)GetWindowLongPtr(hwnd, DWLP_USER);
						InsertCallStack(re, error);
						return TRUE;
					}
#endif
					g_script.Edit(tr.lpstrText);
					return TRUE;
				}
				break;
			}
		}
		break;
	case WM_INITDIALOG:
		InitErrorBox(hwnd, *(ErrorBoxParam *)lParam);
		return FALSE; // "return FALSE to prevent the system from setting the default keyboard focus"
	}
	return FALSE;
}


ResultType Script::ShowError(LPCTSTR aErrorText, ResultType aErrorType, LPCTSTR aExtraInfo, Line *aLine, IObject *aException)
{
	if (!aErrorText)
		aErrorText = _T("");
	if (!aExtraInfo)
		aExtraInfo = _T("");
	if (!aLine)
		aLine = mCurrLine;

#ifdef CONFIG_DEBUGGER
	if (g_Debugger.HasStdErrHook())
	{
		TCHAR buf[LINE_SIZE * 2];
		FormatStdErr(buf, _countof(buf), aErrorText, aExtraInfo
			, aLine ? aLine->mFileIndex : mCurrFileIndex
			, aLine ? aLine->mLineNumber : mCombinedLineNumber
			, aErrorType == WARN);
		g_Debugger.OutputStdErr(buf);
	}
#endif

	static auto sMod = LoadLibrary(_T("riched20.dll")); // RichEdit20W
	//static auto sMod = LoadLibrary(_T("msftedit.dll")); // MSFTEDIT_CLASS (RICHEDIT50W)
	ErrorBoxParam error;
	error.text = aErrorText;
	error.type = aErrorType;
	error.info = aExtraInfo;
	error.line = aLine;
	error.obj = aException;
#ifdef CONFIG_DEBUGGER
	error.stack_index = (aException || !g_script.mIsReadyToExecute) ? -1 : int(g_Debugger.mStack.mTop - g_Debugger.mStack.mBottom);
#endif
	INT_PTR result = DialogBoxParam(NULL, MAKEINTRESOURCE(IDD_ERRORBOX), NULL, ErrorBoxProc, (LPARAM)&error);
	if (result == IDCONTINUE && aErrorType == FAIL_OR_OK)
		return OK;
	if (result == -1) // May have failed to show the custom dialog box.
		MsgBox(aErrorText, MB_TOPMOST); // Keep it simple since it will hopefully never be needed.

	if (aErrorType == CRITICAL_ERROR && mIsReadyToExecute)
		ExitApp(EXIT_CRITICAL); // Pass EXIT_CRITICAL to ensure the program always exits, regardless of OnExit.
	if (aErrorType == WARN && result == IDCANCEL && !mIsReadyToExecute) // Let Escape cancel loading the script.
		ExitApp(EXIT_EXIT);

	return FAIL; // Some callers rely on a FAIL result to propagate failure.
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
	
#ifdef CONFIG_DLL
	if (LibNotifyProblem(aErrorText, aExtraInfo, nullptr))
		return FAIL;
#endif
	
	if (g_script.mErrorStdOut && !g_script.mIsReadyToExecute) // i.e. runtime errors are always displayed via dialog.
	{
		// See LineError() for details.
		PrintErrorStdOut(aErrorText, aExtraInfo, mCurrFileIndex, mCombinedLineNumber);
	}
	else
	{
		ShowError(aErrorText, FAIL, aExtraInfo, nullptr);
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
	return Fail(g_script.RuntimeError(aErrorText, aExtraInfo, FAIL_OR_OK, nullptr, aPrototype));
}

__declspec(noinline)
ResultType ResultToken::Error(LPCTSTR aErrorText, ExprTokenType &aExtraInfo, Object *aPrototype)
{
	TCHAR buf[MAX_NUMBER_SIZE];
	return Error(aErrorText, TokenToString(aExtraInfo, buf), aPrototype);
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
FResult FValueError(LPCTSTR aErrorText, LPCTSTR aExtraInfo)
{
	return FError(aErrorText, aExtraInfo, ErrorPrototype::Value);
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


void TokenTypeAndValue(ExprTokenType &aToken, LPCTSTR &aType, LPCTSTR &aValue, TCHAR *aNBuf)
{
	if (aToken.symbol == SYM_VAR && aToken.var->IsUninitializedNormalVar())
		aType = _T("unset variable"), aValue = aToken.var->mName;
	else if (TokenIsEmptyString(aToken))
		aType = _T("empty string"), aValue = _T("");
	else
		aType = TokenTypeString(aToken), aValue = TokenToString(aToken, aNBuf);
}


__declspec(noinline)
ResultType TypeError(LPCTSTR aExpectedType, ExprTokenType &aActualValue)
{
	TCHAR number_buf[MAX_NUMBER_SIZE];
	LPCTSTR actual_type, value_as_string;
	TokenTypeAndValue(aActualValue, actual_type, value_as_string, number_buf);
	return TypeError(aExpectedType, actual_type, value_as_string);
}

ResultType TypeError(LPCTSTR aExpectedType, LPCTSTR aActualType, LPCTSTR aExtraInfo)
{
	auto an = [](LPCTSTR thing) {
		return _tcschr(_T("aeiou"), ctolower(*thing)) ? _T("n") : _T("");
	};
	TCHAR msg[512];
	sntprintf(msg, _countof(msg), _T("Expected a%s %s but got a%s %s.")
		, an(aExpectedType), aExpectedType, an(aActualType), aActualType);
	return g_script.RuntimeError(msg, aExtraInfo, FAIL_OR_OK, nullptr, ErrorPrototype::Type);
}

ResultType ResultToken::TypeError(LPCTSTR aExpectedType, ExprTokenType &aActualValue)
{
	return Fail(::TypeError(aExpectedType, aActualValue));
}

FResult FTypeError(LPCTSTR aExpectedType, ExprTokenType &aActualValue)
{
	return TypeError(aExpectedType, aActualValue) == OK ? FR_ABORTED : FR_FAIL;
}


__declspec(noinline)
ResultType ParamError(int aIndex, ExprTokenType *aParam, LPCTSTR aExpectedType, LPCTSTR aFunction)
{
	auto an = [](LPCTSTR thing) {
		return _tcschr(_T("aeiou"), ctolower(*thing)) ? _T("n") : _T("");
	};
	TCHAR msg[512];
	TCHAR number_buf[MAX_NUMBER_SIZE];
	LPCTSTR actual_type, value_as_string;
#ifdef CONFIG_DEBUGGER
	if (!aFunction)
		aFunction = g_Debugger.WhatThrew();
#endif
	if (!aParam || aParam->symbol == SYM_MISSING)
	{
#ifdef CONFIG_DEBUGGER
		sntprintf(msg, _countof(msg), _T("Parameter #%i of %s must not be omitted in this case.")
			, aIndex + 1, aFunction);
#else
		sntprintf(msg, _countof(msg), _T("Parameter #%i must not be omitted in this case.")
			, aIndex + 1);
#endif
		return g_script.RuntimeError(msg, nullptr, FAIL_OR_OK, nullptr, ErrorPrototype::Value);
	}
	TokenTypeAndValue(*aParam, actual_type, value_as_string, number_buf);
	if (!*value_as_string && !aExpectedType)
		value_as_string = actual_type;
#ifdef CONFIG_DEBUGGER
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
	return g_script.RuntimeError(msg, value_as_string, FAIL_OR_OK, nullptr
		, aExpectedType ? ErrorPrototype::Type : ErrorPrototype::Value);
}

__declspec(noinline)
ResultType ResultToken::ParamError(int aIndex, ExprTokenType *aParam)
{
	return Fail(::ParamError(aIndex, aParam, nullptr, nullptr));
}

__declspec(noinline)
ResultType ResultToken::ParamError(int aIndex, ExprTokenType *aParam, LPCTSTR aExpectedType)
{
	return Fail(::ParamError(aIndex, aParam, aExpectedType, nullptr));
}

__declspec(noinline)
ResultType ResultToken::ParamError(int aIndex, ExprTokenType *aParam, LPCTSTR aExpectedType, LPCTSTR aFunction)
{
	return Fail(::ParamError(aIndex, aParam, aExpectedType, aFunction));
}

FResult FParamError(int aIndex, ExprTokenType *aParam, LPCTSTR aExpectedType)
{
	return ::ParamError(aIndex, aParam, aExpectedType, nullptr) == OK ? FR_ABORTED : FR_FAIL;
}



ResultType FResultToError(ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount, FResult aResult, int aFirstParam)
{
	if (aResult & FR_OUR_FLAG)
	{
		if (aResult & FR_INT_FLAG)
		{
			// This is a bit of a hack and should probably be revised.
			TCHAR buf[12];
			return aResultToken.Error(ERR_FAILED, _itot(FR_GET_THROWN_INT(aResult), buf, 10));
		}
		auto code = HRESULT_CODE(aResult);
		switch (HRESULT_FACILITY(aResult))
		{
		case FR_FACILITY_CONTROL:
			ASSERT(!code);
			return aResultToken.SetExitResult(FAIL);
		case FR_FACILITY_ARG:
			return aResultToken.ParamError(code, code + aFirstParam < aParamCount ? aParam[code + aFirstParam] : nullptr);
		case FACILITY_WIN32:
			if (!code)
				code = GetLastError();
			return aResultToken.Win32Error(code);
#ifndef _DEBUG
		default: // Using a default case may slightly reduce code size.
#endif
		case FR_FACILITY_ERR:
			switch (code)
			{
			case HRESULT_CODE(FR_E_OUTOFMEM):
				return aResultToken.MemoryError();
			}
		}
		ASSERT(aResult == FR_E_FAILED); // Alert for any unhandled error codes in debug mode.
		return aResultToken.Error(ERR_FAILED);
	}
	else // Presumably a HRESULT error value.
	{
		return aResultToken.Win32Error(aResult);
	}
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
	
#ifdef CONFIG_DLL
	if (LibNotifyProblem(*token))
	{
		g.ThrownToken = token; // See comments above.
		return FAIL;
	}
#endif

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

	if (ShowError(message, aErrorType, extra, aLine, TokenToObject(*token)) == OK)
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
	if (!line) line = mCurrLine;
	int fileIndex = line ? line->mFileIndex : mCurrFileIndex;
	FileIndexType lineNumber = line ? line->mLineNumber : mCombinedLineNumber;
	
#ifdef CONFIG_DLL
	if (LibNotifyProblem(aWarningText, aExtraInfo, line, true))
		return;
#endif
	
	if (warnMode == WARNMODE_OFF)
		return;

	TCHAR buf[MSGBOX_TEXT_SIZE];
	auto n = FormatStdErr(buf, _countof(buf), aWarningText, aExtraInfo, fileIndex, lineNumber, true);

	if (warnMode == WARNMODE_STDOUT)
		PrintErrorStdOut(buf, n);
	else
#ifdef CONFIG_DEBUGGER
	if (!g_Debugger.OutputStdErr(buf))
#endif
		OutputDebugString(buf);

	// In MsgBox mode, MsgBox is in addition to OutputDebug
	if (warnMode == WARNMODE_MSGBOX)
	{
		g_script.ShowError(aWarningText, WARN, aExtraInfo, line);
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
