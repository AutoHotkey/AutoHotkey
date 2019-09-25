/*
Debugger.cpp - Main body of AutoHotkey debugger engine.

Original code by Steve Gray.

This software is provided 'as-is', without any express or implied
warranty. In no event will the authors be held liable for any damages
arising from the use of this software.

Permission is granted to anyone to use this software for any purpose,
including commercial applications, and to alter it and redistribute it
freely, without restriction.
*/

#include "stdafx.h"

#include "defines.h"
#include "globaldata.h" // for access to many global vars
#include "script_object.h"
#include "script_com.h"
#include "TextIO.h"
//#include "Debugger.h" // included by globaldata.h

#ifdef CONFIG_DEBUGGER

// helper macro for WriteF()
#define U4T(s) CStringUTF8FromTChar(s).GetString()

#include <ws2tcpip.h>
#include <wspiapi.h> // for getaddrinfo() on versions of Windows earlier than XP.
#include <stdarg.h>
#include <typeinfo> // for typeid().

Debugger g_Debugger;
CStringA g_DebuggerHost;
CStringA g_DebuggerPort;

// The first breakpoint uses sMaxId + 1. Don't change this without also changing breakpoint_remove.
int Breakpoint::sMaxId = 0;


Debugger::CommandDef Debugger::sCommands[] =
{
	{"run", &run},
	{"step_into", &step_into},
	{"step_over", &step_over},
	{"step_out", &step_out},
	{"break", &_break},
	{"stop", &stop},
	{"detach", &detach},

	{"status", &status},
	
	{"stack_get", &stack_get},
	{"stack_depth", &stack_depth},
	{"context_get", &context_get},
	{"context_names", &context_names},

	{"property_get", &property_get},
	{"property_set", &property_set},
	{"property_value", &property_value},
	
	{"feature_get", &feature_get},
	{"feature_set", &feature_set},
	
	{"breakpoint_set", &breakpoint_set},
	{"breakpoint_get", &breakpoint_get},
	{"breakpoint_update", &breakpoint_update},
	{"breakpoint_remove", &breakpoint_remove},
	{"breakpoint_list", &breakpoint_list},

	{"stdout", &redirect_stdout},
	{"stderr", &redirect_stderr},

	{"typemap_get", &typemap_get},
	
	{"source", &source},
};


// PreExecLine: aLine is about to execute; handle current line marker, breakpoints and step into/over/out.
int Debugger::PreExecLine(Line *aLine)
{
	// Using this->mCurrLine might perform a little better than the alternative, at the expense of a
	// small amount of complexity in stack_get (which is only called by request of the debugger client):
	//	mStack.mTop->line = aLine;
	mCurrLine = aLine;
	
	// Check for a breakpoint on the current line:
	Breakpoint *bp = aLine->mBreakpoint;
	if (bp && bp->state == BS_Enabled)
	{
		if (bp->temporary)
		{
			aLine->mBreakpoint = NULL;
			delete bp;
		}
		return Break();
	}

	if ((mInternalState == DIS_StepInto
		|| mInternalState == DIS_StepOver && mStack.Depth() <= mContinuationDepth
		|| mInternalState == DIS_StepOut && mStack.Depth() < mContinuationDepth) // Due to short-circuit boolean evaluation, mStack.Depth() is only evaluated once and only if mInternalState is StepOver or StepOut.
		// Although IF/ELSE/LOOP skips its block-begin, standalone/function-body block-begin still gets here; we want to skip it:
		&& aLine->mActionType != ACT_BLOCK_BEGIN && (aLine->mActionType != ACT_BLOCK_END || aLine->mAttribute) // Ignore { and }; except for function-end, since we want to break there after a "return" to inspect variables while they're still in scope.
		&& aLine->mLineNumber) // Some scripts (i.e. LowLevel/code.ahk) use mLineNumber==0 to indicate the Line has been generated and injected by the script.
	{
		return Break();
	}
	
	// Check if a command was sent asynchronously (while the script was running).
	// Such commands may also be detected via the AHK_CHECK_DEBUGGER message,
	// but if the program is checking for messages infrequently or not at all,
	// the check here is needed to ensure the debugger is responsive.
	if (HasPendingCommand())
	{
		// A command was sent asynchronously.
		return ProcessCommands();
	}
	
	return DEBUGGER_E_OK;
}


bool Debugger::HasPendingCommand()
// Returns true if there is data in the socket's receive buffer.
// This is used for receiving commands asynchronously.
{
	u_long dataPending;
	if (ioctlsocket(mSocket, FIONREAD, &dataPending) == 0)
		return dataPending > 0;
	return false;
}


int Debugger::EnterBreakState()
{
	if (mInternalState != DIS_Break)
	{
		if (mInternalState != DIS_Starting)
			// Send a response for the previous continuation command.
			if (int err = SendContinuationResponse())
				return err;
		// Remove keyboard/mouse hooks.
		if (mDisabledHooks = GetActiveHooks())
			AddRemoveHooks(0, true);
		// Set break state.
		mInternalState = DIS_Break;
	}
	//else: no continuation command has been received so send no "response".
	// Hooks were already removed and mDisabledHooks must not be overwritten.
	return DEBUGGER_E_OK;
}


void Debugger::ExitBreakState()
{
	// Restore keyboard/mouse hooks if they were previously removed.
	if (mDisabledHooks)
	{
		AddRemoveHooks(mDisabledHooks, true);
		mDisabledHooks = 0;
	}
}


int Debugger::Break()
{
	if (mInternalState == DIS_Break)
		// Already in a break state, so it's likely that we are currently evaluating a
		// DBGp command, such as when property_set releases an object which implements
		// __delete and this causes a breakpoint to be hit.  In that case we must not
		// re-enter the command loop until the current command has completed.
		return DEBUGGER_E_OK;
	int err = EnterBreakState();
	if (!err)
		err = ProcessCommands();
	return err;
}


const int MAX_DBGP_ARGS = 16; // More than currently necessary.

int Debugger::ProcessCommands()
{
	// Disable notification of READ readiness and reset socket to synchronous mode.
	u_long zero = 0;
	WSAAsyncSelect(mSocket, g_hWnd, 0, 0);
	ioctlsocket(mSocket, FIONBIO, &zero);

	int err;

	for (;;)
	{
		int command_length;

		if (err = ReceiveCommand(&command_length))
			break; // Already called FatalError().

		char *command = mCommandBuf.mData;
		char *args = strchr(command, ' ');
		char *argv[MAX_DBGP_ARGS];
		int arg_count = 0;

		// transaction_id is considered optional for the following reasons:
		//  - The spec doesn't explicitly say it is mandatory (just "standard").
		//  - We don't actually need it for anything.
		//  - Rejecting the command doesn't seem useful.
		//  - It makes manually typing DBGp commands easier.
		char *transaction_id = "";

		if (args)
		{
			// Split command name and args.
			*args++ = '\0';
			// Split args into arg vector.
			err = ParseArgs(args, argv, arg_count, transaction_id);
		}
		
		if (!err)
		{
			for (int i = 0; ; ++i)
			{
				if (i == _countof(sCommands))
				{
					err = DEBUGGER_E_UNIMPL_COMMAND;
					break;
				}
				if (!strcmp(sCommands[i].mName, command))
				{
					// EXECUTE THE DBGP COMMAND.
					err = (this->*sCommands[i].mFunc)(argv, arg_count, transaction_id);
					break;
				}
			}
		}

		if (!err)
		{
			if (mResponseBuf.mDataUsed)
				err = SendResponse();
			else
				err = SendStandardResponse(command, transaction_id);
			if (err)
				break; // Already called FatalError().
		}
		else if (err == DEBUGGER_E_CONTINUE)
		{
			ASSERT(mInternalState != DIS_Break);
			// Response will be sent when the debugger breaks again.
			err = DEBUGGER_E_OK;
		}
		else
		{
			// Clear the response buffer in case a response was partially written
			// before the error was encountered (or the command failed because the
			// response buffer is full and cannot be expanded).
			mResponseBuf.Clear();

			if (mSocket == INVALID_SOCKET) // Already disconnected; see FatalError().
				break;

			if (err = SendErrorResponse(command, transaction_id, err))
				break; // Already called FatalError().
		}
		
		// Remove this command and its args from the buffer.
		// (There may be additional commands following it.)
		if (mCommandBuf.mDataUsed) // i.e. it hasn't been cleared as a result of disconnecting.
			mCommandBuf.Remove(command_length + 1);

		// If a command is received asynchronously, the debugger does not
		// enter a break state.  In that case, we need to return after each
		// command to avoid blocking in recv().
		if (mInternalState != DIS_Break)
		{
			// As ExitBreakState() can cause re-entry into ReceiveCommand() via the message pump,
			// it is safe to call only now that the command has been removed from the buffer.
			if (mInternalState != DIS_Starting) // i.e. it hasn't already been called by Disconnect().
				ExitBreakState();
			break;
		}
	}
	ASSERT(mInternalState != DIS_Break);
	// Register for message-based notification of data arrival.  If a command
	// is received asynchronously, control will be passed back to the debugger
	// to process it.  This allows the debugger engine to respond even if the
	// script is sleeping or waiting for messages.
	if (mSocket != INVALID_SOCKET)
		WSAAsyncSelect(mSocket, g_hWnd, AHK_CHECK_DEBUGGER, FD_READ | FD_CLOSE);
	return err;
}

int Debugger::ParseArgs(char *aArgs, char **aArgV, int &aArgCount, char *&aTransactionId)
{
	aArgCount = 0;

	while (*aArgs)
	{
		if (aArgCount == MAX_DBGP_ARGS)
			return DEBUGGER_E_PARSE_ERROR;

		if (*aArgs != '-')
			return DEBUGGER_E_PARSE_ERROR;
		++aArgs;

		char arg_char = *aArgs;
		if (!arg_char)
			return DEBUGGER_E_PARSE_ERROR;
		
		if (aArgs[1] == ' ' && aArgs[2] != '-')
		{
			// Move the arg letter onto the space.
			*++aArgs = arg_char;
		}
		// Store a pointer to the arg letter, followed immediately by its value.
		aArgV[aArgCount++] = aArgs;
		
		if (arg_char == 'i')
		{
			// Handle transaction_id here to simplify many other sections.
			aTransactionId = aArgs + 1;
			--aArgCount;
		}

		if (arg_char == '-') // -- base64-encoded-data
			break;

		char *next_arg;
		if (aArgs[1] == '"')
		{
			char *arg_end;
			for (arg_end = aArgs + 1, next_arg = aArgs + 2; *next_arg != '"'; ++arg_end, ++next_arg)
			{
				if (*next_arg == '\\')
					++next_arg; // Currently only \\ and \" are supported; i.e. mark next char as literal.
				if (!*next_arg)
					return DEBUGGER_E_PARSE_ERROR;
				// Copy the value to eliminate the quotes and escape sequences.
				*arg_end = *next_arg;
			}
			*arg_end = '\0'; // Terminate this arg.
			++next_arg;
			if (!*next_arg)
				break;
			if (strncmp(next_arg, " -", 2))
				return DEBUGGER_E_PARSE_ERROR;
		}
		else
		{
			// Find where this arg's value ends and the next arg begins.
			next_arg = strstr(aArgs + 1, " -");
			if (!next_arg)
				break;
			*next_arg = '\0'; // Terminate this arg.
		}
		
		// Point aArgs to the next arg's hyphen.
		aArgs = next_arg + 1;
	}

	return DEBUGGER_E_OK;
}


//
// DBGP COMMANDS
//

// Calculate base64-encoded size of data, including NULL terminator. org_size must be > 0 if unsigned.
#define DEBUGGER_BASE64_ENCODED_SIZE(org_size) ((((org_size)-1)/3+1)*4 +1)

DEBUGGER_COMMAND(Debugger::status)
{
	if (aArgCount)
		return DEBUGGER_E_INVALID_OPTIONS;

	char *status;
	switch (mInternalState)
	{
	case DIS_Starting:	status = "starting";	break;
	case DIS_Break:		status = "break";		break;
	default:			status = "running";
	}

	return mResponseBuf.WriteF(
		"<response command=\"status\" status=\"%s\" reason=\"ok\" transaction_id=\"%e\"/>"
		, status, aTransactionId);
}

DEBUGGER_COMMAND(Debugger::feature_get)
{
	// feature_get accepts exactly one arg: -n feature_name.
	if (aArgCount != 1 || ArgChar(aArgV, 0) != 'n')
		return DEBUGGER_E_INVALID_OPTIONS;

	char *feature_name = ArgValue(aArgV, 0);

	bool supported = false; // Is %feature_name% a supported feature string?
	char *setting = "";

	if (!strncmp(feature_name, "language_", 9)) {
		if (!strcmp(feature_name + 9, "supports_threads"))
			setting = "0";
		else if (!strcmp(feature_name + 9, "name"))
			setting = DEBUGGER_LANG_NAME;
		else if (!strcmp(feature_name + 9, "version"))
#ifdef UNICODE
			setting = AHK_VERSION " (Unicode)";
#else
			setting = AHK_VERSION;
#endif
	} else if (!strcmp(feature_name, "encoding"))
		setting = "UTF-8";
	else if (!strcmp(feature_name, "protocol_version")
			|| !strcmp(feature_name, "supports_async"))
		setting = "1";
	// Not supported: data_encoding - assume base64.
	// Not supported: breakpoint_languages - assume only %language_name% is supported.
	else if (!strcmp(feature_name, "breakpoint_types"))
		setting = "line";
	else if (!strcmp(feature_name, "multiple_sessions"))
		setting = "0";
	else if (!strcmp(feature_name, "max_data"))
		setting = _itoa(mMaxPropertyData, (char*)_alloca(MAX_INTEGER_SIZE), 10);
	else if (!strcmp(feature_name, "max_children"))
		setting = _ultoa(mMaxChildren, (char*)_alloca(MAX_INTEGER_SIZE), 10);
	else if (!strcmp(feature_name, "max_depth"))
		setting = _ultoa(mMaxDepth, (char*)_alloca(MAX_INTEGER_SIZE), 10);
	// TODO: STOPPING state for retrieving variable values, etc. after the script finishes, then implement supports_postmortem feature name. Requires debugger client support.
	else
	{
		for (int i = 0; i < _countof(sCommands); ++i)
		{
			if (!strcmp(sCommands[i].mName, feature_name))
			{
				supported = true;
				break;
			}
		}
	}

	if (*setting)
		supported = true;

	return mResponseBuf.WriteF(
		"<response command=\"feature_get\" feature_name=\"%e\" supported=\"%i\" transaction_id=\"%e\">%s</response>"
		, feature_name, (int)supported, aTransactionId, setting);
}

DEBUGGER_COMMAND(Debugger::feature_set)
{
	char arg, *value;

	char *feature_name = NULL, *feature_value = NULL;

	for (int i = 0; i < aArgCount; ++i)
	{
		arg = ArgChar(aArgV, i);
		value = ArgValue(aArgV, i);
		switch (arg)
		{
		case 'n': feature_name = value; break;
		case 'v': feature_value = value; break;
		default:
			return DEBUGGER_E_INVALID_OPTIONS;
		}
	}

	if (!feature_name || !feature_value)
		return DEBUGGER_E_INVALID_OPTIONS;

	bool success = false;

	// Since all supported features are positive integers:
	int ival = atoi(feature_value);
	if (ival < 0)
	{
		// Since this value is invalid, return success="0" to indicate the error.
		// Setting the feature to a negative value might cause instability elsewhere.
	}
	else if (success = !strcmp(feature_name, "max_data"))
	{
		if (ival == 0) // Although this isn't in the spec, at least one IDE relies on it.
			ival = INT_MAX; // Strictly following the spec, we should probably return 0 bytes of data.
		mMaxPropertyData = ival;
	}
	else if (success = !strcmp(feature_name, "max_children"))
		mMaxChildren = ival;
	else if (success = !strcmp(feature_name, "max_depth"))
		mMaxDepth = ival;

	return mResponseBuf.WriteF(
		"<response command=\"feature_set\" feature=\"%e\" success=\"%i\" transaction_id=\"%e\"/>"
		, feature_name, (int)success, aTransactionId);
}

DEBUGGER_COMMAND(Debugger::run)
{
	return run_step(aArgV, aArgCount, aTransactionId, "run", DIS_Run);
}

DEBUGGER_COMMAND(Debugger::step_into)
{
	return run_step(aArgV, aArgCount, aTransactionId, "step_into", DIS_StepInto);
}

DEBUGGER_COMMAND(Debugger::step_over)
{
	return run_step(aArgV, aArgCount, aTransactionId, "step_over", DIS_StepOver);
}

DEBUGGER_COMMAND(Debugger::step_out)
{
	return run_step(aArgV, aArgCount, aTransactionId, "step_out", DIS_StepOut);
}

int Debugger::run_step(char **aArgV, int aArgCount, char *aTransactionId, char *aCommandName, DebuggerInternalStateType aNewState)
{
	if (aArgCount)
		return DEBUGGER_E_INVALID_OPTIONS;
	
	if (mInternalState != DIS_Break)
		return DEBUGGER_E_COMMAND_UNAVAIL;

	mInternalState = aNewState;
	mContinuationDepth = mStack.Depth();
	mContinuationTransactionId = aTransactionId;
	
	// Response will be sent when the debugger breaks.
	return DEBUGGER_E_CONTINUE;
}

DEBUGGER_COMMAND(Debugger::_break)
{
	if (aArgCount)
		return DEBUGGER_E_INVALID_OPTIONS;
	if (int err = EnterBreakState())
		return err;
	return DEBUGGER_E_OK;
}

DEBUGGER_COMMAND(Debugger::stop)
{
	mContinuationTransactionId = aTransactionId;

	// Call g_script.TerminateApp instead of g_script.ExitApp to bypass OnExit subroutine.
	g_script.TerminateApp(EXIT_EXIT, 0); // This also causes this->Exit() to be called.
	
	// Should never be reached, but must be here to avoid a compile error:
	return DEBUGGER_E_INTERNAL_ERROR;
}

DEBUGGER_COMMAND(Debugger::detach)
{
	mContinuationTransactionId = aTransactionId; // Seems more appropriate than using the previous ID (if any).
	// User wants to stop the debugger but let the script keep running.
	Exit(EXIT_NONE, "detach"); // Anything but EXIT_ERROR.  Sends "stopped" response, then disconnects.
	return DEBUGGER_E_CONTINUE; // Response already sent.
}

DEBUGGER_COMMAND(Debugger::breakpoint_set)
{
	char arg, *value;
	
	char *type = NULL, state = BS_Enabled, *filename = NULL;
	LineNumberType lineno = 0;
	bool temporary = false;

	for (int i = 0; i < aArgCount; ++i)
	{
		arg = ArgChar(aArgV, i);
		value = ArgValue(aArgV, i);
		switch (arg)
		{
		case 't': // type = line | call | return | exception | conditional | watch
			type = value;
			break;

		case 's': // state = enabled | disabled
			if (!strcmp(value, "enabled"))
				state = BS_Enabled;
			else if (!strcmp(value, "disabled"))
				state = BS_Disabled;
			else
				return DEBUGGER_E_BREAKPOINT_STATE;
			break;

		case 'f': // filename
			filename = value;
			break;

		case 'n': // lineno
			lineno = strtoul(value, NULL, 10);
			break;

		case 'r': // temporary = 0 | 1
			temporary = (*value != '0');
			break;
			
		case 'm': // function
		case 'x': // exception
		case 'h': // hit_value
		case 'o': // hit_condition = >= | == | %
		case '-': // expression for conditional breakpoints
			// These aren't used/supported, but ignored for now.
			break;

		default:
			return DEBUGGER_E_INVALID_OPTIONS;
		}
	}

	if (!type || strcmp(type, "line")) // i.e. type != "line"
		return DEBUGGER_E_BREAKPOINT_TYPE;
	if (lineno < 1)
		return DEBUGGER_E_BREAKPOINT_INVALID;

	int file_index = 0;

	if (filename)
	{
		// Decode filename URI -> path, in-place.
		DecodeURI(filename);
		CStringTCharFromUTF8 filename_t(filename);

		// Find the specified source file.
		for (file_index = 0; file_index < Line::sSourceFileCount; ++file_index)
			if (!_tcsicmp(filename_t, Line::sSourceFile[file_index]))
				break;

		if (file_index >= Line::sSourceFileCount)
			return DEBUGGER_E_BREAKPOINT_INVALID;
	}

	Line *line = NULL, *found_line = NULL;
	// Due to the introduction of expressions in static initializers, lines aren't necessarily in
	// line number order.  First determine if any static initializers match the requested lineno.
	// If not, use the first non-static line at or following that line number.

	if (g_script.mFirstStaticLine)
		for (line = g_script.mFirstStaticLine; ; line = line->mNextLine)
		{
			if (line->mFileIndex == file_index && line->mLineNumber == lineno) // Exact match, unlike normal lines.
			{
				found_line = line;
				break;
			}
			if (line == g_script.mLastStaticLine)
				break;
		}
	if (!found_line)
		// If line is non-NULL, above has left it set to mLastStaticLine, which we want to exclude:
		for (line = line ? line->mNextLine : g_script.mFirstLine; line; line = line->mNextLine)
			if (line->mFileIndex == file_index && line->mLineNumber >= lineno)
			{
				// ACT_ELSE and ACT_BLOCK_BEGIN generally don't cause PreExecLine() to be called,
				// so any breakpoint set on one of those lines would never be hit.  Attempting to
				// set a breakpoint on one of these should act like setting a breakpoint on a line
				// which contains no code: put the breakpoint at the next line instead.
				// Without this check, setting a breakpoint on a line like "else Exit" would not work.
				if (line->mActionType == ACT_ELSE || line->mActionType == ACT_BLOCK_BEGIN)
					continue;
				// Use the first line of code at or after lineno, like Visual Studio.
				// To display the breakpoint correctly, an IDE should use breakpoint_get.
				if (!found_line || found_line->mLineNumber > line->mLineNumber)
					found_line = line;
				// Must keep searching, since class var initializers can cause lines to be listed out of order.
				//break;
			}
	if (found_line)
	{
		if (!found_line->mBreakpoint)
			found_line->mBreakpoint = new Breakpoint();
		found_line->mBreakpoint->state = state;
		found_line->mBreakpoint->temporary = temporary;

		return mResponseBuf.WriteF(
			"<response command=\"breakpoint_set\" transaction_id=\"%e\" state=\"%s\" id=\"%i\"/>"
			, aTransactionId, state ? "enabled" : "disabled", found_line->mBreakpoint->id);
	}
	// There are no lines of code beginning at or after lineno.
	return DEBUGGER_E_BREAKPOINT_INVALID;
}

int Debugger::WriteBreakpointXml(Breakpoint *aBreakpoint, Line *aLine)
{
	mResponseBuf.WriteF("<breakpoint id=\"%i\" type=\"line\" state=\"%s\" filename=\""
					, aBreakpoint->id, aBreakpoint->state ? "enabled" : "disabled");
	mResponseBuf.WriteFileURI(U4T(Line::sSourceFile[aLine->mFileIndex]));
	return mResponseBuf.WriteF("\" lineno=\"%u\"/>", aLine->mLineNumber);
}

DEBUGGER_COMMAND(Debugger::breakpoint_get)
{
	// breakpoint_get accepts exactly one arg: -d breakpoint_id.
	if (aArgCount != 1 || ArgChar(aArgV, 0) != 'd')
		return DEBUGGER_E_INVALID_OPTIONS;

	int breakpoint_id = atoi(ArgValue(aArgV, 0));

	Line *line;
	for (line = g_script.mFirstLine; line; line = line->mNextLine)
	{
		if (line->mBreakpoint && line->mBreakpoint->id == breakpoint_id)
		{
			mResponseBuf.WriteF("<response command=\"breakpoint_get\" transaction_id=\"%e\">", aTransactionId);
			WriteBreakpointXml(line->mBreakpoint, line);
			mResponseBuf.Write("</response>");

			return DEBUGGER_E_OK;
		}
	}

	return DEBUGGER_E_BREAKPOINT_NOT_FOUND;
}

DEBUGGER_COMMAND(Debugger::breakpoint_update)
{
	char arg, *value;
	
	int breakpoint_id = 0; // Breakpoint IDs begin at 1.
	LineNumberType lineno = 0;
	char state = -1;

	for (int i = 0; i < aArgCount; ++i)
	{
		arg = ArgChar(aArgV, i);
		value = ArgValue(aArgV, i);
		switch (arg)
		{
		case 'd':
			breakpoint_id = atoi(value);
			break;

		case 's':
			if (!strcmp(value, "enabled"))
				state = BS_Enabled;
			else if (!strcmp(value, "disabled"))
				state = BS_Disabled;
			else
				return DEBUGGER_E_BREAKPOINT_STATE;
			break;

		case 'n':
			lineno = strtoul(value, NULL, 10);
			break;

		case 'h': // hit_value
		case 'o': // hit_condition
			// These aren't used/supported, but ignored for now.
			break;

		default:
			return DEBUGGER_E_INVALID_OPTIONS;
		}
	}

	if (!breakpoint_id)
		return DEBUGGER_E_INVALID_OPTIONS;

	Line *line;
	for (line = g_script.mFirstLine; line; line = line->mNextLine)
	{
		Breakpoint *bp = line->mBreakpoint;

		if (bp && bp->id == breakpoint_id)
		{
			if (lineno && line->mLineNumber != lineno)
			{
				// Move the breakpoint within its current file.
				int file_index = line->mFileIndex;
				Line *old_line = line;

				for (line = g_script.mFirstLine; line; line = line->mNextLine)
				{
					if (line->mFileIndex == file_index && line->mLineNumber >= lineno)
					{
						line->mBreakpoint = bp;
						break;
					}
				}

				// If line is NULL, the line was not found.
				if (!line)
					return DEBUGGER_E_BREAKPOINT_INVALID;

				// Seems best to only remove the breakpoint from its previous line
				// once we know the breakpoint_update has succeeded.
				old_line->mBreakpoint = NULL;
			}

			if (state != -1)
				bp->state = state;

			return DEBUGGER_E_OK;
		}
	}

	return DEBUGGER_E_BREAKPOINT_NOT_FOUND;
}

DEBUGGER_COMMAND(Debugger::breakpoint_remove)
{
	// breakpoint_remove accepts exactly one arg: -d breakpoint_id.
	if (aArgCount != 1 || ArgChar(aArgV, 0) != 'd')
		return DEBUGGER_E_INVALID_OPTIONS;

	int breakpoint_id = atoi(ArgValue(aArgV, 0));

	Line *line;
	for (line = g_script.mFirstLine; line; line = line->mNextLine)
	{
		if (line->mBreakpoint && line->mBreakpoint->id == breakpoint_id)
		{
			delete line->mBreakpoint;
			line->mBreakpoint = NULL;

			return DEBUGGER_E_OK;
		}
	}

	return DEBUGGER_E_BREAKPOINT_NOT_FOUND;
}

DEBUGGER_COMMAND(Debugger::breakpoint_list)
{
	if (aArgCount)
		return DEBUGGER_E_INVALID_OPTIONS;
	
	mResponseBuf.WriteF("<response command=\"breakpoint_list\" transaction_id=\"%e\">", aTransactionId);
	
	Line *line;
	for (line = g_script.mFirstLine; line; line = line->mNextLine)
	{
		if (line->mBreakpoint)
		{
			WriteBreakpointXml(line->mBreakpoint, line);
		}
	}

	return mResponseBuf.Write("</response>");
}

DEBUGGER_COMMAND(Debugger::stack_depth)
{
	if (aArgCount)
		return DEBUGGER_E_INVALID_OPTIONS;

	return mResponseBuf.WriteF(
		"<response command=\"stack_depth\" depth=\"%i\" transaction_id=\"%e\"/>"
		, mStack.Depth(), aTransactionId);
}

DEBUGGER_COMMAND(Debugger::stack_get)
{
	int depth = -1;

	if (aArgCount)
	{
		// stack_get accepts one optional arg: -d depth.
		if (aArgCount > 1 || ArgChar(aArgV, 0) != 'd')
			return DEBUGGER_E_INVALID_OPTIONS;
		depth = atoi(ArgValue(aArgV, 0));
		if (depth < 0 || depth >= mStack.Depth())
			return DEBUGGER_E_INVALID_STACK_DEPTH;
	}

	mResponseBuf.WriteF("<response command=\"stack_get\" transaction_id=\"%e\">", aTransactionId);
	
	int level = 0;
	DbgStack::Entry *se;
	for (se = mStack.mTop; se >= mStack.mBottom; --se)
	{
		if (depth == -1 || depth == level)
		{
			Line *line;
			if (se == mStack.mTop)
			{
				ASSERT(mCurrLine); // Should always be valid.
				line = mCurrLine; // See PreExecLine() for comments.
			}
			else if (se->type == DbgStack::SE_Thread)
			{
				// !se->line implies se->type == SE_Thread.
				if (se[1].type == DbgStack::SE_UDF)
					line = se[1].udf->func->mJumpToLine;
				else if (se[1].type == DbgStack::SE_Sub)
					line = se[1].sub->mJumpToLine;
				else
					// The auto-execute thread is probably the only one that can exist without
					// a Sub or Func entry immediately above it.  As se != mStack.mTop, se->line
					// has been set to a non-NULL by DbgStack::Push().
					line = se->line;
			}
			else
			{
				line = se->line;
			}
			mResponseBuf.WriteF("<stack level=\"%i\" type=\"file\" filename=\"", level);
			mResponseBuf.WriteFileURI(U4T(Line::sSourceFile[line->mFileIndex]));
			mResponseBuf.WriteF("\" lineno=\"%u\" where=\"", line->mLineNumber);
			switch (se->type)
			{
			case DbgStack::SE_Thread:
				mResponseBuf.WriteF("%e thread", U4T(se->desc)); // %e to escape characters which desc may contain (e.g. "a & b" in hotkey name).
				break;
			case DbgStack::SE_UDF:
				mResponseBuf.WriteF("%e()", U4T(se->udf->func->mName));
				break;
			case DbgStack::SE_Sub:
				mResponseBuf.WriteF("%e sub", U4T(se->sub->mName)); // %e because label/hotkey names may contain almost anything.
				break;
			}
			mResponseBuf.Write("\"/>");
		}
		++level;
	}

	return mResponseBuf.Write("</response>");
}

DEBUGGER_COMMAND(Debugger::context_names)
{
	if (aArgCount)
		return DEBUGGER_E_INVALID_OPTIONS;

	return mResponseBuf.WriteF(
		"<response command=\"context_names\" transaction_id=\"%e\"><context name=\"Local\" id=\"0\"/><context name=\"Global\" id=\"1\"/></response>"
		, aTransactionId);
}

DEBUGGER_COMMAND(Debugger::context_get)
{
	int err = DEBUGGER_E_OK;
	char arg, *value;

	int context_id = 0, depth = 0;

	for (int i = 0; i < aArgCount; ++i)
	{
		arg = ArgChar(aArgV, i);
		value = ArgValue(aArgV, i);
		switch (arg)
		{
		case 'c':	context_id = atoi(value);	break;
		case 'd':
			depth = atoi(value);
			if (depth && (depth < 0 || depth >= mStack.Depth())) // Allow depth 0 even when stack is empty.
				return DEBUGGER_E_INVALID_STACK_DEPTH;
			break;
		default:
			return DEBUGGER_E_INVALID_OPTIONS;
		}
	}

	Var **var = NULL, **var_end = NULL; // An array of pointers-to-var.
	VarBkp *bkp = NULL, *bkp_end = NULL;
	
	// TODO: Include the lazy-var arrays for completeness. Low priority since lazy-var arrays are used only for 10001+ variables, and most conventional debugger interfaces would generally not be useful with that many variables.
	if (context_id == PC_Local)
	{
		mStack.GetLocalVars(depth, var, var_end, bkp, bkp_end);
	}
	else if (context_id == PC_Global)
	{
		var = g_script.mVar;
		var_end = var + g_script.mVarCount;
	}
	else
		return DEBUGGER_E_INVALID_CONTEXT;

	mResponseBuf.WriteF(
		"<response command=\"context_get\" context=\"%i\" transaction_id=\"%e\">"
		, context_id, aTransactionId);

	LPTSTR value_buf = NULL;
	CStringA name_buf;
	PropertyInfo prop(name_buf);
	prop.max_data = mMaxPropertyData;
	prop.pagesize = mMaxChildren;
	prop.max_depth = mMaxDepth;
	for ( ; var < var_end; ++var)
		if (  (err = GetPropertyInfo(**var, prop, value_buf))
			|| (err = WritePropertyXml(prop, (*var)->mName))  )
			break;
	for ( ; bkp < bkp_end; ++bkp)
		if (  (err = GetPropertyInfo(*bkp, prop, value_buf))
			|| (err = WritePropertyXml(prop, bkp->mVar->mName))  )
			break;
	free(value_buf);
	if (err)
		return err;

	return mResponseBuf.Write("</response>");
}

DEBUGGER_COMMAND(Debugger::typemap_get)
{
	if (aArgCount)
		return DEBUGGER_E_INVALID_OPTIONS;
	
	return mResponseBuf.WriteF(
		"<response command=\"typemap_get\" transaction_id=\"%e\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\">"
			"<map type=\"string\" name=\"string\" xsi:type=\"xsd:string\"/>"
			"<map type=\"int\" name=\"integer\" xsi:type=\"xsd:long\"/>"
			"<map type=\"float\" name=\"float\" xsi:type=\"xsd:double\"/>"
			"<map type=\"object\" name=\"object\"/>"
		"</response>"
		, aTransactionId);
}

DEBUGGER_COMMAND(Debugger::property_get)
{
	return property_get_or_value(aArgV, aArgCount, aTransactionId, true);
}

DEBUGGER_COMMAND(Debugger::property_value)
{
	return property_get_or_value(aArgV, aArgCount, aTransactionId, false);
}


int Debugger::GetPropertyInfo(Var &aVar, PropertyInfo &aProp, LPTSTR &aValueBuf)
{
	aProp.is_alias = aVar.mType == VAR_ALIAS;
	aProp.is_static = aVar.IsStatic();
	return GetPropertyValue(aVar, aProp, aValueBuf);
}

int Debugger::GetPropertyInfo(VarBkp &aBkp, PropertyInfo &aProp, LPTSTR &aValueBuf)
{
	aProp.is_static = false;
	aProp.is_builtin = false;
	if (aProp.is_alias = aBkp.mType == VAR_ALIAS)
		return GetPropertyValue(*aBkp.mAliasFor, aProp, aValueBuf);
	aProp.is_binaryclip = aBkp.mAttrib & VAR_ATTRIB_BINARY_CLIP;
	aBkp.ToToken(aProp.value);
	return DEBUGGER_E_OK;
}

int Debugger::GetPropertyValue(Var &aVar, PropertyInfo &aProp, LPTSTR &aValueBuf)
{
	aProp.is_binaryclip = aVar.IsBinaryClip();
	if (aProp.is_builtin = aVar.Type() != VAR_NORMAL)
	{
		size_t approx_size = aVar.Get() + 1;
		if (!aValueBuf || _msize(aValueBuf) < _TSIZE(approx_size))
		{
			free(aValueBuf);
			if (approx_size < MAX_PATH)
				approx_size = MAX_PATH; // Big enough for most built-in vars (avoids repeated reallocation for context_get).
			if (!(aValueBuf = tmalloc(approx_size)))
				return DEBUGGER_E_INTERNAL_ERROR;
		}
		aVar.Get(aValueBuf);
		aProp.value.SetValue(aValueBuf);
		CLOSE_CLIPBOARD_IF_OPEN; // Above may leave the clipboard open if aVar is Clipboard.
	}
	else
	{
		if (aVar.IsUninitializedNormalVar())
		{
			aProp.value.symbol = SYM_MISSING;
			aProp.value.marker = _T("");
		}
		else
			aVar.ToTokenSkipAddRef(aProp.value);
	}
	return DEBUGGER_E_OK;
}

int Debugger::GetPropertyInfo(Object::FieldType &aField, PropertyInfo &aProp)
{
	aField.ToToken(aProp.value);
	return DEBUGGER_E_OK;
}


int Debugger::WritePropertyXml(PropertyInfo &aProp, IObject *aObject)
{
	PropertyWriter pw(*this, aProp, aObject);
	// Ask the object to write out its properties:
	aObject->DebugWriteProperty(&pw, aProp.page, aProp.pagesize, aProp.max_depth);
	aProp.fullname.Truncate(pw.mNameLength);
	// For simplicity/code size, instead of requiring error handling in aObject,
	// any failure during the above sets pw.mError, which causes it to ignore
	// any further method calls.  Since we're finished now, return the error
	// code (even if it's "OK"):
	return pw.mError;
}

void Object::DebugWriteProperty(IDebugProperties *aDebugger, int aPage, int aPageSize, int aDepth)
{
	DebugCookie cookie;
	aDebugger->BeginProperty(NULL, "object", (int)mFieldCount + (mBase != NULL), cookie);

	if (aDepth)
	{
		int i = aPageSize * aPage, j = aPageSize * (aPage + 1);

		if (mBase)
		{
			// Since this object has a "base", let it count as the first field.
			if (i == 0) // i.e. this is the first page.
			{
				aDebugger->WriteProperty("<base>", ExprTokenType(mBase));
				// Now fall through and retrieve field[0] (unless aPageSize == 1).
			}
			// So 20..39 becomes 19..38 when there's a base object:
			else --i; 
			--j;
		}
		if (j > (int)mFieldCount)
			j = (int)mFieldCount;
		// For each field in the requested page...
		for ( ; i < j; ++i)
		{
			Object::FieldType &field = mFields[i];
			
			ExprTokenType key, value;
			if (i >= mKeyOffsetString) // String
				key.symbol = SYM_STRING, key.marker = field.key.s;
			else if (i >= mKeyOffsetObject)
				key.symbol = SYM_OBJECT, key.object = field.key.p;
			else
				key.symbol = SYM_INTEGER, key.value_int64 = field.key.i;
			field.ToToken(value);

			aDebugger->WriteProperty(key, value);
		}
	}

	aDebugger->EndProperty(cookie);
}

int Debugger::WritePropertyXml(PropertyInfo &aProp)
{
	char facetbuf[35]; // Alias Builtin Static ClipboardAll \0
	facetbuf[0] = '\0';
	if (aProp.is_alias)
		strcat(facetbuf, " Alias");
	if (aProp.is_builtin)
		strcat(facetbuf, " Builtin");
	if (aProp.is_static)
		strcat(facetbuf, " Static");
	if (aProp.is_binaryclip)
		strcat(facetbuf, " ClipboardAll");
	aProp.facet = facetbuf + (*facetbuf != '\0'); // Skip the leading space, if non-empty.

	char *type;
	switch (aProp.value.symbol)
	{
	case SYM_OPERAND:
	case SYM_STRING: type = "string"; break;
	case SYM_INTEGER: type = "integer"; break;
	case SYM_FLOAT: type = "float"; break;

	case SYM_OBJECT:
		// Recursively dump object.
		return WritePropertyXml(aProp, aProp.value.object);

	default:
		// Catch SYM_VAR or any invalid symbol in debug mode.  In release mode, treat as undefined
		// (the compiler can omit the SYM_MISSING check because the default branch covers it).
		ASSERT(FALSE);
	case SYM_MISSING:
		type = "undefined";
	}
	// If we fell through, value and type have been set appropriately above.
	mResponseBuf.WriteF("<property name=\"%e\" fullname=\"%e\" type=\"%s\" facet=\"%s\" children=\"0\" encoding=\"base64\" size=\""
		, aProp.name, aProp.fullname.GetString(), type, aProp.facet);
	int err;
	if (err = WritePropertyData(aProp.value, aProp.max_data))
		return err;
	return mResponseBuf.Write("</property>");
}

int Debugger::WritePropertyXml(PropertyInfo &aProp, LPTSTR aName)
{
	StringTCharToUTF8(aName, aProp.fullname);
	aProp.name = aProp.fullname;
	return WritePropertyXml(aProp);
}

void Debugger::AppendKeyName(CStringA &aNameBuf, size_t aParentNameLength, const char *aName)
{
	const char *ccp;
	for (ccp = aName; cisalnum(*ccp) || *ccp == '_'; ++ccp);
	if (!*ccp && ccp != aName) // If it got to the null-terminator, must be empty or alphanumeric.
	{
		// Since this string is purely composed of alphanumeric characters and/or underscore,
		// it doesn't need any quote marks (imitating expression syntax) or escaped characters.
		aNameBuf.AppendFormat(".%s", aName);
	}
	else
	{
		// " must be replaced with "" as in expressions to remove ambiguity.
		char c;
		int extra = 4; // 4 for [""].  Also count double-quote marks:
		for (ccp = aName; *ccp; ++ccp) extra += *ccp=='"';
		char *cp = aNameBuf.GetBufferSetLength(aParentNameLength + strlen(aName) + extra) + aParentNameLength;
		*cp++ = '[';
		*cp++ = '"';
		for (ccp = aName; c = *ccp; ++ccp)
		{
			*cp++ = c;
			if (c == '"')
				*cp++ = '"'; // i.e. replace " with ""
		}
		*cp++ = '"';
		*cp++ = ']';
		aNameBuf.ReleaseBuffer();
	}
}

int Debugger::WritePropertyData(LPCTSTR aData, size_t aDataSize, int aMaxEncodedSize)
// Accepts a "native" string, converts it to UTF-8, base64-encodes it and writes
// the end of the property's size attribute followed by the base64-encoded data.
{
	int err;
	
#ifdef UNICODE
	LPCWSTR utf16_value = aData;
	size_t total_utf16_size = aDataSize;
#else
	// ANSI mode.  For simplicity, convert the entire data to UTF-16 rather than attempting to
	// calculate how much ANSI text will produce the right number of UTF-8 bytes (since UTF-16
	// is needed as an intermediate step anyway).  This would fail if the string length exceeds
	// INT_MAX, but that would only matter if we built ANSI for x64 (and we don't).
	CStringWCharFromChar utf16_buf(aData, aDataSize);
	LPCWSTR utf16_value = utf16_buf.GetString();
	size_t total_utf16_size = utf16_buf.GetLength();
	if (!total_utf16_size && aDataSize) // Conversion failed (too large?)
		return DEBUGGER_E_INTERNAL_ERROR;
#endif
	
	// The spec says: "The IDE should not read more data than the length defined in the packet
	// header.  The IDE can determine if there is more data by using the property data length
	// information."  This has two implications:
	//  1) The size attribute should represent the total size, not the amount of data actually
	//     returned (when limited by aMaxEncodedSize).  This is more useful anyway.
	//  2) Since the data is encoded as UTF-8, the size attribute must be a UTF-8 byte count
	//     for any comparison by the IDE to give the correct result.

	// According to the spec, -m 0 should mean "unlimited".
	if (!aMaxEncodedSize)
		aMaxEncodedSize = INT_MAX;
	
	// Calculate:
	//  - the total size in terms of UTF-8 bytes (even if that exceeds INT_MAX).
	size_t total_utf8_size = 0;
	//  - the maximum number of wide chars to convert, taking aMaxEncodedSize into account.
	int utf16_size = (int)total_utf16_size;
	//  - the required buffer size for conversion, in bytes.
	int utf8_size = -1;

	for (size_t i = 0; i < total_utf16_size; ++i)
	{
		wchar_t wc = utf16_value[i];
		
		int char_size;
		if (wc <= 0x007F)
			char_size = 1;
		else if (wc <= 0x07FF)
			char_size = 2;
		else if (IS_SURROGATE_PAIR(wc, utf16_value[i+1]))
			char_size = 4;
		else
			char_size = 3;
		
		total_utf8_size += char_size;

		if (total_utf8_size > (size_t)aMaxEncodedSize)
		{
			if (utf16_size == total_utf16_size) // i.e. this is the first surplus char.
			{
				// Truncate the input; utf16_value[i] and beyond will not be converted/sent.
				utf16_size = (int)i;
				utf8_size = (int)(total_utf8_size - char_size);
			}
		}
	}
	if (utf8_size == -1) // Data was not limited by aMaxEncodedSize.
		utf8_size = (int)total_utf8_size;
	
	// Calculate maximum length of base64-encoded data.
	int space_needed = DEBUGGER_BASE64_ENCODED_SIZE(utf8_size);
	
	// Reserve enough space for the data's length, "> and encoded data.
	if (err = mResponseBuf.ExpandIfNecessary(mResponseBuf.mDataUsed + space_needed + MAX_INTEGER_LENGTH + 2))
		return err;
	
	// Complete the size attribute by writing the total size, in terms of UTF-8 bytes.
	if (err = mResponseBuf.WriteF("%u\">", total_utf8_size))
		return err;

	// Convert to UTF-8, using mResponseBuf temporarily.
	char *utf8_value = mResponseBuf.mData + mResponseBuf.mDataSize - space_needed;
	utf8_size = WideCharToMultiByte(CP_UTF8, 0, utf16_value, utf16_size, utf8_value, utf8_size, NULL, NULL);
	if (!utf8_size && utf16_size) // Conversion failed.
		return DEBUGGER_E_INTERNAL_ERROR;

	// Base64-encode and write the var data.
	return mResponseBuf.WriteEncodeBase64(utf8_value, (size_t)utf8_size, true);
}

int Debugger::WritePropertyData(ExprTokenType &aValue, int aMaxEncodedSize)
{
	LPTSTR value;
	size_t value_length;
	TCHAR number_buf[MAX_NUMBER_SIZE];

	value = TokenToString(aValue, number_buf);
	value_length = _tcslen(value);
	
	return WritePropertyData(value, value_length, aMaxEncodedSize);
}

int Debugger::ParsePropertyName(LPCSTR aFullName, int aDepth, int aVarScope, bool aVarMustExist
	, PropertySource &aResult)
{
	CStringTCharFromUTF8 name_buf(aFullName);
	LPTSTR name = name_buf.GetBuffer();
	size_t name_length;
	TCHAR c, *name_end, *src, *dst;
	Var *var = NULL;
	VarBkp *varbkp = NULL;
	SymbolType key_type;
	Object::FieldType *field;
	Object::IndexType insert_pos;
	Object *obj;

	aResult.kind = PropNone;

	name_end = StrChrAny(name, _T(".["));
	if (name_end)
	{
		c = *name_end;
		*name_end = '\0'; // Temporarily terminate.
	}
	name_length = _tcslen(name);

	// Validate name for more accurate error-reporting.
	if (name_length > MAX_VAR_NAME_LENGTH || !Var::ValidateName(name, DISPLAY_NO_ERROR))
		return DEBUGGER_E_INVALID_OPTIONS;

	if (aDepth > 0 && aVarScope != FINDVAR_GLOBAL)
	{
		Var **vars = NULL, **vars_end;
		VarBkp *bkps = NULL, *bkps_end;
		mStack.GetLocalVars(aDepth, vars, vars_end, bkps, bkps_end);
		if (bkps)
		{
			for ( ; ; ++bkps)
			{
				if (bkps == bkps_end)
				{
					// No local var at that depth, so make sure to not return the wrong local.
					aVarScope = FINDVAR_GLOBAL;
					break;
				}
				if (!_tcsicmp(bkps->mVar->mName, name))
				{
					varbkp = bkps;
					break;
				}
			}
		}
		else if (vars)
		{
			for ( ; ; ++vars)
			{
				if (vars == vars_end)
				{
					// No local var at that depth, so make sure to not return the wrong local.
					aVarScope = FINDVAR_GLOBAL;
					break;
				}
				if (!_tcsicmp((*vars)->mName, name))
				{
					var = *vars;
					break;
				}
			}
		}
	}

	// If we're allowed to create variables
	if (  !varbkp && !var
		&& (!aVarMustExist
		// or this variable doesn't exist
		|| !(var = g_script.FindVar(name, name_length, NULL, aVarScope))
			// but it is a built-in variable which hasn't been referenced yet:
			&& g_script.GetBuiltInVar(name))  )
		// Find or add the variable.
		var = g_script.FindOrAddVar(name, name_length, aVarScope);

	if (!var && !varbkp)
		return DEBUGGER_E_UNKNOWN_PROPERTY;

	if (!name_end)
	{
		// Just a variable name.
		if (var)
			aResult.var = var, aResult.kind = PropVar;
		else
			aResult.bkp = varbkp, aResult.kind = PropVarBkp;
		return DEBUGGER_E_OK;
	}
	IObject *iobj;
	if (varbkp && varbkp->mType == VAR_ALIAS)
		var = varbkp->mAliasFor;
	if (var)
		iobj = var->HasObject() ? var->Object() : NULL;
	else
		iobj = (varbkp->mAttrib & VAR_ATTRIB_OBJECT) ? varbkp->mObject : NULL;

	if (  !(obj = dynamic_cast<Object *>(iobj))  )
		return DEBUGGER_E_UNKNOWN_PROPERTY;

	// aFullName contains a '.' or '['.  Although it looks like an expression, the IDE should
	// only pass a property name which we gave it in response to a previous command, so we
	// only need to support the subset of expression syntax used by WriteObjectPropertyXml().
	for (*name_end = c; ; )
	{
		name = name_end + 1;
		if (c == '[')
		{
			if (*name == '"')
			{
				// Quoted string which may contain any character.
				// Replace "" with " in-place and find and of string:
				for (dst = src = ++name; c = *src; ++src)
				{
					if (c == '"')
					{
						// Quote mark; but is it a literal quote mark?
						if (*++src != '"')
							// Nope.
							break;
						//else above skipped the second quote mark, so fall through:
					}
					*dst++ = c;
				}
				if (*src != ']') return DEBUGGER_E_INVALID_OPTIONS;
				*dst = '\0'; // Only after the check above, since src might be == dst.
				name_end = src + 1; // Set it up for the next iteration.
				key_type = SYM_STRING;
			}
			else if (!_tcsnicmp(name, _T("Object("), 7))
			{
				// Object(n) where n is the address of a key object, as a literal signed integer.
				name += 7;
				name_end = _tcschr(name, ')');
				if (!name_end || name_end[1] != ']') return DEBUGGER_E_INVALID_OPTIONS;
				*name_end = '\0';
				name_end += 2; // Set it up for the next iteration.
				key_type = SYM_OBJECT;
			}
			else
			{
				// The only other valid form is a literal signed integer.
				name_end = _tcschr(name, ']');
				if (!name_end) return DEBUGGER_E_INVALID_OPTIONS;
				*name_end = '\0'; // Although not actually necessary for _ttoi(), seems best for maintainability.
				++name_end; // Set it up for the next iteration.
				key_type = SYM_INTEGER;
			}
			c = *name_end; // Set for the next iteration.
		}
		else if (c == '.')
		{
			// For simplicity, let this be any string terminated by '.' or '['.
			// Actual expressions require it to contain only alphanumeric chars and/or '_'.
			name_end = StrChrAny(name, _T(".[")); // This also sets it up for the next iteration.
			if (name_end)
			{
				c = *name_end; // Save this for the next iteration.
				*name_end = '\0';
			}
			else
				c = 0;
			//else there won't be a next iteration.
			key_type = IsPureNumeric(name); // SYM_INTEGER or SYM_STRING.
		}
		else
			return DEBUGGER_E_INVALID_OPTIONS;
		
		if (*name != '<' || name[-1] != '.') // Not a pseudo-property; i.e. ["<base>"] is always a key-value pair.
		{
			Object::KeyType key;
			if (key_type == SYM_STRING)
				key.s = name;
			else // SYM_INTEGER or SYM_OBJECT
				key.i = Exp32or64(_ttoi,_ttoi64)(name);
			field = obj->FindField(key_type, key, insert_pos);
		}
		else
			field = NULL;

		if (!field)
		{
			// IDE should request .<base> only if it was returned by property_get or context_get,
			// so this always means the object's base (field is always NULL).  By contrast, .base
			// and ["base"] originate either from a key-value pair or the user "inspecting" an
			// expression like `myObj.base`.  Since no field was found, assume it's the latter.
			if (!_tcsicmp(name, _T("base")) || !_tcsicmp(name - 1, _T(".<base>")))
			{
				if (!c)
				{
					// For property_set, this won't allow the base to be set (success="0").
					// That seems okay since it could only ever be set to NULL anyway.
					aResult.kind = PropValue;
					if (obj->mBase)
						aResult.value.SetValue(obj->mBase);
					else
						aResult.value.SetValue(_T(""));
					return DEBUGGER_E_OK;
				}
				if (  !(obj = dynamic_cast<Object *>(obj->mBase))  )
					return DEBUGGER_E_UNKNOWN_PROPERTY;
				continue; // Search the base object's fields.
			}
			else
				return DEBUGGER_E_UNKNOWN_PROPERTY;
		}

		if (!c)
		{
			// All done!
			aResult.kind = PropField;
			aResult.field = field;
			return DEBUGGER_E_OK;
		}

		if ( field->symbol != SYM_OBJECT || !(obj = dynamic_cast<Object *>(field->object)) )
			// No usable target object for the next iteration, therefore the property mustn't exist.
			return DEBUGGER_E_UNKNOWN_PROPERTY;

	} // infinite loop.
}


int Debugger::property_get_or_value(char **aArgV, int aArgCount, char *aTransactionId, bool aIsPropertyGet)
{
	int err;
	char arg, *value;

	char *name = NULL;
	int context_id = 0, depth = 0; // Var context and stack depth.
	CStringA name_buf;
	PropertyInfo prop(name_buf);
	prop.pagesize = mMaxChildren;
	prop.max_data = aIsPropertyGet ? mMaxPropertyData : 1024*1024*1024; // Limit property_value to 1GB by default.
	prop.max_depth = mMaxDepth; // Max property nesting depth.

	for (int i = 0; i < aArgCount; ++i)
	{
		arg = ArgChar(aArgV, i);
		value = ArgValue(aArgV, i);
		switch (arg)
		{
		// property long name
		case 'n': name = value; break;
		// context id - optional, default zero. see PropertyContextType enum.
		case 'c': context_id = atoi(value); break;
		// stack depth - optional, default zero
		case 'd':
			depth = atoi(value);
			if (depth && (depth < 0 || depth >= mStack.Depth())) // Allow depth 0 even when stack is empty.
				return DEBUGGER_E_INVALID_STACK_DEPTH;
			break;
		// max data size - optional
		case 'm': prop.max_data = atoi(value); break;
		// data page - optional
		case 'p': prop.page = atoi(value); break;

		default:
			return DEBUGGER_E_INVALID_OPTIONS;
		}
	}

	if (!name || prop.max_data < 0)
		return DEBUGGER_E_INVALID_OPTIONS;

	int always_use;
	switch (context_id)
	{
	// It seems best to allow context id zero to retrieve either a local or global,
	// rather than requiring the IDE to check each context when looking up a variable.
	//case PC_Local:	always_use = FINDVAR_LOCAL; break;
	case PC_Local:	always_use = FINDVAR_DEFAULT; break;
	case PC_Global:	always_use = FINDVAR_GLOBAL; break;
	default:
		return DEBUGGER_E_INVALID_CONTEXT;
	}

	if (err = ParsePropertyName(name, depth, always_use, true, prop))
	{
		// Var not found/invalid name.
		if (!aIsPropertyGet)
			return err;

		// NOTEPAD++ DBGP PLUGIN:
		// The DBGp plugin for Notepad++ assumes property_get will always succeed.
		// Property retrieval on mouse hover does not choose words intelligently,
		// so it will attempt to retrieve properties like ";" or " r".
		// If we respond with an <error/> instead of a <property/>, Notepad++ will
		// show an error message and then become unstable. Even after the editor
		// window is closed, notepad++.exe must be terminated forcefully.
		//
		// As a work-around (until this is resolved by the plugin's author),
		// we return a property with an empty value and the 'undefined' type.
		
		return mResponseBuf.WriteF(
			"<response command=\"property_get\" transaction_id=\"%e\">"
				"<property name=\"%e\" fullname=\"%e\" type=\"undefined\" facet=\"\" size=\"0\" children=\"0\"/>"
			"</response>"
			, aTransactionId, name, name);
	}
	//else var and field were set by the called function.

	LPTSTR value_buf = NULL;
	switch (prop.kind)
	{
	case PropVar: err = GetPropertyInfo(*prop.var, prop, value_buf); break;
	case PropVarBkp: err = GetPropertyInfo(*prop.bkp, prop, value_buf); break;
	case PropField: err = GetPropertyInfo(*prop.field, prop); break;
	case PropValue: err = DEBUGGER_E_OK; break;
	}
	if (!err)
	{
		if (aIsPropertyGet)
		{
			mResponseBuf.WriteF(
				"<response command=\"property_get\" transaction_id=\"%e\">"
				, aTransactionId);
			name_buf.SetString(name); // prop.fullname is an alias of name_buf.
			// For simplicity and code size, we use the full caller-specified name
			// instead of trying to parse out the "short" name or record it during
			// ParsePropertyName (which would have to take into account differences
			// between UTF-8 and LPTSTR):
			prop.name = name;
			err = WritePropertyXml(prop);
		}
		else
		{
			mResponseBuf.WriteF(
				"<response command=\"property_value\" transaction_id=\"%e\" encoding=\"base64\" size=\""
				, aTransactionId);
			err = WritePropertyData(prop.value, prop.max_data);
		}
	}
	free(value_buf);
	return err ? err : mResponseBuf.Write("</response>");
}

DEBUGGER_COMMAND(Debugger::property_set)
{
	int err;
	char arg, *value;

	char *name = NULL, *new_value = NULL, *type = "string";
	int context_id = 0, depth = 0;

	for (int i = 0; i < aArgCount; ++i)
	{
		arg = ArgChar(aArgV, i);
		value = ArgValue(aArgV, i);
		switch (arg)
		{
		// property long name
		case 'n': name = value; break;
		// context id - optional, default zero. see PropertyContextType enum.
		case 'c': context_id = atoi(value); break;
		// stack depth - optional, default zero
		case 'd':
			depth = atoi(value);
			if (depth && (depth < 0 || depth >= mStack.Depth())) // Allow depth 0 even when stack is empty.
				return DEBUGGER_E_INVALID_STACK_DEPTH;
			break;
		// new base64-encoded value
		case '-': new_value = value; break;
		// data type: string, integer or float
		case 't': type = value; break;
		// not sure what the use of the 'length' arg is...
		case 'l': break;

		default:
			return DEBUGGER_E_INVALID_OPTIONS;
		}
	}

	if (!name || !new_value)
		return DEBUGGER_E_INVALID_OPTIONS;

	int always_use;
	switch (context_id)
	{
	// For consistency with property_get, create a local only if no global exists.
	//case PC_Local:	always_use = FINDVAR_LOCAL; break;
	case PC_Local:	always_use = FINDVAR_DEFAULT; break;
	case PC_Global:	always_use = FINDVAR_GLOBAL; break;
	default:
		return DEBUGGER_E_INVALID_CONTEXT;
	}

	PropertySource target;
	if (err = ParsePropertyName(name, depth, always_use, false, target))
		return err;

	// "Data must be encoded using base64." : https://xdebug.org/docs-dbgp.php
	// Fixed in v1.1.24.03 to expect base64 even for integer/float:
	int value_length = (int)Base64Decode(new_value, new_value);
	
	CString val_buf;
	ExprTokenType val;
	if (!strcmp(type, "integer"))
	{
		val.symbol = SYM_INTEGER;
		val.value_int64 = _atoi64(new_value);
	}
	else if (!strcmp(type, "float"))
	{
		val.symbol = SYM_FLOAT;
		val.value_double = atof(new_value);
	}
	else // Assume type is "string", since that's the only other supported type.
	{
		StringUTF8ToTChar(new_value, val_buf, value_length);
		val.symbol = SYM_STRING;
		val.marker = (LPTSTR)val_buf.GetString();
	}

	bool success;
	switch (target.kind)
	{
	case PropVarBkp:
		if (target.bkp->mType != VAR_ALIAS)
		{
			VarBkp &bkp = *target.bkp;
			if (bkp.mAttrib & VAR_ATTRIB_OBJECT)
			{
				bkp.mAttrib &= ~VAR_ATTRIB_OBJECT;
				bkp.mObject->Release();
			}
			if (val.symbol == SYM_STRING)
			{
				if ((val.marker_length + 1) * sizeof(TCHAR) > bkp.mByteCapacity && val.marker_length)
				{
					if (bkp.mHowAllocated == ALLOC_MALLOC)
						free(bkp.mCharContents);
					else
						bkp.mHowAllocated = ALLOC_MALLOC;
					bkp.mByteCapacity = (val_buf.GetAllocLength() + 1) * sizeof(TCHAR);
					bkp.mCharContents = val_buf.DetachBuffer();
					bkp.mAttrib &= ~VAR_ATTRIB_OFTEN_REMOVED;
				}
				else
					tmemcpy(bkp.mCharContents, val.marker, val.marker_length + 1);
				bkp.mByteLength = val.marker_length * sizeof(TCHAR);
			}
			else
			{
				bkp.mContentsInt64 = val.value_int64;
				bkp.mAttrib = bkp.mAttrib
					& ~(VAR_ATTRIB_CACHE | VAR_ATTRIB_UNINITIALIZED)
					| (val.symbol == SYM_INTEGER ? VAR_ATTRIB_HAS_VALID_INT64 : VAR_ATTRIB_HAS_VALID_DOUBLE) | VAR_ATTRIB_CONTENTS_OUT_OF_DATE;
			}
			success = true;
			break;
		}
		target.var = target.bkp->mAliasFor;
		// Fall through:
	case PropVar:
		success = !VAR_IS_READONLY(*target.var) && target.var->Assign(val);
		break;
	case PropField:
		success = target.field->Assign(val);
		break;
	default:
		success = false;
	}

	return mResponseBuf.WriteF(
		"<response command=\"property_set\" success=\"%i\" transaction_id=\"%e\"/>"
		, (int)success, aTransactionId);
}

DEBUGGER_COMMAND(Debugger::source)
{
	char arg, *value;

	char *filename = NULL;
	LineNumberType begin_line = 0, end_line = UINT_MAX;

	for (int i = 0; i < aArgCount; ++i)
	{
		arg = ArgChar(aArgV, i);
		value = ArgValue(aArgV, i);
		switch (arg)
		{
		case 'b': begin_line = strtoul(value, NULL, 10); break;
		case 'e': end_line = strtoul(value, NULL, 10); break;
		case 'f': filename = value; break;
		default:
			return DEBUGGER_E_INVALID_OPTIONS;
		}
	}

	if (!filename || begin_line > end_line)
		return DEBUGGER_E_INVALID_OPTIONS;

	int file_index;

	// Decode filename URI -> path, in-place.
	DecodeURI(filename);

	CStringTCharFromUTF8 filename_t(filename);

	// Ensure the file is actually a source file - i.e. don't let the debugger client retrieve any arbitrary file.
	for (file_index = 0; file_index < Line::sSourceFileCount; ++file_index)
	{
		if (!_tcsicmp(filename_t, Line::sSourceFile[file_index]))
		{
			TextFile tf;
			if (!tf.Open(filename_t, DEFAULT_READ_FLAGS, g_DefaultScriptCodepage))
				return DEBUGGER_E_CAN_NOT_OPEN_FILE;
			
			mResponseBuf.WriteF("<response command=\"source\" success=\"1\" transaction_id=\"%e\" encoding=\"base64\">"
								, aTransactionId);

			CStringA utf8_buf;
			TCHAR line_buf[LINE_SIZE + 2]; // May contain up to two characters of the previous line to simplify base64-encoding.
			int line_length;
			int line_remainder = 0;

			LineNumberType current_line = 0;
			
			while (-1 != (line_length = tf.ReadLine(line_buf + line_remainder, LINE_SIZE)))
			{
				if (++current_line >= begin_line)
				{
					if (current_line > end_line)
						break; // done.
					
					// Encode in multiples of 3 characters to avoid inserting padding characters.
					line_length += line_remainder; // Include remainder of previous line.
					line_remainder = line_length % 3;
					line_length -= line_remainder;

					if (line_length)
					{
						// Convert line to UTF-8.
						StringTCharToUTF8(line_buf, utf8_buf, line_length);
						// Base64-encode and write this line and its trailing newline character into the response buffer.
						if (mResponseBuf.WriteEncodeBase64(utf8_buf.GetString(), utf8_buf.GetLength()) != DEBUGGER_E_OK)
							goto break_outer_loop; // fail.
					}
					//else not enough data to encode in this iteration.

					if (line_remainder) // 0, 1 or 2.
					{
						line_buf[0] = line_buf[line_length];
						if (line_remainder > 1)
							line_buf[1] = line_buf[line_length + 1];
					}
				}
			}

			if (line_remainder) // Write any left-over characters.
			{
				StringTCharToUTF8(line_buf, utf8_buf, line_remainder);
				if (mResponseBuf.WriteEncodeBase64(utf8_buf.GetString(), utf8_buf.GetLength()) != DEBUGGER_E_OK)
					break; // fail.
			}

			if (!current_line || current_line < begin_line)
				break; // fail.
			// else if (current_line < end_line) -- just return what we can.

			return mResponseBuf.Write("</response>");
		}
	}
break_outer_loop:
	// If we got here, one of the following is true:
	//	- Something failed and used 'break'.
	//	- The requested file is not a known source file of this script.
	mResponseBuf.Clear();
	return mResponseBuf.WriteF(
		"<response command=\"source\" success=\"0\" transaction_id=\"%e\"/>"
		, aTransactionId);
}

int Debugger::redirect_std(char **aArgV, int aArgCount, char *aTransactionId, char *aCommandName)
{
	int new_mode = -1;
	// stdout and stderr accept exactly one arg: -c mode.
	if (aArgCount == 1 && ArgChar(aArgV, 0) == 'c')
		new_mode = atoi(ArgValue(aArgV, 0));

	if (new_mode < SR_Disabled || new_mode > SR_Redirect)
		return DEBUGGER_E_INVALID_OPTIONS;

	if (!stricmp(aCommandName, "stdout"))
		mStdOutMode = (StreamRedirectType)new_mode;
	else
		mStdErrMode = (StreamRedirectType)new_mode;

	return mResponseBuf.WriteF(
		"<response command=\"%s\" success=\"1\" transaction_id=\"%e\"/>"
		, aCommandName, aTransactionId);
}

DEBUGGER_COMMAND(Debugger::redirect_stdout)
{
	return redirect_std(aArgV, aArgCount, aTransactionId, "stdout");
}

DEBUGGER_COMMAND(Debugger::redirect_stderr)
{
	return redirect_std(aArgV, aArgCount, aTransactionId, "stderr");
}

//
// END DBGP COMMANDS
//

//
// DBGP STREAMS
//

int Debugger::WriteStreamPacket(LPCTSTR aText, LPCSTR aType)
{
	ASSERT(!mResponseBuf.mFailed);
	mResponseBuf.WriteF("<stream type=\"%s\">", aType);
	CStringUTF8FromTChar packet(aText);
	mResponseBuf.WriteEncodeBase64(packet, packet.GetLength() + 1); // Includes the null-terminator.
	mResponseBuf.Write("</stream>");
	return SendResponse();
}

void Debugger::OutputDebug(LPCTSTR aText)
{
	if (mStdErrMode != SR_Disabled) // i.e. SR_Copy or SR_Redirect
		WriteStreamPacket(aText, "stderr");
	if (mStdErrMode != SR_Redirect) // i.e. SR_Disabled or SR_Copy
		OutputDebugString(aText);
}

bool Debugger::FileAppendStdOut(LPCTSTR aText)
{
	if (mStdOutMode != SR_Disabled) // i.e. SR_Copy or SR_Redirect
		WriteStreamPacket(aText, "stdout");
	return mStdOutMode == SR_Redirect;
}

int Debugger::SendErrorResponse(char *aCommandName, char *aTransactionId, int aError, char *aExtraAttributes)
{
	mResponseBuf.WriteF("<response command=\"%s\" transaction_id=\"%e"
		, aCommandName, aTransactionId);
	
	if (aExtraAttributes)
		mResponseBuf.WriteF("\" %s>", aExtraAttributes);
	else
		mResponseBuf.Write("\">");

	mResponseBuf.WriteF("<error code=\"%i\"/></response>", aError);

	return SendResponse();
}

int Debugger::SendStandardResponse(char *aCommandName, char *aTransactionId)
{
	mResponseBuf.WriteF("<response command=\"%s\" transaction_id=\"%e\"/>"
						, aCommandName, aTransactionId);

	return SendResponse();
}

int Debugger::SendContinuationResponse(char *aCommand, char *aStatus, char *aReason)
{
	if (!aCommand)
	{
		switch (mInternalState)
		{
		case DIS_StepInto:	aCommand = "step_into";	break;
		case DIS_StepOver:	aCommand = "step_over";	break;
		case DIS_StepOut:	aCommand = "step_out";	break;
		case DIS_Run:		aCommand = "run";		break;
		// Seems more useful then silently failing:
		default:			aCommand = "";
		}
	}

	mResponseBuf.WriteF("<response command=\"%s\" status=\"%s\" reason=\"%s\" transaction_id=\"%e\"/>"
						, aCommand, aStatus, aReason, (LPCSTR)mContinuationTransactionId);

	return SendResponse();
}

// Debugger::ReceiveCommand
//
// Receives a full command line into mCommandBuf. If part of the next command is
// received, it remains in the buffer until the next call to ReceiveCommand.
//
int Debugger::ReceiveCommand(int *aCommandLength)
{
	ASSERT(mSocket != INVALID_SOCKET); // Shouldn't be at this point; will be caught by recv() anyway.
	ASSERT(!mCommandBuf.mFailed); // Should have been previously reset.

	DWORD u = 0;

	for(;;)
	{
		// Check data received in the previous iteration or by a previous call to ReceiveCommand().
		for ( ; u < mCommandBuf.mDataUsed; ++u)
		{
			if (mCommandBuf.mData[u] == '\0')
			{
				if (aCommandLength)
					*aCommandLength = u; // Does not include the null-terminator.
				return DEBUGGER_E_OK;
			}
		}

		// Init or expand the buffer as necessary.
		if (mCommandBuf.mDataUsed == mCommandBuf.mDataSize && mCommandBuf.Expand() != DEBUGGER_E_OK)
			return FatalError(); // This also calls mCommandBuf.Clear() via Disconnect().

		// Receive and append data.
		int bytes_received = recv(mSocket, mCommandBuf.mData + mCommandBuf.mDataUsed, (int)(mCommandBuf.mDataSize - mCommandBuf.mDataUsed), 0);

		if (bytes_received == SOCKET_ERROR)
			return FatalError();

		mCommandBuf.mDataUsed += bytes_received;
	}
}

// Debugger::SendResponse
//
// Sends a response to a command, using mResponseBuf.mData as the message body.
//
int Debugger::SendResponse()
{
	ASSERT(!mResponseBuf.mFailed);

	char response_header[DEBUGGER_RESPONSE_OVERHEAD];
	
	// Each message is prepended with a stringified integer representing the length of the XML data packet.
	Exp32or64(_itoa,_i64toa)(mResponseBuf.mDataUsed + DEBUGGER_XML_TAG_SIZE, response_header, 10);

	// The length and XML data are separated by a NULL byte.
	char *buf = strchr(response_header, '\0') + 1;

	// The XML document tag must always be present to provide XML version and encoding information.
	buf += sprintf(buf, "%s", DEBUGGER_XML_TAG);

	// Send the response header.
	if (  SOCKET_ERROR == send(mSocket, response_header, (int)(buf - response_header), 0)
	   // Messages sent by the debugger engine must always be NULL terminated.
	   // Failure to write the last byte should be extremely rare, so no attempt
	   // is made to recover from that condition.
	   || DEBUGGER_E_OK != mResponseBuf.Write("\0", 1)
	   // Send the message body.
	   || SOCKET_ERROR == send(mSocket, mResponseBuf.mData, (int)mResponseBuf.mDataUsed, 0)  )
	{
		// Unrecoverable error: disconnect the debugger.
		return FatalError();
	}

	mResponseBuf.Clear();
	return DEBUGGER_E_OK;
}

// Debugger::Connect
//
// Connect to a debugger UI. Returns a Winsock error code on failure, otherwise 0.
//
int Debugger::Connect(const char *aAddress, const char *aPort)
{
	int err;
	WSADATA wsadata;
	SOCKET s;
	
	if (WSAStartup(MAKEWORD(2,2), &wsadata))
		return FatalError();
	
	s = socket(AF_INET, SOCK_STREAM, IPPROTO_TCP);

	if (s != INVALID_SOCKET)
	{
		addrinfo hints = {0};
		addrinfo *res;
		
		hints.ai_family = AF_INET;
		hints.ai_socktype = SOCK_STREAM;
		hints.ai_protocol = IPPROTO_TCP;

		err = getaddrinfo(aAddress, aPort, &hints, &res);
		
		if (err == 0)
		{
			for (;;)
			{
				err = connect(s, res->ai_addr, (int)res->ai_addrlen);
				if (err == 0)
					break;
				switch (MessageBox(g_hWnd, DEBUGGER_ERR_FAILEDTOCONNECT, g_script.mFileSpec, MB_ABORTRETRYIGNORE | MB_ICONSTOP | MB_SETFOREGROUND | MB_APPLMODAL))
				{
				case IDABORT:
					g_script.ExitApp(EXIT_CLOSE);
					// If it didn't exit (due to OnExit), fall through to the next case:
				case IDIGNORE:
					closesocket(s);
					return DEBUGGER_E_INTERNAL_ERROR;
				}
			}
			
			freeaddrinfo(res);
			
			if (err == 0)
			{
				mSocket = s;

				CStringUTF8FromTChar ide_key(CString().GetEnvironmentVariable(_T("DBGP_IDEKEY")));
				CStringUTF8FromTChar session(CString().GetEnvironmentVariable(_T("DBGP_COOKIE")));

				// Clear the buffer in case of a previous failed session.
				mResponseBuf.Clear();

				// Write init message.
				mResponseBuf.WriteF("<init appid=\"" AHK_NAME "\" ide_key=\"%e\" session=\"%e\" thread=\"%u\" parent=\"\" language=\"" DEBUGGER_LANG_NAME "\" protocol_version=\"1.0\" fileuri=\""
					, ide_key.GetString(), session.GetString(), GetCurrentThreadId());
				mResponseBuf.WriteFileURI(U4T(g_script.mFileSpec));
				mResponseBuf.Write("\"/>");

				if (SendResponse() == DEBUGGER_E_OK)
				{
					// mCurrLine isn't updated unless the debugger is connected, so set it now.
					// g_script.mCurrLine should always be non-NULL after the script is loaded,
					// even if no threads are active.
					mCurrLine = g_script.mCurrLine;
					return DEBUGGER_E_OK;
				}

				mSocket = INVALID_SOCKET; // Don't want FatalError() to attempt a second closesocket().
			}
		}

		closesocket(s);
	}

	WSACleanup();
	return FatalError(DEBUGGER_ERR_FAILEDTOCONNECT DEBUGGER_ERR_DISCONNECT_PROMPT);
}

// Debugger::Disconnect
//
// Disconnect the debugger UI.
//
int Debugger::Disconnect()
{
	if (mSocket != INVALID_SOCKET)
	{
		shutdown(mSocket, 2);
		closesocket(mSocket);
		mSocket = INVALID_SOCKET;
		WSACleanup();
	}
	// These are reset in case we re-attach to the debugger client later:
	mCommandBuf.Clear();
	mResponseBuf.Clear();
	mStdOutMode = SR_Disabled;
	mStdErrMode = SR_Disabled;
	if (mInternalState == DIS_Break)
		ExitBreakState();
	mInternalState = DIS_Starting;
	return DEBUGGER_E_OK;
}

// Debugger::Exit
//
// Gracefully end debug session.  Called on script exit.  Also called by "detach" DBGp command.
//
void Debugger::Exit(ExitReasons aExitReason, char *aCommandName)
{
	if (mSocket == INVALID_SOCKET)
		return;
	// Don't care if it fails as we may be exiting due to a previous failure.
	SendContinuationResponse(aCommandName, "stopped", aExitReason == EXIT_ERROR ? "error" : "ok");
	Disconnect();
}

int Debugger::FatalError(LPCTSTR aMessage)
{
	g_Debugger.Disconnect();

	if (IDNO == MessageBox(g_hWnd, aMessage, g_script.mFileSpec, MB_YESNO | MB_ICONSTOP | MB_SETFOREGROUND | MB_APPLMODAL))
	{
		// This might not exit, depending on OnExit:
		g_script.ExitApp(EXIT_CLOSE);
	}
	return DEBUGGER_E_INTERNAL_ERROR;
}

const char *Debugger::sBase64Chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";

#define BINARY_TO_BASE64_CHAR(b) (sBase64Chars[(b) & 63])
#define BASE64_CHAR_TO_BINARY(q) (strchr(sBase64Chars, q)-sBase64Chars)

// Encode base 64 data.
size_t Debugger::Base64Encode(char *aBuf, const char *aInput, size_t aInputSize/* = -1*/)
{
	UINT_PTR buffer;
	size_t i, len = 0;

	if (aInputSize == -1) // Direct comparison since aInputSize is unsigned.
		aInputSize = strlen(aInput);

	for (i = aInputSize; i > 2; i -= 3)
	{
		buffer = (UCHAR)aInput[0] << 16 | (UCHAR)aInput[1] << 8 | (UCHAR)aInput[2]; // L39: Fixed for chars outside the range 0..127. [thanks jackieku]
		aInput += 3;

		aBuf[len + 0] = BINARY_TO_BASE64_CHAR(buffer >> 18);
		aBuf[len + 1] = BINARY_TO_BASE64_CHAR(buffer >> 12);
		aBuf[len + 2] = BINARY_TO_BASE64_CHAR(buffer >> 6);
		aBuf[len + 3] = BINARY_TO_BASE64_CHAR(buffer);
		len += 4;
	}
	if (i > 0)
	{
		buffer = (UCHAR)aInput[0] << 16;
		if (i > 1)
			buffer |= (UCHAR)aInput[1] << 8;
		// aInput not incremented as it is not used below.

		aBuf[len + 0] = BINARY_TO_BASE64_CHAR(buffer >> 18);
		aBuf[len + 1] = BINARY_TO_BASE64_CHAR(buffer >> 12);
		aBuf[len + 2] = (i > 1) ? BINARY_TO_BASE64_CHAR(buffer >> 6) : '=';
		aBuf[len + 3] = '=';
		len += 4;
	}
	aBuf[len] = '\0';
	return len;
}

// Decode base 64 data. aBuf and aInput may point to the same buffer.
size_t Debugger::Base64Decode(char *aBuf, const char *aInput, size_t aInputSize/* = -1*/)
{
	UINT_PTR buffer;
	size_t i, len = 0;

	if (aInputSize == -1) // Direct comparison since aInputSize is unsigned.
		aInputSize = strlen(aInput);

	while (aInputSize > 0 && aInput[aInputSize-1] == '=')
		--aInputSize;

	for (i = aInputSize; i > 3; i -= 4)
	{
		buffer	= BASE64_CHAR_TO_BINARY(aInput[0]) << 18 // L39: Fixed bad reliance on order of *side-effects++. [thanks fincs]
				| BASE64_CHAR_TO_BINARY(aInput[1]) << 12
				| BASE64_CHAR_TO_BINARY(aInput[2]) << 6
				| BASE64_CHAR_TO_BINARY(aInput[3]);
		aInput += 4;

		aBuf[len + 0] = (buffer >> 16) & 255;
		aBuf[len + 1] = (buffer >> 8) & 255;
		aBuf[len + 2] = buffer & 255;
		len += 3;
	}

	if (i > 1)
	{
		buffer  = BASE64_CHAR_TO_BINARY(aInput[0]) << 18
				| BASE64_CHAR_TO_BINARY(aInput[1]) << 12;
		if (i > 2)
			buffer |= BASE64_CHAR_TO_BINARY(aInput[2]) << 6;
		// aInput not incremented as it is not used below.

		aBuf[len++] = (buffer >> 16) & 255;
		if (i > 2)
			aBuf[len++] = (buffer >> 8) & 255;
	}
	aBuf[len] = '\0';
	return len;
}

//
// class Debugger::Buffer - simplifies memory management.
//

// Write data into the buffer, expanding it as necessary.
int Debugger::Buffer::Write(char *aData, size_t aDataSize)
{
	if (mFailed) // See WriteF() for comments.
		return DEBUGGER_E_INTERNAL_ERROR;

	if (aDataSize == -1)
		aDataSize = strlen(aData);

	if (aDataSize == 0)
		return DEBUGGER_E_OK;

	if (ExpandIfNecessary(mDataUsed + aDataSize) != DEBUGGER_E_OK)
		return DEBUGGER_E_INTERNAL_ERROR;

	memcpy(mData + mDataUsed, aData, aDataSize);
	mDataUsed += aDataSize;
	return DEBUGGER_E_OK;
}

// Write formatted data into the buffer. Supports %s (char*), %e (char*, "&'<> escaped), %i (int), %u (unsigned int), %p (UINT_PTR).
int Debugger::Buffer::WriteF(const char *aFormat, ...)
{
	if (mFailed)
	{
		// A prior Write() failed and the now-invalid data hasn't yet been cleared
		// from the buffer.  Abort.  This allows numerous other parts of the code
		// to omit error-checking.
		return DEBUGGER_E_INTERNAL_ERROR;
	}

	int i;
	size_t len;
	char c;
	const char *format_ptr, *s, *param_ptr, *entity;
	char number_buf[MAX_INTEGER_SIZE];
	va_list vl;
	
	for (len = 0, i = 0; i < 2; ++i)
	{
		va_start(vl, aFormat);

		// Calculate the required buffer size.
		for (format_ptr = aFormat; c = *format_ptr; ++format_ptr)
		{
			if (c == '%')
			{
				switch (format_ptr[1])
				{
				case 's': s = va_arg(vl, char*); break;
				case 'i': s = _itoa(va_arg(vl, int), number_buf, 10); break;
				case 'u': s = _ultoa(va_arg(vl, unsigned long), number_buf, 10); break;
				case 'p': s = Exp32or64(_ultoa,_ui64toa)(va_arg(vl, UINT_PTR), number_buf, 10); break;

				case 'e': // String, replace "&'<> with appropriate XML entity.
				{
					s = va_arg(vl, char*);
					if (i == 0)
					{
						for (param_ptr = s; *param_ptr; ++param_ptr)
						{
							switch (*param_ptr)
							{
							case '"': case '\'':	len += 6; break; // &quot; or &apos;
							case '&':				len += 5; break; // &amp;
							case '<': case '>':		len += 4; break; // &lt; or &gt;
							default: ++len;
							}
						}
					}
					else
					{
						for (param_ptr = s; c = *param_ptr; ++param_ptr)
						{
							switch (c)
							{
							case '"': entity = "quot"; break;
							case '\'': entity = "apos"; break;
							case '&': entity = "amp"; break;
							case '<': entity = "lt"; break;
							case '>': entity = "gt"; break;
							default:
								mData[mDataUsed++] = c;
								continue;
							}
							// One of: "'&<> - entity is set to the appropriate entity name.
							mDataUsed += sprintf(mData + mDataUsed, "&%s;", entity);
						}
					}
					++format_ptr; // Skip %, outer loop will skip format char.
					//s = NULL; // Skip section below.
					//break;
					continue;
				}

				default:
					s = NULL; // Skip section below.
				} // switch (format_ptr[1])

				if (s)
				{	// %s, %i or %u.
					if (i == 0)
					{	// Calculate required buffer space on first iteration.
						len += strlen(s);
					}
					else if (len = strlen(s))
					{	// Write into buffer on second iteration.
						// memcpy as we don't want the terminating null character.
						memcpy(mData + mDataUsed, s, len);
						mDataUsed += len;
					}
					++format_ptr; // Skip %, outer loop will skip format char.
					continue;
				}
			} // if (c == '%')
			// Count or copy character as is.
			if (i == 0)
				++len;
			else
				mData[mDataUsed++] = *format_ptr;
		} // for (format_ptr = aFormat; c = *format_ptr; ++format_ptr)

		if (i == 0)
			if (ExpandIfNecessary(mDataUsed + len) != DEBUGGER_E_OK)
				return DEBUGGER_E_INTERNAL_ERROR;
	} // for (len = 0, i = 0; i < 2; ++i)

	return DEBUGGER_E_OK;
}

// Convert a file path to a URI and write it to the buffer.
int Debugger::Buffer::WriteFileURI(const char *aPath)
{
	int c, len = 9; // 8 for "file:///", 1 for '\0' (written by sprintf()).

	// Calculate required buffer size for path after encoding.
	for (const char *ptr = aPath; c = *ptr; ++ptr)
	{
		if (cisalnum(c) || strchr("-_.!~*'()/\\", c))
			++len;
		else
			len += 3;
	}

	// Ensure the buffer contains enough space.
	if (ExpandIfNecessary(mDataUsed + len) != DEBUGGER_E_OK)
		return DEBUGGER_E_INTERNAL_ERROR;

	Write("file:///", 8);

	// Write to the buffer, encoding as we go.
	for (const char *ptr = aPath; c = *ptr; ++ptr)
	{
		if (cisalnum(c) || strchr("-_.!~*()/", c))
		{
			mData[mDataUsed++] = (char)c;
		}
		else if (c == '\\')
		{
			// This could be encoded as %5C, but it's more readable this way:
			mData[mDataUsed++] = '/';
		}
		else
		{
			len = sprintf(mData + mDataUsed, "%%%02X", c & 0xff);
			if (len != -1)
				mDataUsed += len;
		}
	}

	return DEBUGGER_E_OK;
}

int Debugger::Buffer::WriteEncodeBase64(const char *aInput, size_t aInputSize, bool aSkipBufferSizeCheck/* = false*/)
{
	if (aInputSize)
	{
		if (!aSkipBufferSizeCheck)
		{
			// Ensure required buffer space is available.
			if (ExpandIfNecessary(mDataUsed + DEBUGGER_BASE64_ENCODED_SIZE(aInputSize)) != DEBUGGER_E_OK)
				return DEBUGGER_E_INTERNAL_ERROR;
		}
		//else caller has already ensured there is enough space and wants to be absolutely sure mData isn't reallocated.
		ASSERT(mDataUsed + aInputSize < mDataSize);
		
		if (aInput)
			mDataUsed += Debugger::Base64Encode(mData + mDataUsed, aInput, aInputSize);
		//else caller wanted to reserve some buffer space, probably to read the raw data into.
	}
	//else there's nothing to write, just return OK.  Calculations above can't handle (size_t)0
	return DEBUGGER_E_OK;
}

// Decode a file URI in-place.
void Debugger::DecodeURI(char *aUri)
{
	char *end = strchr(aUri, 0);
	char escape[3];
	escape[2] = '\0';

	if (!strnicmp(aUri, "file:///", 8))
	{
		// Use memmove since I'm not sure if strcpy can handle overlap.
		end -= 8;
		memmove(aUri, aUri + 8, end - aUri);
	}
	// Some debugger UI's use file://path even though it is invalid (it should be file:///path or file://server/path).
	// For compatibility with these UI's, support file://.
	else if (!strnicmp(aUri, "file://", 7))
	{
		end -= 7;
		memmove(aUri, aUri + 7, end - aUri);
	}

	char *dst, *src;
	for (src = dst = aUri; src < end; ++src, ++dst)
	{
		if (*src == '%' && src + 2 < end)
		{
			escape[0] = *++src;
			escape[1] = *++src;
			*dst = (char)strtol(escape, NULL, 16);
		}
		else if (*src == '/')
			*dst = '\\';
		else
			*dst = *src;
	}
	*dst = '\0';
}

// Initialize or expand the buffer, don't care how much.
int Debugger::Buffer::Expand()
{
	return ExpandIfNecessary(mDataSize ? mDataSize * 2 : DEBUGGER_INITIAL_BUFFER_SIZE);
}

// Expand as necessary to meet a minimum required size.
int Debugger::Buffer::ExpandIfNecessary(size_t aRequiredSize)
{
	if (mFailed)
		return DEBUGGER_E_INTERNAL_ERROR;

	size_t new_size;
	for (new_size = mDataSize ? mDataSize : DEBUGGER_INITIAL_BUFFER_SIZE
		; new_size < aRequiredSize
		; new_size *= 2);

	if (new_size > mDataSize)
	{
		// For simplicity, this preserves all of mData not just the first mDataUsed bytes.  Some sections may rely on this.
		char *new_data = (char*)realloc(mData, new_size);

		if (new_data == NULL)
		{
			mFailed = TRUE;
			return DEBUGGER_E_INTERNAL_ERROR;
		}

		mData = new_data;
		mDataSize = new_size;
	}
	return DEBUGGER_E_OK;
}

// Remove data from the front of the buffer (i.e. after it is processed).
void Debugger::Buffer::Remove(size_t aDataSize)
{
	ASSERT(aDataSize <= mDataUsed);
	// Move remaining data to the front of the buffer.
	if (aDataSize < mDataUsed)
		memmove(mData, mData + aDataSize, mDataUsed - aDataSize);
	mDataUsed -= aDataSize;
}

void Debugger::Buffer::Clear()
{
	mDataUsed = 0;
	mFailed = FALSE;
}



DbgStack::Entry *DbgStack::Push()
{
	if (mTop == mTopBound)
		Expand();
	if (mTop >= mBottom)
	{
		// Whenever this entry != mTop, Entry::line is used instead of g_Debugger.mCurrLine,
		// which means that it needs to point to the last executed line at that depth. The
		// following fixes a bug where the line number of this entry appears to revert to
		// the first line of the function or sub whenever the debugger steps into another
		// function or sub.  Entry::line is now also used by BIF_Exception even when the
		// debugger is disconnected, which has two consequences:
		//  - g_script.mCurrLine must be used in place of g_Debugger.mCurrLine.
		//  - Changing PreExecLine() to update mStack.mTop->line won't help.
		mTop->line = g_script.mCurrLine;
	}
	return ++mTop;
}

void DbgStack::Push(TCHAR *aDesc)
{
	Entry &s = *Push();
	s.line = NULL;
	s.desc = aDesc;
	s.type = SE_Thread;
}
	
void DbgStack::Push(Label *aSub)
{
	Entry &s = *Push();
	s.line = aSub->mJumpToLine;
	s.sub  = aSub;
	s.type = SE_Sub;
}

void DbgStack::Push(UDFCallInfo *aUDF)
{
	Entry &s = *Push();
	s.line = aUDF->func->mJumpToLine;
	s.udf = aUDF;
	s.type = SE_UDF;
}


void DbgStack::GetLocalVars(int aDepth, Var **&aVar, Var **&aVarEnd, VarBkp *&aBkp, VarBkp *&aBkpEnd)
{
	DbgStack::Entry *se = mTop - aDepth;
	for (;;)
	{
		if (se <= mBottom)
			return;
		if (se->type == DbgStack::SE_UDF)
			break;
		--se;
	}
	Func &func = *se->udf->func;
	if (func.mInstances > 1 && aDepth > 0)
	{
		while (++se <= mTop)
		{
			if (se->type == DbgStack::SE_UDF && se->udf->func == &func)
			{
				// This instance interrupted the target instance, so its backup
				// contains the values of the target instance's local variables.
				aBkp = se->udf->backup;
				aBkpEnd = aBkp + se->udf->backup_count;
				return;
			}
		}
	}
	// Since above did not return, this instance wasn't interrupted.
	aVar = func.mVar;
	aVarEnd = aVar + func.mVarCount;
}


void Debugger::PropertyWriter::WriteProperty(LPCSTR aName, ExprTokenType &aValue)
{
	mProp.fullname.AppendFormat(".%s", aName);
	_WriteProperty(aValue);
}


void Debugger::PropertyWriter::WriteProperty(ExprTokenType &aKey, ExprTokenType &aValue)
{
	switch (aKey.symbol)
	{
	case SYM_INTEGER: mProp.fullname.AppendFormat("[%Ii]", aKey.value_int64); break;
	case SYM_OBJECT: mProp.fullname.AppendFormat("[Object(%Ii)]", aKey.object); break;
	default:
		ASSERT(aKey.symbol == SYM_STRING);
		mDbg.AppendKeyName(mProp.fullname, mNameLength, CStringUTF8FromTChar(aKey.marker));
	}
	_WriteProperty(aValue);
}


void Debugger::PropertyWriter::_WriteProperty(ExprTokenType &aValue)
{
	if (mError)
		return;
	PropertyInfo prop(mProp.fullname);
	// Find the property's "relative" name at the end of the buffer:
	prop.name = mProp.fullname.GetString() + mNameLength;
	if (*prop.name == '.')
		prop.name++;
	// Write the property (and if it contains an object, any child properties):
	prop.value.CopyValueFrom(aValue);
	//prop.page = 0; // "the childrens pages are always the first page."
	prop.pagesize = mProp.pagesize;
	prop.max_data = mProp.max_data;
	prop.max_depth = mProp.max_depth - mDepth;
	mError = mDbg.WritePropertyXml(prop);
	// Truncate back to the parent property name:
	mProp.fullname.Truncate(mNameLength);
}


void Debugger::PropertyWriter::BeginProperty(LPCSTR aName, LPCSTR aType, int aNumChildren, DebugCookie &aCookie)
{
	if (mError)
		return;

	mDepth++;

	if (mDepth == 1) // Write <property> for the object itself.
	{
		LPTSTR classname = mObject->Type();
		mError = mDbg.mResponseBuf.WriteF("<property name=\"%e\" fullname=\"%e\" type=\"%s\" facet=\"%s\" classname=\"%s\" address=\"%p\" size=\"0\" page=\"%i\" pagesize=\"%i\" children=\"%i\" numchildren=\"%i\">"
					, mProp.name, mProp.fullname.GetString(), aType, mProp.facet, U4T(classname), mObject, mProp.page, mProp.pagesize, aNumChildren > 0, aNumChildren);
		return;
	}

	// aName is the raw name of this property.  Before writing it into the response,
	// convert it to the required expression-like form.  Do this conversion within
	// mProp.fullname so it can be used for the fullname attribute:
	mDbg.AppendKeyName(mProp.fullname, mNameLength, aName);

	// Find the property's "relative" name at the end of the buffer:
	LPCSTR name = mProp.fullname.GetString() + mNameLength;
	if (*name == '.')
		name++;
	
	// Save length of outer property name and update mNameLength.
	aCookie = (DebugCookie)mNameLength;
	mNameLength = mProp.fullname.GetLength();

	mError = mDbg.mResponseBuf.WriteF("<property name=\"%e\" fullname=\"%e\" type=\"%s\" size=\"0\" page=\"0\" pagesize=\"%i\" children=\"%i\" numchildren=\"%i\">"
				, name, mProp.fullname.GetString(), aType, mProp.pagesize, aNumChildren > 0, aNumChildren);
}


void Debugger::PropertyWriter::EndProperty(DebugCookie aCookie)
{
	if (mError)
		return;

	mDepth--;

	if (mDepth > 0) // If we just ended a child property...
	{
		mNameLength = (size_t)aCookie; // Restore to the value it had before BeginProperty().
		mProp.fullname.Truncate(mNameLength);
	}
	
	mError = mDbg.mResponseBuf.Write("</property>");
}


#endif


// Helper for Var::ToText().

LPTSTR Var::ObjectToText(LPTSTR aName, LPTSTR aBuf, int aBufSize)
{
	LPTSTR aBuf_orig = aBuf;
	aBuf += sntprintf(aBuf, aBufSize, _T("%s: %s object"), aName, mObject->Type());
	if (ComObject *cobj = dynamic_cast<ComObject *>(mObject))
		aBuf += sntprintf(aBuf, BUF_SPACE_REMAINING, _T(" {wrapper: 0x%IX, vt: 0x%04hX, value: 0x%I64X}"), cobj, cobj->mVarType, cobj->mVal64);
	else
		aBuf += sntprintf(aBuf, BUF_SPACE_REMAINING, _T(" {address: 0x%IX}"), mObject);
	return aBuf;
}
