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
		return ProcessCommands();
	}

	if ((mInternalState == DIS_StepInto
		|| mInternalState == DIS_StepOver && mStack.Depth() <= mContinuationDepth
		|| mInternalState == DIS_StepOut && mStack.Depth() < mContinuationDepth) // Due to short-circuit boolean evaluation, mStack.Depth() is only evaluated once and only if mInternalState is StepOver or StepOut.
		// Although IF/ELSE/LOOP skips its block-begin, standalone/function-body block-begin still gets here; we want to skip it:
		&& aLine->mActionType != ACT_BLOCK_BEGIN && (aLine->mActionType != ACT_BLOCK_END || aLine->mAttribute) // Ignore { and }; except for function-end, since we want to break there after a "return" to inspect variables while they're still in scope.
		&& aLine->mLineNumber) // Some scripts (i.e. LowLevel/code.ahk) use mLineNumber==0 to indicate the Line has been generated and injected by the script.
	{
		return ProcessCommands();
	}
	
	return DEBUGGER_E_OK;
}


int Debugger::ProcessCommands()
{
	int err;

	HookType active_hooks = GetActiveHooks();
	if (active_hooks)
		AddRemoveHooks(0, true);
	
	if (mInternalState != DIS_Starting)
	{
		// Send response for the previous continuation command.
		if (err = SendContinuationResponse())
			goto ProcessCommands_err_return;
			//return err;
	}
	mInternalState = DIS_Break;

	for(;;)
	{
		int command_length;

		if (err = ReceiveCommand(&command_length))
		{	// Internal/winsock/protocol error.
			goto ProcessCommands_err_return;
			//return err;
		}

		char *command = mCommandBuf.mData;
		char *args = strchr(command, ' ');
		// Split command name and args.
		if (args)
			*args++ = '\0';

		if (args && *args)
		{
			err = DEBUGGER_E_UNIMPL_COMMAND;

			if (!strcmp(command, "run"))
				mInternalState = DIS_Run;
			else if (!strncmp(command, "step_", 5))
			{
				if (!strcmp(command + 5, "into"))
					mInternalState = DIS_StepInto;
				else if (!strcmp(command + 5, "over"))
					mInternalState = DIS_StepOver;
				else if (!strcmp(command + 5, "out"))
					mInternalState = DIS_StepOut;
			}
			else if (!strncmp(command, "feature_", 8))
			{
				if (!strcmp(command + 8, "get"))
					err = feature_get(args);
				else if (!strcmp(command + 8, "set"))
					err = feature_set(args);
			}
			else if (!strncmp(command, "breakpoint_", 11))
			{
				if (!strcmp(command + 11, "set"))
					err = breakpoint_set(args);
				else if (!strcmp(command + 11, "get"))
					err = breakpoint_get(args);
				else if (!strcmp(command + 11, "update"))
					err = breakpoint_update(args);
				else if (!strcmp(command + 11, "remove"))
					err = breakpoint_remove(args);
				else if (!strcmp(command + 11, "list"))
					err = breakpoint_list(args);
			}
			else if (!strncmp(command, "stack_", 6))
			{
				if (!strcmp(command + 6, "depth"))
					err = stack_depth(args);
				else if (!strcmp(command + 6, "get"))
					err = stack_get(args);
			}
			else if (!strncmp(command, "context_", 8))
			{
				if (!strcmp(command + 8, "names"))
					err = context_names(args);
				else if (!strcmp(command + 8, "get"))
					err = context_get(args);
			}
			else if (!strncmp(command, "property_", 9))
			{
				if (!strcmp(command + 9, "get"))
					err = property_get(args);
				else if (!strcmp(command + 9, "set"))
					err = property_set(args);
				else if (!strcmp(command + 9, "value"))
					err = property_value(args);
			}
			else if (!strcmp(command, "status"))
				err = status(args);
			else if (!strcmp(command, "source"))
				err = source(args);
			else if (!strcmp(command, "stdout"))
				err = redirect_stdout(args);
			else if (!strcmp(command, "stderr"))
				err = redirect_stderr(args);
			else if (!strcmp(command, "typemap_get"))
				err = typemap_get(args);
			else if (!strcmp(command, "detach"))
			{	// User wants to stop the debugger but let the script keep running.
				Exit(EXIT_NONE); // Anything but EXIT_ERROR.  Sends "stopped" response, then disconnects.
				return DEBUGGER_E_OK;
			}
			else if (!strcmp(command, "stop"))
			{
				err = stop(args);
				// Above should exit the program.
			}
		}
		else
		{	// All commands require at least the -i arg.
			err = DEBUGGER_E_INVALID_OPTIONS;
		}

		mCommandBuf.Remove(command_length + 1);

		if (err == DEBUGGER_E_INTERNAL_ERROR)
			goto ProcessCommands_err_return;
			//return err;

		if (err) // Command returned error, is not implemented, or is a continuation command.
		{
			char *transaction_id;
			
			if (args && (transaction_id = strstr(args, "-i ")))
			{
				transaction_id += 3;
				char *transaction_id_end;
				if (transaction_id_end = strstr(transaction_id, " -"))
					*transaction_id_end = '\0';
			}
			else transaction_id = "";
			
			if (mInternalState != DIS_Break) // Received a continuation command.
			{
				if (*transaction_id)
				{
					mContinuationDepth = mStack.Depth();
					mContinuationTransactionId = transaction_id;
					break;
				}
				err = DEBUGGER_E_INVALID_OPTIONS;
			}
			
			// Assume command (if called) has not sent a response.
			if (err = SendErrorResponse(command, transaction_id, err))
			{	// Internal/winsock/protocol error.
				goto ProcessCommands_err_return;
				//return err;
			}
			
			// Continue processing commands.
		}
	}
	// Received a continuation command.
	// Script execution should continue until a break condition is met:
	//	If step_into was used, break on next line.
	//	If step_out was used, break on next line after return.
	//	If step_over was used, break on next line at same recursion depth.
	//	Also break when an active breakpoint is reached.
	//return DEBUGGER_E_OK;
	err = DEBUGGER_E_OK;

ProcessCommands_err_return:
	if (active_hooks)
		AddRemoveHooks(active_hooks, true);
	return err;
}


//
// DBGP COMMANDS
//

// For simple commands that accept only -i transaction_id:
#define DEBUGGER_COMMAND_INIT_TRANSACTION_ID \
	if (strncmp(aArgs, "-i ", 3) || strstr(aArgs + 3, " -")) \
		return DEBUGGER_E_INVALID_OPTIONS; \
	char *transaction_id = aArgs + 3

// Calculate base64-encoded size of data, including NULL terminator. org_size must be > 0 if unsigned.
#define DEBUGGER_BASE64_ENCODED_SIZE(org_size) ((((DWORD)(org_size)-1)/3+1)*4 +1)

DEBUGGER_COMMAND(Debugger::status)
{
	DEBUGGER_COMMAND_INIT_TRANSACTION_ID;

	char *status;
	switch (mInternalState)
	{
	case DIS_Starting:	status = "starting";	break;
	case DIS_Run:		status = "running";		break;
	default:			status = "break";
	}

	mResponseBuf.WriteF("<response command=\"status\" status=\"%s\" reason=\"ok\" transaction_id=\"%e\"/>"
						, status, transaction_id);

	return SendResponse();
}

DEBUGGER_COMMAND(Debugger::feature_get)
{
	int err;
	char arg, *value;

	char *transaction_id, *feature_name;

	for(;;)
	{
		if (err = GetNextArg(aArgs, arg, value))
			return err;
		if (!arg)
			break;
		switch (arg)
		{
		case 'i': transaction_id = value; break;
		case 'n': feature_name = value; break;
		default:
			return DEBUGGER_E_INVALID_OPTIONS;
		}
	}

	if (!transaction_id || !feature_name)
		return DEBUGGER_E_INVALID_OPTIONS;

	bool supported = false; // Is %feature_name% a supported feature string?
	char *setting = "";

	if (!strncmp(feature_name, "language_", 9)) {
		if (supported = !strcmp(feature_name + 9, "supports_threads"))
			setting = "0";
		else if (supported = !strcmp(feature_name + 9, "name"))
			setting = DEBUGGER_LANG_NAME;
		else if (supported = !strcmp(feature_name + 9, "version"))
#ifdef UNICODE
			setting = AHK_VERSION " (Unicode)";
#else
			setting = AHK_VERSION;
#endif
	} else if (supported = !strcmp(feature_name, "encoding"))
		setting = "UTF-8";
	else if (supported = !strcmp(feature_name, "protocol_version"))
		setting = "1";
	else if (supported = !strcmp(feature_name, "supports_async"))
		setting = "0"; // TODO: Async support for status, breakpoint_set, break, etc.
	// Not supported: data_encoding - assume base64.
	// Not supported: breakpoint_languages - assume only %language_name% is supported.
	else if (supported = !strcmp(feature_name, "breakpoint_types"))
		setting = "line";
	else if (supported = !strcmp(feature_name, "multiple_sessions"))
		setting = "0";
	else if (supported = !strcmp(feature_name, "max_data"))
		setting = _itoa(mMaxPropertyData, (char*)_alloca(MAX_INTEGER_SIZE), 10);
	else if (supported = !strcmp(feature_name, "max_children"))
		setting = _ultoa(mMaxChildren, (char*)_alloca(MAX_INTEGER_SIZE), 10);
	else if (supported = !strcmp(feature_name, "max_depth"))
		setting = _ultoa(mMaxDepth, (char*)_alloca(MAX_INTEGER_SIZE), 10);
	// TODO: STOPPING state for retrieving variable values, etc. after the script finishes, then implement supports_postmortem feature name. Requires debugger client support.
	else
	{
		supported = !strcmp(feature_name, "run")
			|| !strncmp(feature_name, "step_", 5)
				&& (!strcmp(feature_name + 5, "into")
				 || !strcmp(feature_name + 5, "over")
				 || !strcmp(feature_name + 5, "out"))
			|| !strncmp(feature_name, "feature_", 8)
				&& (!strcmp(feature_name + 8, "get")
				 || !strcmp(feature_name + 8, "set"))
			|| !strncmp(feature_name, "breakpoint_", 11)
				&& (!strcmp(feature_name + 11, "set")
				 || !strcmp(feature_name + 11, "get")
				 || !strcmp(feature_name + 11, "update")
				 || !strcmp(feature_name + 11, "remove")
				 || !strcmp(feature_name + 11, "list"))
			|| !strncmp(feature_name, "stack_", 6)
				&& (!strcmp(feature_name + 6, "depth")
				 || !strcmp(feature_name + 6, "get"))
			|| !strncmp(feature_name, "context_", 8)
				&& (!strcmp(feature_name + 8, "names")
				 || !strcmp(feature_name + 8, "get"))
			|| !strncmp(feature_name, "property_", 9)
				&& (!strcmp(feature_name + 9, "get")
				 || !strcmp(feature_name + 9, "set")
				 || !strcmp(feature_name + 9, "value"))
			|| !strcmp(feature_name, "status")
			|| !strcmp(feature_name, "source")
			|| !strcmp(feature_name, "stdout")
			|| !strcmp(feature_name, "stderr")
			|| !strcmp(feature_name, "typemap_get")
			|| !strcmp(feature_name, "detach")
			|| !strcmp(feature_name, "stop");
	}

	mResponseBuf.WriteF("<response command=\"feature_get\" feature_name=\"%e\" supported=\"%i\" transaction_id=\"%e\">%s</response>"
						, feature_name, (int)supported, transaction_id, setting);

	return SendResponse();
}

DEBUGGER_COMMAND(Debugger::feature_set)
{
	int err;
	char arg, *value;

	char *transaction_id, *feature_name, *feature_value;

	for(;;)
	{
		if (err = GetNextArg(aArgs, arg, value))
			return err;
		if (!arg)
			break;
		switch (arg)
		{
		case 'i': transaction_id = value; break;
		case 'n': feature_name = value; break;
		case 'v': feature_value = value; break;
		default:
			return DEBUGGER_E_INVALID_OPTIONS;
		}
	}

	if (!transaction_id || !feature_name || !feature_value)
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

	mResponseBuf.WriteF("<response command=\"feature_set\" feature=\"%e\" success=\"%i\" transaction_id=\"%e\"/>"
						, feature_name, (int)success, transaction_id);

	return SendResponse();
}

DEBUGGER_COMMAND(Debugger::stop)
{
	// See DEBUGGER_COMMAND_INIT_TRANSACTION_ID.
	if (!strncmp(aArgs, "-i ", 3) && !strstr(aArgs + 3, " -"))
		mContinuationTransactionId = aArgs + 3;
	else // Seems appropriate to ignore invalid args for stop command.
		mContinuationTransactionId = "";

	// Call g_script.TerminateApp instead of g_script.ExitApp to bypass OnExit subroutine.
	g_script.TerminateApp(EXIT_EXIT, 0); // This also causes this->Exit() to be called.
	
	// Should never be reached, but must be here to avoid a compile error:
	return DEBUGGER_E_INTERNAL_ERROR;
}

DEBUGGER_COMMAND(Debugger::breakpoint_set)
{
	int err;
	char arg, *value;
	
	char *type = NULL, state = BS_Enabled, *transaction_id = NULL, *filename = NULL;
	LineNumberType lineno = 0;
	bool temporary = false;

	for(;;)
	{
		if (err = GetNextArg(aArgs, arg, value))
			return err;
		if (!arg)
			break;
		switch (arg)
		{
		case 'i':
			transaction_id = value;
			break;

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

	if (!transaction_id || !type)
		return DEBUGGER_E_INVALID_OPTIONS;
	if (strcmp(type, "line")) // i.e. type != "line"
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

	Line *line = NULL;
	bool found_line = false;
	// Due to the introduction of expressions in static initializers, lines aren't necessarily in
	// line number order.  First determine if any static initializers match the requested lineno.
	// If not, use the first non-static line at or following that line number.

	if (g_script.mFirstStaticLine)
		for (line = g_script.mFirstStaticLine; ; line = line->mNextLine)
		{
			if (line->mFileIndex == file_index && line->mLineNumber == lineno)
			{
				found_line = true;
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
				// Use the first line of code at or after lineno, like Visual Studio.
				// To display the breakpoint correctly, an IDE should use breakpoint_get.
				found_line = true;
				break;
			}
	if (found_line)
	{
		if (!line->mBreakpoint)
			line->mBreakpoint = new Breakpoint();
		line->mBreakpoint->state = state;
		line->mBreakpoint->temporary = temporary;

		mResponseBuf.WriteF("<response command=\"breakpoint_set\" transaction_id=\"%e\" state=\"%s\" id=\"%i\"/>"
			, transaction_id, state ? "enabled" : "disabled", line->mBreakpoint->id);

		return SendResponse();
	}
	// There are no lines of code beginning at or after lineno.
	return DEBUGGER_E_BREAKPOINT_INVALID;
}

int Debugger::WriteBreakpointXml(Breakpoint *aBreakpoint, Line *aLine)
{
	return mResponseBuf.WriteF("<breakpoint id=\"%i\" type=\"line\" state=\"%s\" filename=\""
								, aBreakpoint->id, aBreakpoint->state ? "enabled" : "disabled")
								|| mResponseBuf.WriteFileURI(U4T(Line::sSourceFile[aLine->mFileIndex]))
		|| mResponseBuf.WriteF("\" lineno=\"%u\"/>", aLine->mLineNumber);
}

DEBUGGER_COMMAND(Debugger::breakpoint_get)
{
	int err;
	char arg, *value;
	
	char *transaction_id = NULL;
	int breakpoint_id = 0; // Breakpoint IDs begin at 1.

	for(;;)
	{
		if (err = GetNextArg(aArgs, arg, value))
			return err;
		if (!arg)
			break;
		switch (arg)
		{
		case 'i':	transaction_id = value;			break;
		case 'd':	breakpoint_id = atoi(value);	break;
		default:	return DEBUGGER_E_INVALID_OPTIONS;
		}
	}

	if (!transaction_id || !breakpoint_id)
		return DEBUGGER_E_INVALID_OPTIONS;

	Line *line;
	for (line = g_script.mFirstLine; line; line = line->mNextLine)
	{
		if (line->mBreakpoint && line->mBreakpoint->id == breakpoint_id)
		{
			mResponseBuf.WriteF("<response command=\"breakpoint_get\" transaction_id=\"%e\">", transaction_id);
			WriteBreakpointXml(line->mBreakpoint, line);
			mResponseBuf.Write("</response>");

			return SendResponse();
		}
	}

	return DEBUGGER_E_BREAKPOINT_NOT_FOUND;
}

DEBUGGER_COMMAND(Debugger::breakpoint_update)
{
	int err;
	char arg, *value;
	
	char *transaction_id = NULL;
	
	int breakpoint_id = 0; // Breakpoint IDs begin at 1.
	LineNumberType lineno = 0;
	char state = -1;

	for(;;)
	{
		if (err = GetNextArg(aArgs, arg, value))
			return err;
		if (!arg)
			break;
		switch (arg)
		{
		case 'i':
			transaction_id = value;
			break;

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

	if (!transaction_id || !breakpoint_id)
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

			return SendStandardResponse("breakpoint_update", transaction_id);
		}
	}

	return DEBUGGER_E_BREAKPOINT_NOT_FOUND;
}

DEBUGGER_COMMAND(Debugger::breakpoint_remove)
{
	int err;
	char arg, *value;
	
	char *transaction_id = NULL;
	int breakpoint_id = 0; // Breakpoint IDs begin at 1.

	for(;;)
	{
		if (err = GetNextArg(aArgs, arg, value))
			return err;
		if (!arg)
			break;
		switch (arg)
		{
		case 'i':	transaction_id = value;			break;
		case 'd':	breakpoint_id = atoi(value);	break;
		default:	return DEBUGGER_E_INVALID_OPTIONS;
		}
	}

	if (!transaction_id || !breakpoint_id)
		return DEBUGGER_E_INVALID_OPTIONS;

	Line *line;
	for (line = g_script.mFirstLine; line; line = line->mNextLine)
	{
		if (line->mBreakpoint && line->mBreakpoint->id == breakpoint_id)
		{
			delete line->mBreakpoint;
			line->mBreakpoint = NULL;

			return SendStandardResponse("breakpoint_remove", transaction_id);
		}
	}

	return DEBUGGER_E_BREAKPOINT_NOT_FOUND;
}

DEBUGGER_COMMAND(Debugger::breakpoint_list)
{
	DEBUGGER_COMMAND_INIT_TRANSACTION_ID;
	
	mResponseBuf.WriteF("<response command=\"breakpoint_list\" transaction_id=\"%e\">", transaction_id);
	
	Line *line;
	for (line = g_script.mFirstLine; line; line = line->mNextLine)
	{
		if (line->mBreakpoint)
		{
			WriteBreakpointXml(line->mBreakpoint, line);
		}
	}

	mResponseBuf.Write("</response>");

	return SendResponse();
}

DEBUGGER_COMMAND(Debugger::stack_depth)
{
	DEBUGGER_COMMAND_INIT_TRANSACTION_ID;

	mResponseBuf.WriteF("<response command=\"stack_depth\" depth=\"%i\" transaction_id=\"%e\"/>"
						, mStack.Depth(), transaction_id);

	return SendResponse();
}

DEBUGGER_COMMAND(Debugger::stack_get)
{
	int err;
	char arg, *value;

	char *transaction_id = NULL;
	int depth = -1;

	for(;;)
	{
		if (err = GetNextArg(aArgs, arg, value))
			return err;
		if (!arg)
			break;
		switch (arg)
		{
		case 'i':
			transaction_id = value;
			break;

		case 'd':
			depth = atoi(value);
			if (depth < 0 || depth >= mStack.Depth())
				return DEBUGGER_E_INVALID_STACK_DEPTH;
			break;

		default:
			return DEBUGGER_E_INVALID_OPTIONS;
		}
	}

	if (!transaction_id)
		return DEBUGGER_E_INVALID_OPTIONS;

	mResponseBuf.WriteF("<response command=\"stack_get\" transaction_id=\"%e\">", transaction_id);
	
	int level = 0;
	DbgStack::Entry *se;
	FileIndexType file_index;
	LineNumberType line_number;
	for (se = mStack.mTop; se >= mStack.mBottom; --se)
	{
		if (depth == -1 || depth == level)
		{
			if (se == mStack.mTop && mCurrLine) // See PreExecLine() for comments.
			{
				file_index = mCurrLine->mFileIndex;
				line_number = mCurrLine->mLineNumber;
			}
			else
			{
				file_index = se->line->mFileIndex;
				line_number = se->line->mLineNumber;
			}
			mResponseBuf.WriteF("<stack level=\"%i\" type=\"file\" filename=\"", level);
			mResponseBuf.WriteFileURI(U4T(Line::sSourceFile[file_index]));
			mResponseBuf.WriteF("\" lineno=\"%u\" where=\"", line_number);
			switch (se->type)
			{
			case DbgStack::SE_Thread:
				mResponseBuf.WriteF("%e (thread)", U4T(se->desc)); // %e to escape characters which desc may contain (e.g. "a & b" in hotkey name).
				break;
			case DbgStack::SE_Func:
				mResponseBuf.WriteF("%s()", U4T(se->func->mName)); // %s because function names should never contain characters which need escaping.
				break;
			case DbgStack::SE_Sub:
				mResponseBuf.WriteF("%e:", U4T(se->sub->mName)); // %e because label/hotkey names may contain almost anything.
				break;
			}
			mResponseBuf.Write("\"/>");
		}
		++level;
	}

	mResponseBuf.Write("</response>");

	return SendResponse();
}

DEBUGGER_COMMAND(Debugger::context_names)
{
	DEBUGGER_COMMAND_INIT_TRANSACTION_ID;

	mResponseBuf.WriteF("<response command=\"context_names\" transaction_id=\"%e\"><context name=\"Local\" id=\"0\"/><context name=\"Global\" id=\"1\"/></response>"
						, transaction_id);

	return SendResponse();
}

DEBUGGER_COMMAND(Debugger::context_get)
{
	int err;
	char arg, *value;

	char *transaction_id = NULL;
	int context_id = 0, depth = 0;

	for(;;)
	{
		if (err = GetNextArg(aArgs, arg, value))
			return err;
		if (!arg)
			break;
		switch (arg)
		{
		case 'i':	transaction_id = value;		break;
		case 'c':	context_id = atoi(value);	break;
		case 'd':	depth = atoi(value);		break;
		default:
			return DEBUGGER_E_INVALID_OPTIONS;
		}
	}

	if (!transaction_id)
		return DEBUGGER_E_INVALID_OPTIONS;

	// TODO: Support setting/retrieving variables at a given stack depth. See also property_get_or_value and property_set.
	if (depth != 0)
		return DEBUGGER_E_INVALID_STACK_DEPTH;

	Var **var, **var_end; // An array of pointers-to-var.
	
	// TODO: Include the lazy-var arrays for completeness. Low priority since lazy-var arrays are used only for 10001+ variables, and most conventional debugger interfaces would generally not be useful with that many variables.
	if (context_id == PC_Local)
	{
		if (g->CurrentFunc)
		{
			var = g->CurrentFunc->mVar;
			var_end = var + g->CurrentFunc->mVarCount;
		}
		else
			var_end = var = NULL;
	}
	else if (context_id == PC_Global)
	{
		var = g_script.mVar;
		var_end = var + g_script.mVarCount;
	}
	else
		return DEBUGGER_E_INVALID_CONTEXT;

	mResponseBuf.WriteF("<response command=\"context_get\" context=\"%i\" transaction_id=\"%e\">"
						, context_id, transaction_id);

	for ( ; var < var_end; ++var)
		WritePropertyXml(**var, mMaxPropertyData);

	mResponseBuf.Write("</response>");

	return SendResponse();
}

DEBUGGER_COMMAND(Debugger::typemap_get)
{
	DEBUGGER_COMMAND_INIT_TRANSACTION_ID;
	
	// Send a basic type-map with string = string.
	mResponseBuf.WriteF("<response command=\"typemap_get\" transaction_id=\"%e\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\">"
							"<map type=\"string\" name=\"string\" xsi:type=\"xsd:string\"/>"
							"<map type=\"int\" name=\"integer\" xsi:type=\"xsd:long\"/>"
							"<map type=\"float\" name=\"float\" xsi:type=\"xsd:double\"/>"
							"<map type=\"object\" name=\"object\"/>"
						"</response>"
						, transaction_id);

	return SendResponse();
}

DEBUGGER_COMMAND(Debugger::property_get)
{
	return property_get_or_value(aArgs, true);
}

DEBUGGER_COMMAND(Debugger::property_value)
{
	return property_get_or_value(aArgs, false);
}


int Debugger::WritePropertyXml(Var &aVar, int aMaxEncodedSize, int aPage)
{
	char facet[35]; // Alias Builtin Static ClipboardAll
	facet[0] = '\0';
	VarAttribType attrib;
	if (aVar.mType == VAR_ALIAS)
	{
		strcat(facet, "Alias ");
		attrib = aVar.mAliasFor->mAttrib;
	}
	else
		attrib = aVar.mAttrib;
	if (aVar.Type() == VAR_BUILTIN)
		strcat(facet, "Builtin ");
	if (aVar.IsStatic())
		strcat(facet, "Static ");
	if (aVar.IsBinaryClip())
		strcat(facet, "ClipboardAll ");
	if (facet[0] != '\0') // Remove the final space.
		facet[strlen(facet)-1] = '\0';

	if (attrib & VAR_ATTRIB_OBJECT)
	{
		CStringUTF8FromTChar name_buf(aVar.mName);
		return WritePropertyXml(aVar.Object(), name_buf.GetString(), name_buf, aPage, mMaxChildren, mMaxDepth, aMaxEncodedSize, facet);
	}

	char *type;
	if (attrib & VAR_ATTRIB_HAS_VALID_INT64)
		type = "integer";
	else if (attrib & VAR_ATTRIB_HAS_VALID_DOUBLE)
		type = "float";
	else
		type = "string";

	// Write as much as possible now to simplify calculation of required buffer space.
	// We can't write the data size yet as it needs to be UTF-8 encoded (see WritePropertyData).
	mResponseBuf.WriteF("<property name=\"%s\" fullname=\"%s\" type=\"%s\" facet=\"%s\" children=\"0\" encoding=\"base64\" size=\""
						, U4T(aVar.mName), U4T(aVar.mName), type, facet);

	WritePropertyData(aVar, aMaxEncodedSize);

	mResponseBuf.Write("</property>");
	
	return DEBUGGER_E_OK;
}

int Debugger::WritePropertyXml(IObject *aObject, const char *aName, CStringA &aNameBuf, int aPage, int aPageSize, int aDepthRemaining, int aMaxEncodedSize, char *aFacet)
{
	INT_PTR numchildren,
		i = aPageSize * aPage,
		j = aPageSize * (aPage + 1);
	size_t aNameBuf_marker = aNameBuf.GetLength();
	Object *objptr = dynamic_cast<Object *>(aObject);
	ComObject *comobjptr = dynamic_cast<ComObject *>(aObject);
	char *cp;
	int err;

	if (objptr)
		numchildren = objptr->mFieldCount;
	else if (comobjptr)
		numchildren = 2; // For simplicity, assume page == 0 and pagesize >= 2.
	else
		numchildren = 0;

	LPCSTR classname = typeid(*aObject).name();
	if (!strncmp(classname, "class ", 6))
		classname += 6;

	mResponseBuf.WriteF("<property name=\"%e\" fullname=\"%e\" type=\"object\" classname=\"%s\" address=\"%p\" facet=\"%s\" size=\"0\" page=\"%i\" pagesize=\"%i\" children=\"1\" numchildren=\"%p\">"
						, aName, aNameBuf.GetString(), classname, aObject, aFacet, aPage, aPageSize, numchildren);

	if (aDepthRemaining)
	{
		if (objptr)
		{
			Object &obj = *objptr;
			if (obj.mBase)
			{
				// Since this object has a "base", let it count as the first field.
				if (i == 0) // i.e. this is the first page.
				{
					cp = aNameBuf.GetBufferSetLength(aNameBuf_marker + 5) + aNameBuf_marker;
					strcpy(cp, ".base");
					aNameBuf.ReleaseBuffer();
					aName = "base";
					if (err = WritePropertyXml(obj.mBase, aName, aNameBuf, 0, aPageSize, aDepthRemaining - 1, aMaxEncodedSize))
						return err;
					// Now fall through and retrieve field[0] (unless aPageSize == 1).
				}
				// So 20..39 becomes 19..38 when there's a base object:
				else --i; 
				--j;
			}
			if (j > obj.mFieldCount)
				j = obj.mFieldCount;
			// For each field in the requested page...
			for ( ; i < j; ++i)
			{
				Object::FieldType &field = obj.mFields[i];

				// Append the key to the name buffer, attempting to maintain valid expression syntax.
				// String:	var.alphanumeric or var["non-alphanumeric"]
				// Integer:	var.123 or var[-123]
				// Object:	var[Object(address)]
				if (i >= obj.mKeyOffsetString) // String
				{
					LPTSTR tp;
					for (tp = field.key.s; cisalnum(*tp) || *tp == '_'; ++tp);
					bool use_dot = !*tp && tp != field.key.s; // If it got to the null-terminator, must be empty or alphanumeric.

					CStringUTF8FromTChar buf(field.key.s);
					LPCSTR key = buf.GetString();

					if (use_dot)
					{
						// Since this string is purely composed of alphanumeric characters and/or underscore,
						// it doesn't need any quote marks (imitating expression syntax) or escaped characters.
						cp = aNameBuf.GetBufferSetLength(aNameBuf_marker + buf.GetLength() + 1) + aNameBuf_marker;
						*cp++ = '.';
						strcpy(cp, key);
						// aNameBuf is released below.
					}
					else
					{
						// " must be replaced with "" as in expressions to remove ambiguity.
						char c;
						int extra = 4; // 4 for [""].  Also count double-quote marks:
						for ( ; *key; ++key) extra += *key=='"';
						cp = aNameBuf.GetBufferSetLength(aNameBuf_marker + buf.GetLength() + extra) + aNameBuf_marker;
						*cp++ = '[';
						*cp++ = '"';
						for (key = buf.GetString(); c = *key; ++key)
						{
							*cp++ = c;
							if (c == '"')
								*cp++ = '"'; // i.e. replace " with ""
						}
						strcpy(cp, "\"]");
						// aNameBuf is released below.
					}
				}
				else // INT_PTR or IObject*
				{
					INT_PTR key = field.key.i; // Also copies field.key.p (IObject*) via union.
					cp = aNameBuf.GetBufferSetLength(aNameBuf_marker + MAX_INTEGER_LENGTH + 10) + aNameBuf_marker; // +10 for "[Object()]"
					sprintf(cp, (i >= obj.mKeyOffsetObject) ? "[Object(%Ii)]" : key<0 ? "[%Ii]" : ".%Ii", field.key.i);
				}
				aNameBuf.ReleaseBuffer();
				aName = aNameBuf.GetString() + aNameBuf_marker;
				if (*aName == '.') // omit '.' but not []
					++aName;

				if (err = WritePropertyXml(field, aName, aNameBuf, aPageSize, aDepthRemaining, aMaxEncodedSize))
					return err;
			}
		}
		else if (comobjptr)
		{
			static LPCSTR sComObjNames[] = { "__Value", "__VarType" };
			LPCSTR name = aNameBuf.GetString();
			size_t value_len;
			TCHAR number_buf[MAX_NUMBER_SIZE];
			cp = (char *)number_buf;
			for ( ; i < 2 && i < j; ++i)
			{
				if (i == 0)
					_ui64toa(comobjptr->mVal64, cp, 10);
				else
					_ultoa(comobjptr->mVarType, cp, 10);
				value_len = strlen(cp);
				mResponseBuf.WriteF("<property name=\"%e\" fullname=\"%e.%e\" type=\"integer\" facet=\"\" children=\"0\" encoding=\"base64\" size=\"%p\">"
					, sComObjNames[i], name, sComObjNames[i], value_len);
				mResponseBuf.WriteEncodeBase64(cp, value_len);
				mResponseBuf.Write("</property>");
			}
		}
	}

	mResponseBuf.Write("</property>");

	aNameBuf.Truncate(aNameBuf_marker);
	return DEBUGGER_E_OK;
}

int Debugger::WritePropertyXml(Object::FieldType &aField, const char *aName, CStringA &aNameBuf, int aPageSize, int aDepthRemaining, int aMaxEncodedSize)
// This function has an equivalent WritePropertyData() for property_value, so maintain the two together.
{
	LPTSTR value;
	char *type;
	TCHAR number_buf[MAX_NUMBER_SIZE];

	switch (aField.symbol)
	{
	case SYM_OPERAND:
		value = aField.marker;
		type = "string";
		break;

	case SYM_INTEGER:
	case SYM_FLOAT:
		// The following tries to use the same methods to convert the number to a string as
		// the script would use when converting a variable's cached binary number to a string:
		if (aField.symbol == SYM_INTEGER)
		{
			ITOA64(aField.n_int64, number_buf);
			type = "integer";
		}
		else
		{
			sntprintf(number_buf, _countof(number_buf), g->FormatFloat, aField.n_double);
			type = "float";
		}
		value = number_buf;
		break;

	case SYM_OBJECT:
		// Recursively dump object.
		return WritePropertyXml(aField.object, aName, aNameBuf, 0, aPageSize, aDepthRemaining - 1, aMaxEncodedSize);

	default:
		ASSERT(FALSE);
	}
	// If we fell through, value and type have been set appropriately above.
	mResponseBuf.WriteF("<property name=\"%e\" fullname=\"%e\" type=\"%s\" facet=\"\" children=\"0\" encoding=\"base64\" size=\"", aName, aNameBuf.GetString(), type);
	int err;
	if (err = WritePropertyData(value, (int)_tcslen(value), aMaxEncodedSize))
		return err;
	return mResponseBuf.Write("</property>");
}

int Debugger::WritePropertyData(LPCTSTR aData, int aDataSize, int aMaxEncodedSize)
// Accepts a "native" string, converts it to UTF-8, base64-encodes it and writes
// the end of the property's size attribute followed by the base64-encoded data.
{
	int wide_size, value_size, space_needed; // Could use size_t, but WideCharToMultiByte is limited to int.
	char *value;
	int err;
	
#ifdef UNICODE
	LPCWSTR wide_value = aData;
	wide_size = aDataSize;
#else
	CStringWCharFromChar wide_buf(aData, aDataSize);
	LPCWSTR wide_value = wide_buf.GetString();
	wide_size = (int)wide_buf.GetLength();
#endif
	
	// Calculate the required buffer size to convert the value to UTF-8 without a null-terminator.
	// Actual conversion is not done until we've reserved some (shared) space in mResponseBuf for it.
	value_size = WideCharToMultiByte(CP_UTF8, 0, wide_value, wide_size, NULL, 0, NULL, NULL);

	if (value_size > aMaxEncodedSize) // Limit length of source data; see below.
		value_size = aMaxEncodedSize;
	// Calculate maximum length of base64-encoded data.
	// This should also ensure there is enough space to temporarily hold the raw value (aVar.Get()).
	space_needed = (value_size > 0) ? DEBUGGER_BASE64_ENCODED_SIZE(value_size) : sizeof(TCHAR);
	ASSERT(space_needed >= value_size + 1);
	
	// Reserve enough space for the data's length, "> and encoded data.
	if (err = mResponseBuf.ExpandIfNecessary(mResponseBuf.mDataUsed + space_needed + MAX_INTEGER_LENGTH + 2))
		return err;

	// Convert to UTF-8, using mResponseBuf temporarily.
	value = mResponseBuf.mData + mResponseBuf.mDataSize - space_needed;
	value_size = WideCharToMultiByte(CP_UTF8, 0, wide_value, wide_size, value, space_needed, NULL, NULL);

	// Now that we know the actual length for sure, write the end of the size attribute.
	mResponseBuf.WriteF("%u\">", value_size);

	// Limit length of value returned based on -m arg, client-requested max_data, defaulted value, etc.
	if (value_size > aMaxEncodedSize)
		value_size = aMaxEncodedSize;

	// Base64-encode and write the var data.
	return mResponseBuf.WriteEncodeBase64(value, (size_t)value_size, true);
}

int Debugger::WritePropertyData(Var &aVar, int aMaxEncodedSize)
{
	CString buf;
	LPTSTR value;
	int value_size;

	if (aVar.Type() == VAR_NORMAL)
	{
		value = aVar.Contents();
		value_size = (int)aVar.CharLength();
	}
	else
	{
		// In this case, allocating some memory to temporarily hold the var's value is unavoidable.
		if (value = buf.GetBufferSetLength(aVar.Get()))
			aVar.Get(value);
		else
			return DEBUGGER_E_INTERNAL_ERROR;
		CLOSE_CLIPBOARD_IF_OPEN; // Above may leave the clipboard open if aVar is Clipboard.
		value_size = (int)_tcslen(value);
	}

	return WritePropertyData(value, value_size, aMaxEncodedSize);
}

int Debugger::WritePropertyData(Object::FieldType &aField, int aMaxEncodedSize)
// Write object field: used only by property_value, so does not recursively dump objects.
// This function has an equivalent WritePropertyXml(), so maintain the two together.
{
	LPTSTR value;
	char *type;
	TCHAR number_buf[MAX_NUMBER_SIZE];

	switch (aField.symbol)
	{
	case SYM_OPERAND:
		value = aField.marker;
		type = "string";
		break;

	case SYM_INTEGER:
	case SYM_FLOAT:
		// The following tries to use the same methods to convert the number to a string as
		// the script would use when converting a variable's cached binary number to a string:
		if (aField.symbol == SYM_INTEGER)
		{
			ITOA64(aField.n_int64, number_buf);
			type = "integer";
		}
		else
		{
			sntprintf(number_buf, _countof(number_buf), g->FormatFloat, aField.n_double);
			type = "float";
		}
		value = number_buf;
		break;

	case SYM_OBJECT:
		value = _T("");
		type = "object";
		break;
	
	default:
		ASSERT(FALSE);
	}
	// If we fell through, value and type have been set appropriately above.
	return WritePropertyData(value, (int)_tcslen(value) + 1, aMaxEncodedSize);
}

int Debugger::ParsePropertyName(const char *aFullName, int aVarScope, bool aVarMustExist, Var *&aVar, Object::FieldType *&aField)
{
	CStringTCharFromUTF8 name_buf(aFullName);
	LPTSTR name = name_buf.GetBuffer();
	size_t name_length;
	TCHAR c, *name_end, *src, *dst;
	Var *var;
	SymbolType key_type;
	Object::KeyType key;
    Object::FieldType *field;
	Object::IndexType insert_pos;
	Object *obj;

	name_end = StrChrAny(name, _T(".["));
	if (name_end)
	{
		c = *name_end;
		*name_end = '\0'; // Temporarily terminate.
	}
	name_length = _tcslen(name);

	// Validate name for more accurate error-reporting.
	if (name_length > MAX_VAR_NAME_LENGTH || !Var::ValidateName(name, false, DISPLAY_NO_ERROR))
		return DEBUGGER_E_INVALID_OPTIONS;

	// If we're allowed to create variables
	if ( !aVarMustExist
		// or this variable doesn't exist
		|| !(var = g_script.FindVar(name, name_length, NULL, aVarScope))
			// but it is a built-in variable which hasn't been referenced yet:
			&& g_script.GetVarType(name) > (void*)VAR_LAST_TYPE )
		// Find or add the variable.
		var = g_script.FindOrAddVar(name, name_length, aVarScope);

	if (!var)
		return DEBUGGER_E_UNKNOWN_PROPERTY;

	if (!name_end)
	{
		// Just a variable name.
		aVar = var;
		aField = NULL;
		return DEBUGGER_E_OK;
	}

	if ( !var->HasObject() || !(obj = dynamic_cast<Object *>(var->Object())) )
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
			key_type = IsPureNumeric(name); // SYM_INTEGER or SYM_STRING.
			if (name_end)
			{
				c = *name_end; // Save this for the next iteration.
				*name_end = '\0';
			}
			//else there won't be a next iteration.
		}
		else
			return DEBUGGER_E_INVALID_OPTIONS;
		
		if (key_type == SYM_STRING)
			key.s = name;
		else // SYM_INTEGER or SYM_OBJECT
			key.i = Exp32or64(_ttoi,_ttoi64)(name);

		if ( !(field = obj->FindField(key_type, key, insert_pos)) )
		{
			if (!_tcsicmp(name, _T("base")) && (aVarMustExist || name_end && c)) // i.e. don't return the fake field to property_set since it would Release() the base object but not actually work.
			{
				// Since "base" doesn't usually correspond to an actual field, let it resolve
				// to a fake one for simplicity (dynamically allocated since we never want the
				// destructor to be called):
				static Object::FieldType *sBaseField = new Object::FieldType();
				if (obj->mBase)
				{
					sBaseField->symbol = SYM_OBJECT;
					sBaseField->object = obj->mBase;
				}
				else
				{
					sBaseField->symbol = SYM_OPERAND;
					sBaseField->marker = _T("");
					sBaseField->size = 0;
				}
				field = sBaseField;
				// If this is the end of 'name', sBaseField will be returned to our caller.
				// Otherwise the next iteration will either fail (mBase == NULL) or search
				// the base object's fields.
			}
			else
				return DEBUGGER_E_UNKNOWN_PROPERTY;
		}

		if (!name_end || !c)
		{
			// All done!
			aVar = NULL;
			aField = field;
			return DEBUGGER_E_OK;
		}

		if ( field->symbol != SYM_OBJECT || !(obj = dynamic_cast<Object *>(field->object)) )
			// No usable target object for the next iteration, therefore the property mustn't exist.
			return DEBUGGER_E_UNKNOWN_PROPERTY;

	} // infite loop.
}


int Debugger::property_get_or_value(char *aArgs, bool aIsPropertyGet)
{
	int err;
	char arg, *value;

	char *transaction_id = NULL, *name = NULL;
	int context_id = 0, depth = 0, page = 0;
	int max_data = aIsPropertyGet ? mMaxPropertyData : INT_MAX;

	for(;;)
	{
		if (err = GetNextArg(aArgs, arg, value))
			return err;
		if (!arg)
			break;
		switch (arg)
		{
		case 'i': transaction_id = value; break;
		// property long name
		case 'n': name = value; break;
		// context id - optional, default zero. see PropertyContextType enum.
		case 'c': context_id = atoi(value); break;
		// stack depth - optional, default zero
		case 'd': depth = atoi(value); break;
		// max data size - optional
		case 'm': max_data = atoi(value); break;
		// data page - optional
		case 'p': page = atoi(value); break;

		default:
			return DEBUGGER_E_INVALID_OPTIONS;
		}
	}

	if (!transaction_id || !name || max_data < 0)
		return DEBUGGER_E_INVALID_OPTIONS;

	// Currently only stack depth 0 (the top-most running function) is supported.
	if (depth != 0)
		return DEBUGGER_E_INVALID_STACK_DEPTH;

	int always_use;
	switch (context_id)
	{
	// It seems best to allow context id zero to retrieve either a local or global,
	// rather than requiring the IDE to check each context when looking up a variable.
	//case PC_Local:	always_use = ALWAYS_USE_LOCAL;	break;
	case PC_Local:	always_use = ALWAYS_PREFER_LOCAL;	break;
	case PC_Global:	always_use = ALWAYS_USE_GLOBAL;		break;
	default:
		return DEBUGGER_E_INVALID_CONTEXT;
	}

	Var *var;
	Object::FieldType *field;
	if (err = ParsePropertyName(name, always_use, true, var, field))
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
		
		mResponseBuf.WriteF("<response command=\"property_get\" transaction_id=\"%e\"><property name=\"%e\" fullname=\"%e\" type=\"undefined\" facet=\"\" size=\"0\" children=\"0\"/></response>"
							, transaction_id, name, name);

		return SendResponse();
	}
	//else var and field were set by the called function.

	if (aIsPropertyGet)
	{
		mResponseBuf.WriteF("<response command=\"property_get\" transaction_id=\"%e\">"
							, transaction_id);
		
		if (var)
			WritePropertyXml(*var, max_data, page);
		else
			WritePropertyXml(*field, name, CStringA(name), mMaxChildren, mMaxDepth, max_data);
	}
	else
	{
		mResponseBuf.WriteF("<response command=\"property_value\" transaction_id=\"%e\" encoding=\"base64\" size=\""
							, transaction_id);

		if (var)
			WritePropertyData(*var, max_data);
		else
			WritePropertyData(*field, max_data);
	}
	mResponseBuf.Write("</response>");

	return SendResponse();
}

DEBUGGER_COMMAND(Debugger::property_set)
{
	int err;
	char arg, *value;

	char *transaction_id = NULL, *name = NULL, *new_value = NULL, *type = "string";
	int context_id = 0, depth = 0;

	for(;;)
	{
		if (err = GetNextArg(aArgs, arg, value))
			return err;
		if (!arg)
			break;
		switch (arg)
		{
		case 'i': transaction_id = value; break;
		// property long name
		case 'n': name = value; break;
		// context id - optional, default zero. see PropertyContextType enum.
		case 'c': context_id = atoi(value); break;
		// stack depth - optional, default zero
		case 'd': depth = atoi(value); break;
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

	if (!transaction_id || !name || !new_value)
		return DEBUGGER_E_INVALID_OPTIONS;

	// Currently only stack depth 0 (the top-most running function) is supported.
	if (depth != 0)
		return DEBUGGER_E_INVALID_STACK_DEPTH;

	int always_use;
	switch (context_id)
	{
	// For consistency with property_get, create a local only if no global exists.
	//case PC_Local:	always_use = ALWAYS_USE_LOCAL;	break;
	case PC_Local:	always_use = ALWAYS_PREFER_LOCAL;	break;
	case PC_Global:	always_use = ALWAYS_USE_GLOBAL;		break;
	default:
		return DEBUGGER_E_INVALID_CONTEXT;
	}

	Var *var;
	Object::FieldType *field;
	if (err = ParsePropertyName(name, always_use, false, var, field))
		return err;
	
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
		StringUTF8ToTChar(new_value, val_buf, (int)Base64Decode(new_value, new_value));
		val.symbol = SYM_STRING;
		val.marker = (LPTSTR)val_buf.GetString();
	}

	bool success;
	if (var)
		success = !VAR_IS_READONLY(*var) && var->Assign(val); // Relies on shortcircuit boolean evaluation.
	else
		success = field->Assign(val);

	mResponseBuf.WriteF("<response command=\"property_set\" success=\"%i\" transaction_id=\"%e\"/>"
						, (int)success, transaction_id);

	return SendResponse();
}

DEBUGGER_COMMAND(Debugger::source)
{
	int err;
	char arg, *value;

	char *transaction_id = NULL, *filename = NULL;
	LineNumberType begin_line = 0, end_line = UINT_MAX;

	for(;;)
	{
		if (err = GetNextArg(aArgs, arg, value))
			return err;
		if (!arg)
			break;
		switch (arg)
		{
		case 'i': transaction_id = value; break;
		case 'b': begin_line = strtoul(value, NULL, 10); break;
		case 'e': end_line = strtoul(value, NULL, 10); break;
		case 'f': filename = value; break;
		default:
			return DEBUGGER_E_INVALID_OPTIONS;
		}
	}

	if (!transaction_id || !filename || begin_line > end_line)
		return DEBUGGER_E_INVALID_OPTIONS;

	int file_index;
	FILE *source_file = NULL;

	// Decode filename URI -> path, in-place.
	DecodeURI(filename);

	CStringTCharFromUTF8 filename_t(filename);

	// Ensure the file is actually a source file - i.e. don't let the debugger client retrieve any arbitrary file.
	for (file_index = 0; file_index < Line::sSourceFileCount; ++file_index)
	{
		if (!_tcsicmp(filename_t, Line::sSourceFile[file_index]))
		{
			// TODO: Change this section to support different file encodings, perhaps via TextFile class.

			if (!(source_file = _tfopen(filename_t, _T("rb"))))
				return DEBUGGER_E_CAN_NOT_OPEN_FILE;
			
			mResponseBuf.WriteF("<response command=\"source\" success=\"1\" transaction_id=\"%e\" encoding=\"base64\">"
								, transaction_id);

			if (begin_line == 0 && end_line == UINT_MAX)
			{	// RETURN THE ENTIRE FILE:
				long file_size;
				// Seek to end of file,
				if (fseek(source_file, 0, SEEK_END)
					// get file size
					|| (file_size = ftell(source_file)) < 1
					// and seek back to the beginning.
					|| fseek(source_file, 0, SEEK_SET))
					break; // fail.

				// Reserve some space in advance to read the raw file data into.
				if (err = mResponseBuf.WriteEncodeBase64(NULL, file_size))
					return err;

				// Read into the end of the response buffer.
				char *buf = mResponseBuf.mData + mResponseBuf.mDataSize - (file_size + 1);

				if (file_size != fread(buf, 1, file_size, source_file))
					break; // fail.

				// Base64-encode and write the file data into the response buffer.
				if (err = mResponseBuf.WriteEncodeBase64(buf, file_size))
					return err;
			}
			else
			{	// RETURN SPECIFIC LINES:
				char line_buf[LINE_SIZE + 2]; // May contain up to two characters of the previous line to simplify base64-encoding.
				size_t line_length;
				int line_remainder = 0;

				LineNumberType current_line = 0;
				
				while (fgets(line_buf + line_remainder, LINE_SIZE, source_file))
				{
					if (++current_line >= begin_line)
					{
						if (current_line > end_line)
							break; // done.
					
						if (line_length = strlen(line_buf))
						{
							// Encode in multiples of 3 characters to avoid inserting padding characters.
							line_remainder = line_length % 3;
							line_length -= line_remainder;

							if (line_length)
							{
								// Base64-encode and write this line and its trailing newline character into the response buffer.
								if (err = mResponseBuf.WriteEncodeBase64(line_buf, line_length))
									return err;
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
				}

				// Write any left-over characters (if line_remainder is 0, this does nothing).
				if (err = mResponseBuf.WriteEncodeBase64(line_buf, line_remainder))
					return err;

				if (!current_line || current_line < begin_line)
					break; // fail.
				// else if (current_line < end_line) -- just return what we can.
			}
			fclose(source_file);

			mResponseBuf.Write("</response>");

			return SendResponse();
		}
	}
	if (source_file)
		fclose(source_file);
	// If we got here, one of the following is true:
	//	- Something failed and used 'break'.
	//	- The requested file is not a known source file of this script.
	mResponseBuf.mDataUsed = 0;
	mResponseBuf.WriteF("<response command=\"source\" success=\"0\" transaction_id=\"%e\"/>"
						, transaction_id);
	
	return SendResponse();
}

int Debugger::redirect_std(char *aArgs, char *aCommandName)
{
	int err;
	char arg, *value;

	char *transaction_id = NULL;
	int new_mode = -1;

	for(;;)
	{
		if (err = GetNextArg(aArgs, arg, value))
			return err;
		if (!arg)
			break;
		switch (arg)
		{
		case 'i': transaction_id = value; break;
		case 'c': new_mode = atoi(value); break;
		default:
			return DEBUGGER_E_INVALID_OPTIONS;
		}
	}

	if (!transaction_id || new_mode < SR_Disabled || new_mode > SR_Redirect)
		return DEBUGGER_E_INVALID_OPTIONS;

	if (!stricmp(aCommandName, "stdout"))
		mStdOutMode = (StreamRedirectType)new_mode;
	else
		mStdErrMode = (StreamRedirectType)new_mode;

	mResponseBuf.WriteF("<response command=\"%s\" success=\"1\" transaction_id=\"%e\"/>"
						, aCommandName, transaction_id);

	return SendResponse();
}

DEBUGGER_COMMAND(Debugger::redirect_stdout)
{
	return redirect_std(aArgs, "stdout");
}

DEBUGGER_COMMAND(Debugger::redirect_stderr)
{
	return redirect_std(aArgs, "stderr");
}

//
// END DBGP COMMANDS
//

//
// DBGP STREAMS
//

int Debugger::WriteStreamPacket(LPCTSTR aText, LPCSTR aType)
{
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
	if (mStdOutMode != SR_Redirect) // i.e. SR_Disabled or SR_Copy
		return _fputts(aText, stdout) >= 0;
	return true;
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

int Debugger::SendContinuationResponse(char *aStatus, char *aReason)
{
	char *command;

	switch (mInternalState)
	{
	case DIS_StepInto:	command = "step_into";	break;
	case DIS_StepOver:	command = "step_over";	break;
	case DIS_StepOut:	command = "step_out";	break;
	//case DIS_Run:
	default:
		command = "run";
	}

	mResponseBuf.WriteF("<response command=\"%s\" status=\"%s\" reason=\"%s\" transaction_id=\"%e\"/>"
						, command, aStatus, aReason, mContinuationTransactionId);

	return SendResponse();
}

// Debugger::ReceiveCommand
//
// Receives a full command line into mCommandBuf. If part of the next command is
// received, it remains in the buffer until the next call to ReceiveCommand.
//
int Debugger::ReceiveCommand(int *aCommandLength)
{
	if (mSocket == INVALID_SOCKET)
		return FatalError(DEBUGGER_E_INTERNAL_ERROR);

	int err;
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
		if (mCommandBuf.mDataUsed == mCommandBuf.mDataSize && (err = mCommandBuf.Expand()))
			return err;

		// Receive and append data.
		int bytes_received = recv(mSocket, mCommandBuf.mData + mCommandBuf.mDataUsed, (int)(mCommandBuf.mDataSize - mCommandBuf.mDataUsed), 0);

		if (bytes_received == SOCKET_ERROR)
			return FatalError(DEBUGGER_E_INTERNAL_ERROR);

		mCommandBuf.mDataUsed += bytes_received;
	}
}

// Debugger::SendResponse
//
// Sends a response to a command, using mResponseBuf.mData as the message body.
//
int Debugger::SendResponse()
{
	int err;
	char response_header[DEBUGGER_RESPONSE_OVERHEAD];
	
	// Each message is prepended with a stringified integer representing the length of the XML data packet.
	Exp32or64(_itoa,_i64toa)(mResponseBuf.mDataUsed + DEBUGGER_XML_TAG_SIZE, response_header, 10);

	// The length and XML data are separated by a NULL byte.
	char *buf = strchr(response_header, '\0') + 1;

	// The XML document tag must always be present to provide XML version and encoding information.
	buf += sprintf(buf, "%s", DEBUGGER_XML_TAG);

	// Send the response header.
	if (SOCKET_ERROR == send(mSocket, response_header, (int)(buf - response_header), 0))
		return FatalError(DEBUGGER_E_INTERNAL_ERROR);

	// The messages sent by the debugger engine must always be NULL terminated.
	if (err = mResponseBuf.Write("\0", 1))
	{
		return err;
	}

	// Send the message body.
	if (SOCKET_ERROR == send(mSocket, mResponseBuf.mData, (int)mResponseBuf.mDataUsed, 0))
		return FatalError(DEBUGGER_E_INTERNAL_ERROR);

	mResponseBuf.mDataUsed = 0;
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
		return FatalError(DEBUGGER_E_INTERNAL_ERROR);
	
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
					g_script.ExitApp(EXIT_ERROR, _T(""));
					// Above should always exit, but if it doesn't, fall through to the next case:
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

				// Write init message.
				mResponseBuf.WriteF("<init appid=\"" AHK_NAME "\" ide_key=\"%e\" session=\"%e\" thread=\"%u\" parent=\"\" language=\"" DEBUGGER_LANG_NAME "\" protocol_version=\"1.0\" fileuri=\""
					, ide_key.GetString(), session.GetString(), GetCurrentThreadId());
				mResponseBuf.WriteFileURI(U4T(g_script.mFileSpec));
				mResponseBuf.Write("\"/>");

				if (SendResponse() == DEBUGGER_E_OK)
				{
					return DEBUGGER_E_OK;
				}

				mSocket = INVALID_SOCKET; // Don't want FatalError() to attempt a second closesocket().
			}
		}

		closesocket(s);
	}

	WSACleanup();
	return FatalError(DEBUGGER_E_INTERNAL_ERROR, DEBUGGER_ERR_FAILEDTOCONNECT DEBUGGER_ERR_DISCONNECT_PROMPT);
}

// Debugger::Disconnect
//
// Disconnect the debugger UI.
//
int Debugger::Disconnect()
{
	if (mSocket != INVALID_SOCKET)
	{
		closesocket(mSocket);
		mSocket = INVALID_SOCKET;
		WSACleanup();
	}
	// These are reset in case we re-attach to the debugger client later:
	mCommandBuf.mDataUsed = 0;
	mResponseBuf.mDataUsed = 0;
	mStdOutMode = SR_Disabled;
	mStdErrMode = SR_Disabled;
	mInternalState = DIS_Starting;
	return DEBUGGER_E_OK;
}

// Debugger::Exit
//
// Gracefully end debug session.  Called on script exit.  Also called by "detach" DBGp command.
//
void Debugger::Exit(ExitReasons aExitReason)
{
	if (mSocket == INVALID_SOCKET)
		return;
	// Don't care if it fails as we may be exiting due to a previous failure.
	SendContinuationResponse("stopped", aExitReason == EXIT_ERROR ? "error" : "ok");
	Disconnect();
}

// Debugger::GetNextArg
//
// Used to parse args for received DBGp commands. See Debugger.h for more info.
//
int Debugger::GetNextArg(char *&aArgs, char &aArg, char *&aValue)
{
	aArg = '\0';

	if (!aArgs)
	{
		aValue = NULL;
		return DEBUGGER_E_OK;
	}

	// Command line args must begin with an -arg.
	if (*aArgs != '-')
		return DEBUGGER_E_PARSE_ERROR;

	// Consume the hyphen.
	++aArgs;

	if (*aArgs == '-')
	{	// The rest of the command line is base64-encoded data.
		aArg = '-';
		aValue = aArgs + 1;
		if (*aValue == ' ')
			++aValue;
		aArgs = NULL;
		return DEBUGGER_E_OK;
	}
	
	aArg = *aArgs;
	
	if (aArg == '\0')
		return DEBUGGER_E_PARSE_ERROR;

	++aArgs;
	if (*aArgs == ' ')
	{
		++aArgs;
		if (*aArgs == '-')
		{	// Reached the next arg; i.e. this arg has no value.
			aValue = NULL;
			//return DEBUGGER_E_OK;
			// Currently every arg of every command requires a value.
			return DEBUGGER_E_INVALID_OPTIONS;
		}
	}
	// Else if *aArgs == '-', since there was no space, assume it is
	// the value of this arg rather than the beginning of the next arg.
	if (*aArgs == '\0')
	{	// End of string following "-a" or "-a ".
		aValue = NULL;
		aArgs = NULL;
		return DEBUGGER_E_OK;
	}

	aValue = aArgs;

	// Find where this arg's value ends and the next arg begins.
	char *next_arg = strstr(aArgs, " -");

	if (next_arg != NULL)
	{
		// Terminate aValue.
		*next_arg = '\0';
		
		// Set aArgs to the hyphen of the next arg, ready for the next call.
		aArgs = next_arg + 1;
	}
	else
	{
		// This was the last arg.
		aArgs = NULL;
	}

	return DEBUGGER_E_OK;
}

int Debugger::FatalError(int aErrorCode, LPCTSTR aMessage)
{
	g_Debugger.Disconnect();

	if (IDNO == MessageBox(g_hWnd, aMessage, g_script.mFileSpec, MB_YESNO | MB_ICONSTOP | MB_SETFOREGROUND | MB_APPLMODAL))
	{
		// The following will exit even if the OnExit subroutine does not use ExitApp:
		g_script.ExitApp(EXIT_ERROR, _T(""));
	}
	return aErrorCode;
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
	int err;

	if (aDataSize == -1)
		aDataSize = strlen(aData);

	if (aDataSize == 0)
		return DEBUGGER_E_OK;

	if (err = ExpandIfNecessary(mDataUsed + aDataSize))
		return err;

	memcpy(mData + mDataUsed, aData, aDataSize);
	mDataUsed += aDataSize;
	return DEBUGGER_E_OK;
}

// Write formatted data into the buffer. Supports %s (char*), %e (char*, "&'<> escaped), %i (int), %u (unsigned int), %p (UINT_PTR).
int Debugger::Buffer::WriteF(const char *aFormat, ...)
{
	int i, err;
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

		va_end(vl);

		if (i == 0)
			if (err = ExpandIfNecessary(mDataUsed + len))
				return err;
	} // for (len = 0, i = 0; i < 2; ++i)

	return DEBUGGER_E_OK;
}

// Convert a file path to a URI and write it to the buffer.
int Debugger::Buffer::WriteFileURI(const char *aPath)
{
	int err, c, len = 9; // 8 for "file:///", 1 for '\0' (written by sprintf()).

	// Calculate required buffer size for path after encoding.
	for (const char *ptr = aPath; c = *ptr; ++ptr)
	{
		if (cisalnum(c) || strchr("-_.!~*'()/\\", c))
			++len;
		else
			len += 3;
	}

	// Ensure the buffer contains enough space.
	if (err = ExpandIfNecessary(mDataUsed + len))
		return err;

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
	int err;
	if (aInputSize)
	{
		if (!aSkipBufferSizeCheck)
		{
			// Ensure required buffer space is available.
			if (err = ExpandIfNecessary(mDataUsed + DEBUGGER_BASE64_ENCODED_SIZE(aInputSize)))
				return err;
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
	size_t new_size;
	for (new_size = mDataSize ? mDataSize : DEBUGGER_INITIAL_BUFFER_SIZE
		; new_size < aRequiredSize
		; new_size *= 2);

	if (new_size > mDataSize)
	{
		// For simplicity, this preserves all of mData not just the first mDataUsed bytes.  Some sections may rely on this.
		char *new_data = (char*)realloc(mData, new_size);

		if (new_data == NULL)
			// Note: FatalError() calls Disconnect(), which "clears" mResponseBuf.
			return FatalError(DEBUGGER_E_INTERNAL_ERROR);

		mData = new_data;
		mDataSize = new_size;
	}
	return DEBUGGER_E_OK;
}

// Remove data from the front of the buffer (i.e. after it is processed).
void Debugger::Buffer::Remove(size_t aDataSize)
{
	// Move remaining data to the front of the buffer.
	if (aDataSize < mDataUsed)
		memmove(mData, mData + aDataSize, mDataUsed - aDataSize);
	mDataUsed -= aDataSize;
}

#endif


// Helper for Var::ToText().

LPTSTR Var::ObjectToText(LPTSTR aBuf, int aBufSize)
{
	LPTSTR aBuf_orig = aBuf;
	aBuf += sntprintf(aBuf, aBufSize, _T("%s[Object]: 0x%p"), mName, mObject);
	if (ComObject *cobj = dynamic_cast<ComObject *>(mObject))
		aBuf += sntprintf(aBuf, BUF_SPACE_REMAINING, _T(" <= ComObject(0x%04hX, 0x%I64X)"), cobj->mVarType, cobj->mVal64);
	return aBuf;
}
