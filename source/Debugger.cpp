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

#ifdef SCRIPT_DEBUG

#include "defines.h"
#include <Ws2tcpip.h>
#include <Wspiapi.h> // for getaddrinfo() on versions of Windows earlier than XP.
#include <stdarg.h>
//#include "Debugger.h" // included by globaldata.h
#include "globaldata.h" // for access to many global vars

Debugger g_Debugger;
char *g_DebuggerHost = NULL;
char *g_DebuggerPort = NULL;

// The first breakpoint uses sMaxId + 1. Don't change this without also changing breakpoint_remove.
int Breakpoint::sMaxId = 0;


// aLine is about to execute.
int Debugger::PreExecLine(Line *aLine)
{
	Breakpoint *&bp = aLine->mBreakpoint;

	mStackTop->line = aLine;

	if (bp && bp->state == BS_Enabled)
	{
		if (bp->temporary)
		{
			delete bp;
			bp = NULL;
		}
		return ProcessCommands();
	}
	else if ((mInternalState == DIS_StepInto
			|| mInternalState == DIS_StepOut && mStackDepth < mContinuationDepth
			|| mInternalState == DIS_StepOver && mStackDepth <= mContinuationDepth)
			// L31: The following check is no longer done because a) ACT_BLOCK_BEGIN belonging to an IF/ELSE/LOOP is now skipped; and b) allowing step to break at ACT_BLOCK_END makes program flow a little easier to follow in some cases.
			// L40: Stepping on braces seems more bothersome than helpful. Although IF/ELSE/LOOP skips its block-begin, standalone/function-body block-begin still gets here.
			&& aLine->mActionType != ACT_BLOCK_BEGIN && (aLine->mActionType != ACT_BLOCK_END || aLine->mAttribute) // For now, ignore { and }, except for function-end.
			&& aLine->mLineNumber) // Some scripts (i.e. LowLevel/code.ahk) use mLineNumber==0 to indicate the Line has been generated and injected by the script.
	{
		return ProcessCommands();
	}
	
	return DEBUGGER_E_OK;
}

int Debugger::StackPush(StackEntry *aEntry)
{
	mStackTop->upper = aEntry;
	aEntry->lower = mStackTop;
	mStackTop = aEntry;
	++mStackDepth;
	return DEBUGGER_E_OK;
}

int Debugger::StackPop()
{
	if (mStackTop != mStack)
	{
		mStackTop = mStackTop->lower;
		mStackTop->upper = NULL;
		--mStackDepth;
		// Stepping out or over from the auto-execute section requires the script to reach
		// a depth of < or <= 1 to break. This will never happen, since new threads start
		// at depth 2 - 1 for the thread, 1 for the sub or function it executes.
		// As a workaround, if all threads finished, step_out/over will act like step_into.
		if (mStackDepth == 0)
			mContinuationDepth = INT_MAX;
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
					mContinuationDepth = mStackDepth;
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

// For use by numerous commands:
#define DEBUGGER_EXPAND_RESPONSEBUF_IF_NECESSARY(space_needed) \
	if (mResponseBuf.mDataUsed + space_needed > mResponseBuf.mDataSize && (err = mResponseBuf.Expand(mResponseBuf.mDataUsed + space_needed))) \
		return err
		// If an error occurs, FatalError() calls Disconnect(), which clears mResponseBuf.

// Calculate base64-encoded size of data, including NULL terminator. org_size must be > 0.
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
			setting = NAME_P;
		else if (supported = !strcmp(feature_name + 9, "version"))
			setting = NAME_VERSION;
	} else if (supported = !strcmp(feature_name, "encoding"))
		setting = "windows-1252";
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
	// Not supported: max_children, max_depth - not applicable until we have arrays/objects.
	else if (!strcmp(feature_name, "max_data"))
		setting = _ultoa(mMaxPropertyData, (char*)_alloca(MAX_NUMBER_SIZE), 10);
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

	if (!strcmp(feature_name, "max_data"))
	{
		mMaxPropertyData = strtoul(feature_value, NULL, 10);
		if (mMaxPropertyData == 0)
			mMaxPropertyData = MAXDWORD;
		success = true;
	}

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

		// Find the specified source file.
		for (file_index = 0; file_index < Line::sSourceFileCount; ++file_index)
			if (!lstrcmpi(filename, Line::sSourceFile[file_index]))
				break;

		if (file_index >= Line::sSourceFileCount)
			return DEBUGGER_E_BREAKPOINT_INVALID;
	}

	Line *line;
	for (line = g_script.mFirstLine; line; line = line->mNextLine)
	{
		if (line->mFileIndex == file_index)
		{
			// Use the first line of code at or after lineno, like Visual Studio.
			// To display the breakpoint correctly, an IDE should use breakpoint_get.
			if (line->mLineNumber >= lineno)
			{
				if (!line->mBreakpoint)
					line->mBreakpoint = new Breakpoint();
				line->mBreakpoint->state = state;
				line->mBreakpoint->temporary = temporary;

				mResponseBuf.WriteF("<response command=\"breakpoint_set\" transaction_id=\"%e\" state=\"%s\" id=\"%i\"/>"
					, transaction_id, state ? "enabled" : "disabled", line->mBreakpoint->id);

				return SendResponse();
			}
			//else if (line->mLineNumber > lineno)
			//{
			//	// Passed the breakpoint line without finding any code.
			//	return DEBUGGER_E_BREAKPOINT_NO_CODE;
			//}
		}
	}
	// lineno is beyond the end of the file.
	return DEBUGGER_E_BREAKPOINT_INVALID;
}

int Debugger::WriteBreakpointXml(Breakpoint *aBreakpoint, Line *aLine)
{
	return mResponseBuf.WriteF("<breakpoint id=\"%i\" type=\"line\" state=\"%s\" filename=\""
								, aBreakpoint->id, aBreakpoint->state ? "enabled" : "disabled")
		|| mResponseBuf.WriteFileURI(Line::sSourceFile[aLine->mFileIndex])
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
						, mStackDepth, transaction_id);

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
			if (depth < 0 || depth >= mStackDepth)
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
	for (StackEntry *se = mStackTop; se != mStack; se = se->lower)
	{
		if (depth == -1 || depth == level)
		{
			mResponseBuf.WriteF("<stack level=\"%i\" type=\"file\" filename=\"", level);
			mResponseBuf.WriteFileURI(Line::sSourceFile[se->line->mFileIndex]);
			mResponseBuf.WriteF("\" lineno=\"%u\" where=\"", se->line->mLineNumber);
			switch (se->type)
			{
			case SE_Thread:
				mResponseBuf.WriteF("%e (thread)", se->desc); // %e to escape characters which desc may contain (e.g. "a & b" in hotkey name).
				break;
			case SE_Func:
				mResponseBuf.WriteF("%s()", se->func->mName); // %s because function names should never contain characters which need escaping.
				break;
			case SE_Sub:
				mResponseBuf.WriteF("%e:", se->sub->mName); // %e because label/hotkey names may contain almost anything.
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
		WritePropertyXml(*var, mMaxPropertyData);

	mResponseBuf.Write("</response>");

	return SendResponse();
}

DEBUGGER_COMMAND(Debugger::typemap_get)
{
	DEBUGGER_COMMAND_INIT_TRANSACTION_ID;
	
	// Send a basic type-map with string = string.
	mResponseBuf.WriteF("<response command=\"typemap_get\" transaction_id=\"%e\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\"><map type=\"string\" name=\"string\" xsi:type=\"xsd:string\"/></response>"
						, transaction_id);

	return SendResponse();
}

int Debugger::WritePropertyXml(Var *aVar, VarSizeType aMaxData)
{
	char facet[35]; // Alias Builtin Static ClipboardAll
	facet[0] = '\0';
	if (aVar->RealType() == VAR_ALIAS)
		strcat(facet, "Alias ");
	if (aVar->RealType() == VAR_BUILTIN)
		strcat(facet, "Builtin ");
	if (aVar->IsStatic())
		strcat(facet, "Static ");
	if (aVar->IsBinaryClip())
		strcat(facet, "ClipboardAll ");
	if (facet[0] != '\0') // Remove the final space.
		facet[strlen(facet)-1] = '\0';

	// Write as much as possible now to simplify calculation of required buffer space.
	// We can't write the var length yet as it may not be accurate.
	mResponseBuf.WriteF("<property name=\"%s\" fullname=\"%s\" type=\"string\" facet=\"%s\" children=\"0\" encoding=\"base64\" size=\""
						, aVar->mName, aVar->mName, facet);

	WriteVarSizeAndData(aVar, aMaxData);

	mResponseBuf.Write("</property>");
	
	return DEBUGGER_E_OK;
}

int Debugger::WriteVarSizeAndData(Var *aVar, VarSizeType aMaxData)
{
	VarSizeType length, min_space_needed = 0;

	if (aVar->Type() == VAR_NORMAL)
	{
		// Use Capacity() instead of Length() in case the script stores binary
		// data in the variable and the IDE has some way to display it.
		// UPDATE: Because we say this property is a "string", the IDE is likely to
		// display it as such, and may display the "garbage" after the null-terminator.
		//length = aVar->Capacity();
		length = aVar->Length();
	}
	else
	{
		// Get() returns the maximum length, not always the actual length of the value.
		// We need at least this much buffer space to retrieve the variable data.
		min_space_needed = (length = aVar->Get()) + 1;
	}

	// Calculate maximum length of base64-encoded variable data.
	DWORD length_base64 = length ? DEBUGGER_BASE64_ENCODED_SIZE(min(length,aMaxData)) : 0;
	
	// Reserve at least enough space to write the var length"> and data.
	DWORD space_needed = max(length_base64, min_space_needed) + 12; // 12 = max VarSizeType length + ">
	int err;
	DEBUGGER_EXPAND_RESPONSEBUF_IF_NECESSARY(space_needed);
	// Above may return on failure.

	char *contents;

	if (aVar->Type() == VAR_NORMAL)
	{
		// Copy and base64-encode directly from the variable.
		contents = aVar->Contents();
	}
	else
	{
		// Copy the variable data to the *end* of the buffer so that we can
		// base64-encode it without allocating more memory to hold the input data.
		contents = mResponseBuf.mData + mResponseBuf.mDataSize - min_space_needed;
		// Up to this point, length may be inaccurate.
		length = aVar->Get(contents);
	}

	// Now that we know the actual length, write the end of the size attribute.
	mResponseBuf.WriteF("%u\">", length);

	// Limit length of data returned based on -m arg, max_data, defaulted value, etc.
	if (length > aMaxData)
		length = aMaxData;

	// Base64-encode and write the var data.
	mResponseBuf.mDataUsed += Base64Encode(mResponseBuf.mData + mResponseBuf.mDataUsed, contents, length);
	
	return DEBUGGER_E_OK;
}

DEBUGGER_COMMAND(Debugger::property_get)
{
	return property_get_or_value(aArgs, true);
}

DEBUGGER_COMMAND(Debugger::property_value)
{
	return property_get_or_value(aArgs, false);
}

int Debugger::property_get_or_value(char *aArgs, bool aIsPropertyGet)
{
	int err;
	char arg, *value;

	char *transaction_id = NULL, *name = NULL;
	int context_id = 0, depth = 0;
	VarSizeType max_data = aIsPropertyGet ? mMaxPropertyData : VARSIZE_MAX;

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
		case 'm': max_data = strtoul(value, NULL, 10); break;

		default:
			return DEBUGGER_E_INVALID_OPTIONS;
		}
	}

	if (!transaction_id || !name)
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

	Var *var = g_script.FindVar(name, 0, NULL, always_use);

	if (!var) // Var not found.
	{
		if (!aIsPropertyGet)
			return DEBUGGER_E_UNKNOWN_PROPERTY;

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
		
		// If the debugger client includes XML-reserved characters in the property name,
		// the response XML will be badly formatted. Temporary workaround follows:
		if (strlen(name) > MAX_VAR_NAME_LENGTH || !Var::ValidateName(name, true, DISPLAY_NO_ERROR))
			name = "(invalid)";

		mResponseBuf.WriteF("<response command=\"property_get\" transaction_id=\"%e\"><property name=\"%e\" fullname=\"%e\" type=\"undefined\" facet=\"\" size=\"0\" children=\"0\"/></response>"
							, transaction_id, name, name);

		return SendResponse();
	}

	if (aIsPropertyGet)
	{
		mResponseBuf.WriteF("<response command=\"property_get\" transaction_id=\"%e\">"
							, transaction_id);
		
		WritePropertyXml(var, max_data);
	}
	else
	{
		mResponseBuf.WriteF("<response command=\"property_value\" transaction_id=\"%e\" encoding=\"base64\" size=\""
							, transaction_id);

		WriteVarSizeAndData(var, max_data);
	}
	mResponseBuf.Write("</response>");

	return SendResponse();
}

DEBUGGER_COMMAND(Debugger::property_set)
{
	int err;
	char arg, *value;

	char *transaction_id = NULL, *name = NULL, *new_value = NULL;
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

	Var *var = g_script.FindOrAddVar(name, strlen(name), always_use);

	bool success = var && var->Assign(new_value, Base64Decode(new_value, new_value));

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
	DWORD space_needed;

	// Decode filename URI -> path, in-place.
	DecodeURI(filename);

	// Ensure the file is actually a source file - i.e. don't let the debugger client retrieve any arbitrary file.
	for (file_index = 0; file_index < Line::sSourceFileCount; ++file_index)
	{
		if (!lstrcmpi(filename, Line::sSourceFile[file_index]))
		{
			if (!(source_file = fopen(filename, "rb")))
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

				// Calculate base64-encoded length. Relies on above ensuring file_size > 0, since below does file_size-1 and file_size is unsigned.
				space_needed = DEBUGGER_BASE64_ENCODED_SIZE(file_size);

				DEBUGGER_EXPAND_RESPONSEBUF_IF_NECESSARY(space_needed);
				// Above may return on failure.

				// Read into the end of the response buffer.
				char *buf = mResponseBuf.mData + mResponseBuf.mDataSize - (file_size + 1);

				if (file_size != fread(buf, 1, file_size, source_file))
					break; // fail.

				// Base64-encode and write the file data into the response buffer.
				mResponseBuf.mDataUsed += Base64Encode(mResponseBuf.mData + mResponseBuf.mDataUsed, buf, file_size);
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

							space_needed = DEBUGGER_BASE64_ENCODED_SIZE(line_length) + 1;

							DEBUGGER_EXPAND_RESPONSEBUF_IF_NECESSARY(space_needed);
							// Above may return on failure.

							// Base64-encode and write this line and its trailing newline character into the response buffer.
							mResponseBuf.mDataUsed += Base64Encode(mResponseBuf.mData + mResponseBuf.mDataUsed, line_buf, line_length);

							if (line_remainder) // 1 or 2.
							{
								line_buf[0] = line_buf[line_length];
								if (line_remainder > 1)
									line_buf[1] = line_buf[line_length + 1];
							}
						}
					}
				}

				if (line_remainder)
					mResponseBuf.mDataUsed += Base64Encode(mResponseBuf.mData + mResponseBuf.mDataUsed, line_buf, line_remainder);

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

	char *transaction_id = NULL, *new_mode = NULL;

	for(;;)
	{
		if (err = GetNextArg(aArgs, arg, value))
			return err;
		if (!arg)
			break;
		switch (arg)
		{
		case 'i': transaction_id = value; break;
		case 'c': new_mode = value; break;
		default:
			return DEBUGGER_E_INVALID_OPTIONS;
		}
	}

	if (!transaction_id || !new_mode)
		return DEBUGGER_E_INVALID_OPTIONS;

	// TODO: Support redirecting of stdout, or at least the * (stdout) mode of FileAppend.
	// TODO: Support reporting of non-critical errors through "stderr" redirection. Alternately redirect OutputDebug.
	mResponseBuf.WriteF("<response command=\"%s\" success=\"%s\" transaction_id=\"%e\"/>"
						, aCommandName, *new_mode=='0' ? "1" : "0", transaction_id);

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
		int bytes_received = recv(mSocket, mCommandBuf.mData + mCommandBuf.mDataUsed, mCommandBuf.mDataSize - mCommandBuf.mDataUsed, 0);

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
	_itoa(mResponseBuf.mDataUsed + DEBUGGER_XML_TAG_SIZE, response_header, 10);

	// The length and XML data are separated by a NULL byte.
	char *buf = strchr(response_header, '\0') + 1;

	// The XML document tag must always be present to provide XML version and encoding information.
	buf += sprintf(buf, "%s", DEBUGGER_XML_TAG);

	// Send the response header.
	if (SOCKET_ERROR == send(mSocket, response_header, buf - response_header, 0))
		return FatalError(DEBUGGER_E_INTERNAL_ERROR);

	// The messages sent by the debugger engine must always be NULL terminated.
	if (err = mResponseBuf.Write("\0", 1))
	{
		return err;
	}

	// Send the message body.
	if (SOCKET_ERROR == send(mSocket, mResponseBuf.mData, mResponseBuf.mDataUsed, 0))
		return FatalError(DEBUGGER_E_INTERNAL_ERROR);

	mResponseBuf.mDataUsed = 0;
	return DEBUGGER_E_OK;
}

// Debugger::Connect
//
// Connect to a debugger UI. Returns a Winsock error code on failure, otherwise 0.
//
int Debugger::Connect(char *aAddress, char *aPort)
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
			err = connect(s, res->ai_addr, res->ai_addrlen);
			
			freeaddrinfo(res);
			
			if (err == 0)
			{
				mSocket = s;

				char *ide_key = "", *session = "";
				int size;

				if (size = GetEnvironmentVariable("DBGP_IDEKEY", NULL, 0))
				{
					ide_key = (char*)_alloca(size);
					GetEnvironmentVariable("DBGP_IDEKEY", ide_key, size);
				}
				if (size = GetEnvironmentVariable("DBGP_COOKIE", NULL, 0))
				{
					session = (char*)_alloca(size);
					GetEnvironmentVariable("DBGP_COOKIE", session, size);
				}

				// Write init message.
				mResponseBuf.WriteF("<init appid=\"" NAME_P "\" ide_key=\"%e\" session=\"%e\" thread=\"%u\" parent=\"\" language=\"" NAME_P "\" protocol_version=\"1.0\" fileuri=\""
					, ide_key, session, GetCurrentThreadId());
				mResponseBuf.WriteFileURI(g_script.mFileSpec);
				mResponseBuf.Write("\"/>");

				if (SendResponse() == DEBUGGER_E_OK)
				{
					return DEBUGGER_E_OK;
				}
			}
		}

		closesocket(s);
	}

	WSACleanup();
	return FatalError(DEBUGGER_E_INTERNAL_ERROR);
}

// Debugger::Disconnect
//
// Disconnect the debugger UI. Returns a Winsock error code on failure, otherwise 0.
//
int Debugger::Disconnect()
{
	if (mSocket != INVALID_SOCKET)
	{
		closesocket(mSocket);
		//if (closesocket(mSocket) == SOCKET_ERROR)
		//{
		//	return DEBUGGER_E_INTERNAL_ERROR;
		//}
		mSocket = INVALID_SOCKET;
		WSACleanup();
	}
	mCommandBuf.mDataUsed = 0;
	mResponseBuf.mDataUsed = 0;
	return DEBUGGER_E_OK;
}

// Debugger::Exit
//
// Call when exiting to gracefully end debug session.
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

int Debugger::FatalError(int aErrorCode, char *aMessage)
{
	g_Debugger.Disconnect();

	if (!aMessage)
		aMessage = DEBUGGER_ERR_INTERNAL DEBUGGER_ERR_DISCONNECT_PROMPT;
	
	if (IDNO == MessageBox(g_hWnd, aMessage, g_script.mFileSpec, MB_YESNO | MB_ICONSTOP | MB_SETFOREGROUND | MB_APPLMODAL))
	{
		// The following will exit even if the OnExit subroutine does not use ExitApp:
		g_script.ExitApp(EXIT_ERROR, "");
	}
	return aErrorCode;
}

char *Debugger::sBase64Chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";

#define BINARY_TO_BASE64_CHAR(b) (sBase64Chars[(b) & 63])
#define BASE64_CHAR_TO_BINARY(q) (strchr(sBase64Chars, q)-sBase64Chars)

// Encode base 64 data.
int Debugger::Base64Encode(char *aBuf, const char *aInput, int aInputSize)
{
	int buffer, i, len = 0;

	if (aInputSize == -1)
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
int Debugger::Base64Decode(char *aBuf, const char *aInput, int aInputSize)
{
	int buffer, i, len = 0;

	if (aInputSize == -1)
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

		aBuf[len + 0] = buffer >> 16;
		aBuf[len + 1] = buffer >> 8;
		aBuf[len + 2] = buffer;
		len += 3;
	}

	if (i > 1)
	{
		buffer  = BASE64_CHAR_TO_BINARY(aInput[0]) << 18
				| BASE64_CHAR_TO_BINARY(aInput[1]) << 12;
		if (i > 2)
			buffer |= BASE64_CHAR_TO_BINARY(aInput[2]) << 6;
		// aInput not incremented as it is not used below.

		aBuf[len++] = buffer >> 16;
		if (i > 2)
			aBuf[len++] = buffer >> 8;
	}
	aBuf[len] = '\0';
	return len;
}

//
// class Debugger::Buffer - simplifies memory management.
//

// Write data into the buffer, expanding it as necessary.
int Debugger::Buffer::Write(char *aData, DWORD aDataSize)
{
	int err;

	if (aDataSize == MAXDWORD)
		aDataSize = strlen(aData);

	if (aDataSize == 0)
		return DEBUGGER_E_OK;

	if (mDataUsed + aDataSize > mDataSize && (err = Expand(mDataUsed + aDataSize)))
		return err;

	memcpy(mData + mDataUsed, aData, aDataSize);
	mDataUsed += aDataSize;
	return DEBUGGER_E_OK;
}

// Write formatted data into the buffer. Supports %s (char*), %e (char*, "&'<> escaped), %i (int), %u (unsigned int).
int Debugger::Buffer::WriteF(char *aFormat, ...)
{
	int i, err;
	DWORD len;
	char c, *format_ptr, *s, *param_ptr, *entity;
	char number_buf[MAX_NUMBER_SIZE];
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
		{	// Expand the buffer as necessary.
			if (mDataUsed + len > mDataSize && (err = Expand(mDataUsed + len)))
				return err;
		}
	} // for (len = 0, i = 0; i < 2; ++i)

	return DEBUGGER_E_OK;
}

// Convert a file path to a URI and write it to the buffer.
int Debugger::Buffer::WriteFileURI(char *aPath)
{
	int err, c, len = 9; // 8 for "file:///", 1 for '\0' (written by sprintf()).

	// Calculate required buffer size for path after encoding.
	for (char *ptr = aPath; c = *ptr; ++ptr)
	{
		if (isalnum(c) || strchr("-_.!~*'()/\\", c))
			++len;
		else
			len += 3;
	}

	// Ensure the buffer contains enough space.
	if (mDataUsed + len > mDataSize && (err = Expand(mDataUsed + len)))
		return err;

	Write("file:///", 8);

	// Write to the buffer, encoding as we go.
	for (char *ptr = aPath; c = *ptr; ++ptr)
	{
		if (isalnum(c) || strchr("-_.!~*()/", c))
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
	return Expand(mDataSize ? mDataSize * 2 : DEBUGGER_INITIAL_BUFFER_SIZE);
}

// Expand, if necessary, to meet a minimum required size.
int Debugger::Buffer::Expand(DWORD aRequiredSize)
{
	DWORD new_size;
	for (new_size = mDataSize ? mDataSize : DEBUGGER_INITIAL_BUFFER_SIZE
		; new_size < aRequiredSize
		; new_size *= 2);

	if (new_size > mDataSize)
	{
		char *new_data = (char*)realloc(mData, new_size);

		if (new_data == NULL)
			return FatalError(DEBUGGER_E_INTERNAL_ERROR);

		mData = new_data;
		mDataSize = new_size;
	}
	return DEBUGGER_E_OK;
}

// Remove data from the front of the buffer (i.e. after it is processed).
void Debugger::Buffer::Remove(DWORD aDataSize)
{
	// Move remaining data to the front of the buffer.
	if (aDataSize < mDataUsed)
		memmove(mData, mData + aDataSize, mDataUsed - aDataSize);
	mDataUsed -= aDataSize;
}

#endif