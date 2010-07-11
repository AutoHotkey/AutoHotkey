/*
Debugger.h

Original code by Steve Gray.

This software is provided 'as-is', without any express or implied
warranty. In no event will the authors be held liable for any damages
arising from the use of this software.

Permission is granted to anyone to use this software for any purpose,
including commercial applications, and to alter it and redistribute it
freely, without restriction.
*/

#pragma once

#ifndef CONFIG_DEBUGGER

#define DEBUGGER_STACK_PUSH(a,b,c,d)
#define DEBUGGER_STACK_POP()

#else

#ifndef Debugger_h
#define Debugger_h

#include <winsock2.h>


#define DEBUGGER_INITIAL_BUFFER_SIZE 2048

#define DEBUGGER_XML_TAG "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
#define DEBUGGER_XML_TAG_SIZE (_countof(DEBUGGER_XML_TAG)-1)

// DBGp Error Codes
#define DEBUGGER_E_OK					0
#define DEBUGGER_E_PARSE_ERROR			1
#define DEBUGGER_E_INVALID_OPTIONS		3
#define DEBUGGER_E_UNIMPL_COMMAND		4

#define DEBUGGER_E_CAN_NOT_OPEN_FILE	100

#define DEBUGGER_E_BREAKPOINT_TYPE		201 // Breakpoint type not supported.
#define DEBUGGER_E_BREAKPOINT_INVALID	202 // Invalid line number or filename.
#define DEBUGGER_E_BREAKPOINT_NO_CODE	203 // No code on breakpoint line.
#define DEBUGGER_E_BREAKPOINT_STATE		204 // Invalid breakpoint state.
#define DEBUGGER_E_BREAKPOINT_NOT_FOUND	205 // No such breakpoint.

#define DEBUGGER_E_UNKNOWN_PROPERTY		300
#define DEBUGGER_E_INVALID_STACK_DEPTH	301
#define DEBUGGER_E_INVALID_CONTEXT		302

#define DEBUGGER_E_INTERNAL_ERROR		998 // Unrecoverable internal error, usually the result of a Winsock error.

// Error messages: these are shown directly to the user, so are in the native string format.
#define DEBUGGER_ERR_INTERNAL			_T("An internal error has occurred in the debugger engine.")
#define DEBUGGER_ERR_DISCONNECT_PROMPT	_T("\nContinue running the script without the debugger?")
#define DEBUGGER_ERR_FAILEDTOCONNECT	_T("Failed to connect to an active debugger client.")

// Buffer size required for a given XML message size, plus protocol overhead.
// Format: data_length NULL xml_tag data NULL
//#define DEBUGGER_XML_SIZE_REQUIRED(xml_size) (MAX_NUMBER_LENGTH + DEBUGGER_XML_TAG_SIZE + xml_size + 2)
#define DEBUGGER_RESPONSE_OVERHEAD (MAX_INTEGER_LENGTH + DEBUGGER_XML_TAG_SIZE + 2)

class Debugger;

extern Debugger g_Debugger;
// jackieku: modified to hold the buffer.
extern CStringA g_DebuggerHost;
extern CStringA g_DebuggerPort;


enum BreakpointTypeType {BT_Line, BT_Call, BT_Return, BT_Exception, BT_Conditional, BT_Watch};
enum BreakpointStateType {BS_Disabled=0, BS_Enabled};

class Breakpoint
{
public:
	int id;
	char type;
	char state;
	bool temporary;
	
	// Not yet supported: function, hit_count, hit_value, hit_condition, exception

	Breakpoint() : id(++sMaxId), type(BT_Line), state(BS_Enabled), temporary(false)
	{
	}

private:
	static int sMaxId; // Highest used breakpoint ID.
};


// Forward-declarations (this file is included in script.h before these are declared).
class Line;
class Func;
class Label;


enum StackEntryTypeType {SE_Thread, SE_Sub, SE_Func};

struct StackEntry
{
	StackEntryTypeType type;
	Line *line;
	union {
		TCHAR *desc; // SE_Thread: "auto-exec", hotkey/hotstring name, "timer", etc.
		Label *sub;
		Func *func;
	};
	StackEntry *upper, *lower;
};

#define DEBUGGER_STACK_PUSH(aType, aLine, aDataType, aData) \
	if (g_Debugger.IsConnected()) \
	{ \
		StackEntry *se = (StackEntry*) _alloca(sizeof(StackEntry)); \
		se->type = aType; \
		se->line = aLine; \
		se->aDataType = aData; \
		g_Debugger.StackPush(se); \
	}
// Func::Call() calls StackPop() directly rather than using this macro, as it requires extra work to allow the user
// to inspect variables before actually returning. If this macro is changed, also update that section.
#define DEBUGGER_STACK_POP() \
	if (g_Debugger.IsConnected()) \
		g_Debugger.StackPop();


enum PropertyContextType {PC_Local=0, PC_Global};


class Debugger
{
public:
	int Connect(const char *aAddress, const char *aPort);
	int Disconnect();
	void Exit(ExitReasons aExitReason); // Called when exiting AutoHotkey.
	inline bool IsConnected() { return mSocket != INVALID_SOCKET; }
	inline bool IsStepping() { return mInternalState >= DIS_StepInto; }
	inline bool HasStdErrHook() { return mStdErrMode != SR_Disabled; }
	inline bool HasStdOutHook() { return mStdOutMode != SR_Disabled; }
	inline bool ShouldBreakAfterFunctionCall()
	{
		return mInternalState == DIS_StepInto
			|| (mInternalState == DIS_StepOut || mInternalState == DIS_StepOver) && mStackDepth < mContinuationDepth;
	}

	// Code flow notification functions:
	int PreExecLine(Line *aLine); // Called before executing each line.
	
	// Call-stack: track threads, function-calls and gosub.
	int StackPush(StackEntry *aEntry);
	int StackPop();


	// Receive and process commands. Returns when a continuation command is received.
	int ProcessCommands();

	// Streams
	int WriteStreamPacket(LPCTSTR aText, LPCSTR aType);
	void OutputDebug(LPCTSTR aText);
	bool FileAppendStdOut(LPCTSTR aText);

	#define DEBUGGER_COMMAND(cmd)	int cmd(char *aArgs)
	
	//
	// Debugger commands.
	//
	DEBUGGER_COMMAND(status);

	DEBUGGER_COMMAND(feature_get);
	DEBUGGER_COMMAND(feature_set);
	
	/*DEBUGGER_COMMAND(run);
	DEBUGGER_COMMAND(step_into);
	DEBUGGER_COMMAND(step_over);
	DEBUGGER_COMMAND(step_out);*/
	DEBUGGER_COMMAND(stop);
	
	DEBUGGER_COMMAND(breakpoint_set);
	DEBUGGER_COMMAND(breakpoint_get);
	DEBUGGER_COMMAND(breakpoint_update);
	DEBUGGER_COMMAND(breakpoint_remove);
	DEBUGGER_COMMAND(breakpoint_list);
	
	DEBUGGER_COMMAND(stack_depth);
	DEBUGGER_COMMAND(stack_get);
	DEBUGGER_COMMAND(context_names);
	DEBUGGER_COMMAND(context_get);

	DEBUGGER_COMMAND(typemap_get);
	DEBUGGER_COMMAND(property_get);
	DEBUGGER_COMMAND(property_set);
	DEBUGGER_COMMAND(property_value);
	
	DEBUGGER_COMMAND(source);

	DEBUGGER_COMMAND(redirect_stdout);
	DEBUGGER_COMMAND(redirect_stderr);


	Debugger() : mSocket(INVALID_SOCKET), mInternalState(DIS_Starting), mStackDepth(0)
		, mMaxPropertyData(1024), mContinuationTransactionId(""), mStdErrMode(SR_Disabled), mStdOutMode(SR_Disabled)
	{
		// Create root entry for simplicity.
		mStackTop = mStack = new StackEntry();
	}


private:
	SOCKET mSocket;

	class Buffer
	{
	public:
		int Write(char *aData, size_t aDataSize=-1);
		int WriteF(const char *aFormat, ...);
		int WriteFileURI(const char *aPath);
		int WriteEncodeBase64(const char *aData, size_t aDataSize, bool aSkipBufferSizeCheck = false);
		int Expand();
		int ExpandIfNecessary(size_t aRequiredSize);
		void Remove(size_t aDataSize);

		Buffer() : mData(NULL), mDataSize(0), mDataUsed(0) {}
	
		char *mData;
		size_t mDataSize;
		size_t mDataUsed;

		~Buffer() {
			if (mData)
				free(mData);
		}
	} mCommandBuf, mResponseBuf;

	enum DebuggerInternalStateType {
		DIS_Starting,
		DIS_Run,
		DIS_Break,
		DIS_StepInto,
		DIS_StepOver,
		DIS_StepOut
	} mInternalState;

	enum StreamRedirectType {
		SR_Disabled = 0,
		SR_Copy = 1,
		SR_Redirect = 2
	} mStdErrMode, mStdOutMode;

	// Stack - keeps track of threads, function calls and gosubs.
	StackEntry *mStack, *mStackTop;
	int mStackDepth;

	int mContinuationDepth; // Stack depth at last continuation command, for step_into/step_over.
	char *mContinuationTransactionId; // transaction_id of last continuation command.

	VarSizeType mMaxPropertyData;


	// Receive next command from debugger UI:
	int ReceiveCommand(int *aCommandSize=NULL);

	// Send XML response to debugger UI:
	int SendResponse();
	int SendErrorResponse(char *aCommandName, char *aTransactionId, int aError=999, char *aExtraAttributes=NULL);
	int SendStandardResponse(char *aCommandName, char *aTransactionId);
	int SendContinuationResponse(char *aStatus="break", char *aReason="ok");

	int WriteBreakpointXml(Breakpoint *aBreakpoint, Line *aLine);
	int WritePropertyXml(Var *aVar, VarSizeType aMaxData=VARSIZE_MAX);
	int WriteVarSizeAndData(Var *aVar, VarSizeType aMaxData=VARSIZE_MAX);

	int property_get_or_value(char *aArgs, bool aIsPropertyGet);
	int redirect_std(char *aArgs, char *aCommandName);

	// Decode a file URI in-place.
	void DecodeURI(char *aUri);
	
	static const char *sBase64Chars;
	static size_t Base64Encode(char *aBuf, const char *aInput, size_t aInputSize = -1);
	static size_t Base64Decode(char *aBuf, const char *aInput, size_t aInputSize = -1);


	// Debugger::GetNextArg
	//
	// Returns DEBUGGER_E_OK on success, or a DBGp error code otherwise.
	// aArgs is set to the beginning of the next arg, or NULL if no more args.
	// aArg is set to the arg character, or '\0' if no args.
	// aValue is set to the value of the arg, or NULL if no value.
	//
	// The Xdebug/DBGp documentation is very vague about command line rules,
	// so this function has no special treatment of quotes, backslash, etc.
	// There is currently no way to include a literal " -" in an arg as it
	// would be recognized as the beginning of the next arg.
	//
	int GetNextArg(char *&aArgs, char &aArg, char *&aValue);

	// Fatal debugger error. Prompt user to terminate script or only disconnect debugger.
	static int FatalError(int aErrorCode, LPCTSTR aMessage = DEBUGGER_ERR_INTERNAL DEBUGGER_ERR_DISCONNECT_PROMPT);
};

#endif
#endif