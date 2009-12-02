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

#ifndef SCRIPT_DEBUG

#define DEBUGGER_STACK_PUSH(a,b,c,d)
#define DEBUGGER_STACK_POP()

#else

#ifndef Debugger_h
#define Debugger_h

#include <WinSock2.h>


#define DEBUGGER_INITIAL_BUFFER_SIZE 2048

#define DEBUGGER_XML_TAG "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
#define DEBUGGER_XML_TAG_SIZE 38

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

// Error Messages
#define DEBUGGER_ERR_FAILEDTOCONNECT "Failed to connect to debugger UI."
#define DEBUGGER_ERR_INTERNAL "An internal error has occurred in the debugger engine."
#define DEBUGGER_ERR_DISCONNECT_PROMPT " Continue running the script without the debugger?"

// Buffer size required for a given XML message size, plus protocol overhead.
// Format: data_length NULL xml_tag data NULL
//#define DEBUGGER_XML_SIZE_REQUIRED(xml_size) (MAX_NUMBER_LENGTH + DEBUGGER_XML_TAG_SIZE + xml_size + 2)
#define DEBUGGER_RESPONSE_OVERHEAD (MAX_INTEGER_LENGTH + DEBUGGER_XML_TAG_SIZE + 2)

class Debugger;

extern Debugger g_Debugger;
extern char *g_DebuggerHost;
extern char *g_DebuggerPort;


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
		char *desc; // SE_Thread: "auto-exec", hotkey/hotstring name, "timer", etc.
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
	int Connect(char *aAddress, char *aPort);
	int Disconnect();
	void Exit(ExitReasons aExitReason); // Called when exiting AutoHotkey.
	inline bool IsConnected() { return mSocket != INVALID_SOCKET; }
	inline bool IsStepping() { return mInternalState >= DIS_StepInto; }
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
		, mMaxPropertyData(1024), mContinuationTransactionId("")
	{
		// Create root entry for simplicity.
		mStackTop = mStack = new StackEntry();
	}


private:
	SOCKET mSocket;

	class Buffer
	{
	public:
		int Write(char *aData, DWORD aDataSize=MAXDWORD);
		int WriteF(char *aFormat, ...);
		int WriteFileURI(char *aPath);
		int Expand();
		int Expand(DWORD aRequiredSize);
		void Remove(DWORD aDataSize);

		Buffer() : mData(NULL), mDataSize(0), mDataUsed(0) {}
	
		char *mData;
		DWORD mDataSize;
		DWORD mDataUsed;

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
	
	static char *sBase64Chars;
	int Base64Encode(char *aBuf, const char *aInput, int aInputSize=-1);
	int Base64Decode(char *aBuf, const char *aInput, int aInputSize=-1);


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
	static int FatalError(int aErrorCode, char *aMessage=NULL);
};

#endif
#endif