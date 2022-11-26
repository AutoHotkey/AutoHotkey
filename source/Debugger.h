﻿/*
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

#define DEBUGGER_STACK_PUSH(...)
#define DEBUGGER_STACK_POP()

#else

#ifndef Debugger_h
#define Debugger_h

#include <winsock2.h>
#include "script_object.h"


#define DEBUGGER_INITIAL_BUFFER_SIZE 2048

#define DEBUGGER_XML_TAG "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
#define DEBUGGER_XML_TAG_SIZE (_countof(DEBUGGER_XML_TAG)-1)

#define DEBUGGER_LANG_NAME AHK_NAME

// DBGp Error Codes
#define DEBUGGER_E_OK					0
#define DEBUGGER_E_PARSE_ERROR			1
#define DEBUGGER_E_INVALID_OPTIONS		3
#define DEBUGGER_E_UNIMPL_COMMAND		4
#define DEBUGGER_E_COMMAND_UNAVAIL		5

#define DEBUGGER_E_CAN_NOT_OPEN_FILE	100

#define DEBUGGER_E_BREAKPOINT_TYPE		201 // Breakpoint type not supported.
#define DEBUGGER_E_BREAKPOINT_INVALID	202 // Invalid line number or filename.
#define DEBUGGER_E_BREAKPOINT_NO_CODE	203 // No code on breakpoint line.
#define DEBUGGER_E_BREAKPOINT_STATE		204 // Invalid breakpoint state.
#define DEBUGGER_E_BREAKPOINT_NOT_FOUND	205 // No such breakpoint.
#define DEBUGGER_E_EVAL_FAIL			206

#define DEBUGGER_E_UNKNOWN_PROPERTY		300
#define DEBUGGER_E_INVALID_STACK_DEPTH	301
#define DEBUGGER_E_INVALID_CONTEXT		302

#define DEBUGGER_E_INTERNAL_ERROR		998 // Unrecoverable internal error, usually the result of a Winsock error.

#define DEBUGGER_E_CONTINUE				-1 // Internal code used by continuation commands.

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

extern LPCTSTR g_AutoExecuteThreadDesc;


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

	Breakpoint() : id(AllocateID()), type(BT_Line), state(BS_Enabled), temporary(false)
	{
	}

	static int AllocateID() { return ++sMaxId; }

private:
	static int sMaxId; // Highest used breakpoint ID.
};


// Forward-declarations (this file is included in script.h before these are declared).
class Line;
class NativeFunc;
struct UDFCallInfo;

template<typename T, int S> struct ScriptItemList;
typedef ScriptItemList<Var, VARLIST_INITIAL_SIZE> VarList;

struct DbgStack
{
	enum StackEntryType {SE_Thread, SE_BIF, SE_UDF};
	struct Entry
	{
		Line *line;
		union
		{
			LPCTSTR desc; // SE_Thread -- "auto-exec", hotkey/hotstring name, "timer", etc.
			NativeFunc *func; // SE_BIF
			UDFCallInfo *udf; // SE_UDF
		};
		StackEntryType type;
		LPCTSTR Name();
	};

	Entry *mBottom, *mTop, *mTopBound;
	size_t mSize; // i.e. capacity.

	DbgStack()
	{
		// We don't want to set the following too low since the stack would need to be reallocated,
		// but also don't want to set it too high since the average script mightn't recurse deeply;
		// if the stack size never approaches its maximum, there'll be wasted memory:
		mSize = 128;
		mBottom = (Entry *)malloc(mSize * sizeof(Entry));
		mTop = mBottom - 1; // ++mTop will be the first entry.
		mTopBound = mTop + mSize; // Topmost valid position.
	}

	int Depth()
	{
		return (int)(mTop + 1 - mBottom);
	}

	// noinline currently seems to have a slight effect on benchmarks.
	// Since it should be called very rarely, we don't want it inlined.
	void __declspec(noinline) Expand()
	{
		mSize *= 2;
		// To keep the design as simple as possible, assume the allocation will never fail.
		// These reallocations should be very rare: if size starts at 128 and doubles each
		// time, three expansions would bring it to 1024 entries, which is probably larger
		// than any script could need.  (Generally the program's stack will run out before
		// script recursion gets anywhere near that deep.)
		Entry *new_bottom = (Entry *)realloc(mBottom, mSize * sizeof(Entry));
		// Recalculate top of stack.
		mTop = mTop - mBottom + new_bottom;
		mBottom = new_bottom; // After using its old value.
		// Pre-calculate upper bound to keep Push() simple.
		mTopBound = mBottom - 1 + mSize;
	}

	Entry *Push();

	void Pop();

	void Push(LPCTSTR aDesc);
	void Push(NativeFunc *aFunc);
	void Push(UDFCallInfo *aRecurse);

	void GetLocalVars(int aDepth, VarList *&aVars, VarList *&aStaticVars, VarBkp *&aBkp, VarBkp *&aBkpEnd);
};

#define DEBUGGER_STACK_PUSH(aWhat)	g_Debugger.mStack.Push(aWhat);
#define DEBUGGER_STACK_POP()		g_Debugger.mStack.Pop();


enum PropertyContextType {PC_Local=0, PC_Global};


class Debugger
{
public:
	int Connect(const char *aAddress, const char *aPort);
	int Disconnect();
	void Exit(ExitReasons aExitReason, char *aCommandName=NULL); // Called when exiting AutoHotkey.
	inline bool IsConnected() { return mSocket != INVALID_SOCKET; }
	inline bool IsStepping() { return mInternalState >= DIS_StepInto; }
	inline bool HasStdErrHook() { return mStdErrMode != SR_Disabled; }
	inline bool HasStdOutHook() { return mStdOutMode != SR_Disabled; }
	inline bool BreakOnExceptionIsEnabled() { return mBreakOnException; }

	LPCTSTR WhatThrew();

	__declspec(noinline) // Avoiding inlining should reduce the code size of ExpandExpression(), which might help performance since this is only called when the debugger is connected.
	void PostExecFunctionCall(Line *aExpressionLine)
	{
		// If the debugger is stepping into/over/out from a function call, we want to
		// break at the line which called that function, since the next line to execute
		// might be a line in some other function (i.e. because the line which called
		// the function is "return func()" or calls another function after this one).
		if ((mInternalState == DIS_StepInto
			|| ((mInternalState == DIS_StepOut || mInternalState == DIS_StepOver)
				// Always '<' since '<=' (for StepOver) shouldn't be possible,
				// since we just returned from a function call:
				&& mStack.Depth() < mContinuationDepth))
			// The final check ensures we don't repeatedly break at a line containing
			// multiple built-in function calls; i.e. don't break unless some script
			// has been executed since we began evaluating aExpressionLine.  Something
			// like "return recursivefunc()" should work if this is StepInto or StepOver
			// since mCurrLine would probably be the '}' of that function:
			&& mCurrLine != aExpressionLine)
			PreExecLine(aExpressionLine);
	}

	// Code flow notification functions:
	int PreExecLine(Line *aLine); // Called before executing each line.
	bool PreThrow(ExprTokenType *aException);
	
	// Receive and process commands. Returns when a continuation command is received.
	int ProcessCommands(LPCSTR aBreakReason = nullptr);
	int Break(LPCSTR aReason = "ok");
	
	bool HasPendingCommand();

	// Streams
	int WriteStreamPacket(LPCTSTR aText, LPCSTR aType);
	bool OutputStdErr(LPCTSTR aText);
	bool OutputStdOut(LPCTSTR aText);

	#define DEBUGGER_COMMAND(cmd)	int cmd(char **aArgV, int aArgCount, char *aTransactionId)
	
	//
	// Debugger commands.
	//
	DEBUGGER_COMMAND(status);

	DEBUGGER_COMMAND(feature_get);
	DEBUGGER_COMMAND(feature_set);
	
	DEBUGGER_COMMAND(run);
	DEBUGGER_COMMAND(step_into);
	DEBUGGER_COMMAND(step_over);
	DEBUGGER_COMMAND(step_out);
	DEBUGGER_COMMAND(_break);
	DEBUGGER_COMMAND(stop);
	DEBUGGER_COMMAND(detach);
	
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


	Debugger() : mSocket(INVALID_SOCKET), mInternalState(DIS_Starting)
		, mMaxPropertyData(1024), mContinuationTransactionId(""), mStdErrMode(SR_Disabled), mStdOutMode(SR_Disabled)
		, mMaxChildren(20), mMaxDepth(2), mDisabledHooks(0)
		, mThrownToken(NULL), mBreakOnExceptionID(0), mBreakOnExceptionWasSet(false), mBreakOnExceptionIsTemporary(false), mBreakOnException(false)
	{
	}

	
	// Stack - keeps track of threads and function calls.
	DbgStack mStack;
	friend struct DbgStack;

private:
	SOCKET mSocket;
	Line *mCurrLine; // Similar to g_script.mCurrLine, but may be different when breaking post-function-call, before continuing expression evaluation.
	ExprTokenType *mThrownToken; // The exception that triggered the current exception breakpoint.
	bool mBreakOnExceptionWasSet, mBreakOnExceptionIsTemporary, mBreakOnException; // Supports a single coverall breakpoint exception.
	int mBreakOnExceptionID;

	class Buffer
	{
	public:
		int Write(char *aData, size_t aDataSize=-1);
		int WriteF(const char *aFormat, ...);
		int WriteEncodeBase64(const char *aData, size_t aDataSize, bool aSkipBufferSizeCheck = false);
		int Expand();
		int ExpandIfNecessary(size_t aRequiredSize);
		void Delete(size_t aDataSize);
		void Clear();

		Buffer() : mData(NULL), mDataSize(0), mDataUsed(0), mFailed(FALSE) {}
	
		char *mData;
		size_t mDataSize;
		size_t mDataUsed;
		BOOL mFailed;

		~Buffer() {
			if (mData)
				free(mData);
		}
	private:
		int EstimateFileURILength(LPCTSTR aPath);
		void WriteFileURI(LPCTSTR aPath);
	} mCommandBuf, mResponseBuf;

	enum DebuggerInternalStateType {
		DIS_None = 0,
		DIS_Starting = DIS_None,
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

	int mContinuationDepth; // Stack depth at last continuation command, for step_into/step_over.
	CStringA mContinuationTransactionId; // transaction_id of last continuation command.

	int mMaxPropertyData, mMaxChildren, mMaxDepth;

	HookType mDisabledHooks;


	enum PropertyType
	{
		PropNone = 0,
		PropVar,
		PropVarBkp,
		PropValue, // Read-only (ExprTokenType)
		PropEnum
	};

	struct PropertySource
	{
		PropertyType kind;
		Var *var;
		VarBkp *bkp;
		Object::Variant *field;
		ResultToken value;
		IObject *this_object = nullptr;
		PropertySource(LPTSTR aResultBuf)
		{
			value.InitResult(aResultBuf);
		}
		~PropertySource()
		{
			if (this_object)
				this_object->Release();
			value.Free();
		}
	};

	struct PropertyInfo : PropertySource
	{
		LPCSTR name;
		CStringA &fullname;
		LPSTR facet; // Initialised during writing.
		bool is_alias = false, is_builtin = false, is_static = false; // Facets.
		int page = 0, pagesize; // Attributes which are also parameters.
		int max_data; // Parameters.
		int max_depth;

		PropertyInfo(CStringA &aNameBuf, LPTSTR aResultBuf)
			: PropertySource(aResultBuf), fullname(aNameBuf) {}
	};

	
	struct PropertyWriter : public IDebugProperties
	{
		Debugger &mDbg;
		PropertyInfo &mProp;
		IObject *mObject;
		size_t mNameLength;
		int mDepth;
		int mError;

		PropertyWriter(Debugger &aDbg, PropertyInfo &aProp, IObject *aObject)
			: mDbg(aDbg)
			, mProp(aProp)
			, mObject(aObject)
			, mNameLength(aProp.fullname.GetLength())
			, mDepth(0)
			, mError(0)
		{
		}

		void WriteProperty(LPCSTR aName, ExprTokenType &aValue);
		void WriteProperty(LPCWSTR aName, ExprTokenType &aValue);
		void WriteProperty(ExprTokenType &aKey, ExprTokenType &aValue);
		void WriteBaseProperty(IObject *aBase);
		void WriteDynamicProperty(LPTSTR aName);
		void WriteEnumItems(IObject *aEnumerable, int aStart, int aEnd);

		void _WriteProperty(ExprTokenType &aValue, IObject *aThisOverride = nullptr);

		void BeginProperty(LPCSTR aName, LPCSTR aType, int aNumChildren, DebugCookie &aCookie);
		void EndProperty(DebugCookie aCookie);
	};


	// Receive next command from debugger UI:
	int ReceiveCommand(int *aCommandSize=NULL);

	// Send XML response to debugger UI:
	int SendResponse();
	int SendErrorResponse(char *aCommandName, char *aTransactionId, int aError=999, char *aExtraAttributes=NULL);
	int SendStandardResponse(char *aCommandName, char *aTransactionId);
	int SendContinuationResponse(LPCSTR aCommand = nullptr, LPCSTR aStatus = "break", LPCSTR aReason = "ok");

	int EnterBreakState(LPCSTR aReason = "ok");
	void ExitBreakState();

	int WriteBreakpointXml(Breakpoint *aBreakpoint, Line *aLine);
	int WriteExceptionBreakpointXml();
	Line *FindFirstLineForBreakpoint(int file_index, UINT line_no);

	void AppendPropertyName(CStringA &aNameBuf, size_t aParentNameLength, const char *aName);
	void AppendStringKey(CStringA &aNameBuf, size_t aParentNameLength, const char *aKey);

	int GetPropertyInfo(Var &aVar, PropertyInfo &aProp);
	int GetPropertyInfo(VarBkp &aBkp, PropertyInfo &aProp);
	
	int GetPropertyValue(Var &aVar, PropertySource &aProp);

	int WritePropertyXml(PropertyInfo &aProp);
	int WritePropertyXml(PropertyInfo &aProp, LPTSTR aName);
	int WritePropertyXml(PropertyInfo &aProp, IObject *aObject);

	int WritePropertyData(LPCTSTR aData, size_t aDataSize, int aMaxEncodedSize);
	int WritePropertyData(ExprTokenType &aValue, int aMaxEncodedSize);

	int WriteEnumItems(PropertyInfo &aProp, IObject *aObject);

	int ParsePropertyName(LPCSTR aFullName, int aDepth, int aVarScope, ExprTokenType *aSetValue
		, PropertySource &aResult);
	int property_get_or_value(char **aArgV, int aArgCount, char *aTransactionId, bool aIsPropertyGet);
	int redirect_std(char **aArgV, int aArgCount, char *aTransactionId, char *aCommandName);
	int run_step(char **aArgV, int aArgCount, char *aTransactionId, char *aCommandName, DebuggerInternalStateType aNewState);

	// Decode a file URI in-place.
	void DecodeURI(char *aUri);
	
	static const char *sBase64Chars;
	static size_t Base64Encode(char *aBuf, const char *aInput, size_t aInputSize = -1);
	static size_t Base64Decode(char *aBuf, const char *aInput, size_t aInputSize = -1);


	//typedef int (Debugger::*CommandFunc)(char **aArgV, int aArgCount, char *aTransactionId);
	typedef DEBUGGER_COMMAND((Debugger::*CommandFunc));
	
	struct CommandDef
	{
		const char *mName;
		CommandFunc mFunc;
	};

	static CommandDef sCommands[];
	

	// Debugger::ParseArgs
	//
	// Returns DEBUGGER_E_OK on success, or a DBGp error code otherwise.
	//
	// The Xdebug/DBGp documentation is very vague about command line rules,
	// so this function has no special treatment of quotes, backslash, etc.
	// There is currently no way to include a literal " -" in an arg as it
	// would be recognized as the beginning of the next arg.
	//
	int ParseArgs(char *aArgs, char **aArgV, int &aArgCount, char *&aTransactionId);
	
	// Caller must verify that aArg is within bounds:
	inline char *ArgValue(char **aArgV, int aArg) { return aArgV[aArg] + 1; }
	inline char  ArgChar(char **aArgV, int aArg) { return *aArgV[aArg]; }

	// Fatal debugger error. Prompt user to terminate script or only disconnect debugger.
	static int FatalError(LPCTSTR aMessage = DEBUGGER_ERR_INTERNAL DEBUGGER_ERR_DISCONNECT_PROMPT);
};


constexpr auto SCRIPT_STACK_BUF_SIZE = 2048; // Character limit for Error().Stack and stack traces generated by the error dialog.

void GetScriptStack(LPTSTR aBuf, int aBufSize, DbgStack::Entry *aTop = nullptr);


#endif
#endif