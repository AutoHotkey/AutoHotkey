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

#ifndef defines_h
#define defines_h


// Disable silly performance warning about converting int to bool:
// Unlike other typecasts from a larger type to a smaller, I'm 99% sure
// that all compilers are supposed to do something special for bool,
// not just truncate.  Example:
// bool x = 0xF0000000
// The above should give the value "true" to x, not false which is
// what would happen if:
// char x = 0xF0000000
//
#ifdef _MSC_VER
#pragma warning(disable:4800)
#endif

#include "ahkversion.h"
#define AHK_WEBSITE "https://autohotkey.com"

#ifdef AUTOHOTKEYSC
#define SCRIPT_RESOURCE_NAME _T(">AUTOHOTKEY SCRIPT<")
#else
#define SCRIPT_RESOURCE_SPEC _T("*#1")
#define SCRIPT_RESOURCE_NAME MAKEINTRESOURCE(1)
#define SCRIPT_PRESOURCE_SPEC _T("*#2")
#define SCRIPT_PRESOURCE_NAME MAKEINTRESOURCE(2)
#endif

// Window class names: Changing these may result in new versions not being able to detect any old instances
// that may be running (such as the use of FindWindow() in WinMain()).  It may also have other unwanted
// effects, such as anything in the OS that relies on the class name that the user may have changed the
// settings for, such as whether to hide the tray icon (though it probably doesn't use the class name
// in that example).
// MSDN: "Because window classes are process specific, window class names need to be unique only within
// the same process. Also, because class names occupy space in the system's private atom table, you
// should keep class name strings as short a possible:
#define WINDOW_CLASS_MAIN _T("AutoHotkey")
#define WINDOW_CLASS_GUI _T("AutoHotkeyGUI") // There's a section in Script::Edit() that relies on these all starting with "AutoHotkey".

#define EXT_AUTOHOTKEY _T(".ahk")
#define AHK_HELP_FILE _T("AutoHotkey.chm")

// AutoIt2 supports lines up to 16384 characters long, and we want to be able to do so too
// so that really long lines from aut2 scripts, such as a chain of IF commands, can be
// brought in and parsed.  In addition, it also allows continuation sections to be long.
#define LINE_SIZE (16384 + 1)  // +1 for terminator.  Don't increase LINE_SIZE above 65535 without considering ArgStruct::length's type (WORD).

// Maximum length of a Unicode file path, plus null-terminator.
#define MAX_WIDE_PATH 32768
#ifdef UNICODE
#define T_MAX_PATH MAX_WIDE_PATH
#else
#define T_MAX_PATH MAX_PATH
#endif

// The following avoid having to link to OLDNAMES.lib, but they probably don't
// reduce code size at all.
#define stricmp(str1, str2) _stricmp(str1, str2)
#define strnicmp(str1, str2, size) _strnicmp(str1, str2, size)

// Items that may be needed for VC++ 6.X:
#ifndef VK_XBUTTON1
	#define VK_XBUTTON1       0x05    /* NOT contiguous with L & RBUTTON */
	#define VK_XBUTTON2       0x06    /* NOT contiguous with L & RBUTTON */
	#define WM_NCXBUTTONDOWN                0x00AB
	#define WM_NCXBUTTONUP                  0x00AC
	#define WM_NCXBUTTONDBLCLK              0x00AD
	#define GET_WHEEL_DELTA_WPARAM(wParam)  ((short)HIWORD(wParam))
	#define WM_XBUTTONDOWN                  0x020B
	#define WM_XBUTTONUP                    0x020C
	#define WM_XBUTTONDBLCLK                0x020D
	#define GET_KEYSTATE_WPARAM(wParam)     (LOWORD(wParam))
	#define GET_NCHITTEST_WPARAM(wParam)    ((short)LOWORD(wParam))
	#define GET_XBUTTON_WPARAM(wParam)      (HIWORD(wParam))
	/* XButton values are WORD flags */
	#define XBUTTON1      0x0001
	#define XBUTTON2      0x0002
#endif
#ifndef HIMETRIC_INCH
	#define HIMETRIC_INCH 2540
#endif

#define IS_32BIT(signed_value_64) (signed_value_64 >= INT_MIN && signed_value_64 <= INT_MAX)
#define GET_BIT(buf,n) (((buf) & (1 << (n))) >> (n))
#define SET_BIT(buf,n,val) ((val) ? ((buf) |= (1<<(n))) : (buf &= ~(1<<(n))))

// FAIL = 0 to remind that FAIL should have the value zero instead of something arbitrary
// because some callers may simply evaluate the return result as true or false
// (and false is a failure):
enum ResultType {FAIL = 0, OK, WARN = OK, CRITICAL_ERROR  // Some things might rely on OK==1 (i.e. boolean "true")
	, CONDITION_TRUE, CONDITION_FALSE
	, LOOP_BREAK, LOOP_CONTINUE
	, EARLY_RETURN, EARLY_EXIT // EARLY_EXIT needs to be distinct from FAIL for ExitApp() and AutoExecSection().
	, FAIL_OR_OK // For LineError/RuntimeError, error is continuable.
};

typedef HRESULT FResult;

enum ExcptModeType {EXCPTMODE_NONE = 0
	//, EXCPTMODE_TRY = 1 // Currently unused: Try block present.
	, EXCPTMODE_CATCH = 2 // Exception will be suppressed or caught.
	, EXCPTMODE_DELETE = 4 // Unhandled exceptions will display ERR_ABORT_DELETE vs. ERR_ABORT.
	, EXCPTMODE_CAUGHT = 0x10 // An exception is already being handled within a CATCH, and is not shadowed by TRY.
};

#define SEND_MODES { _T("Event"), _T("Input"), _T("Play"), _T("InputThenPlay") } // Must match the enum below.
enum SendModes {SM_EVENT, SM_INPUT, SM_PLAY, SM_INPUT_FALLBACK_TO_PLAY, SM_INVALID}; // SM_EVENT must be zero.
// In above, SM_INPUT falls back to SM_EVENT when the SendInput mode would be defeated by the presence
// of a keyboard/mouse hooks in another script (it does this because SendEvent is superior to a
// crippled/interruptible SendInput due to SendEvent being able to dynamically adjust to changing
// conditions [such as the user releasing a modifier key during the Send]).  By contrast,
// SM_INPUT_FALLBACK_TO_PLAY falls back to the SendPlay mode.  SendInput has this extra fallback behavior
// because it's likely to become the most popular sending method.

enum SendRawModes {SCM_NOT_RAW = FALSE, SCM_RAW, SCM_RAW_TEXT};
typedef UCHAR SendRawType;

enum ExitReasons {EXIT_CRITICAL = -2, EXIT_DESTROY = -1, EXIT_NONE = 0, EXIT_ERROR, EXIT_LOGOFF, EXIT_SHUTDOWN
	, EXIT_CLOSE, EXIT_MENU, EXIT_EXIT, EXIT_RELOAD, EXIT_SINGLEINSTANCE};
#define EXITREASON_MUST_EXIT(er) (static_cast<ExitReasons>(er) <= EXIT_DESTROY)

enum WarnType {WARN_LOCAL_SAME_AS_GLOBAL, WARN_UNREACHABLE, WARN_VAR_UNSET, WARN_ALL};
#define WARN_TYPE_STRINGS _T("LocalSameAsGlobal"), _T("Unreachable"), _T("VarUnset"), _T("All")

enum WarnMode {WARNMODE_OFF, WARNMODE_OUTPUTDEBUG, WARNMODE_MSGBOX, WARNMODE_STDOUT};	// WARNMODE_OFF must be zero.
#define WARN_MODE_STRINGS _T("Off"), _T("OutputDebug"), _T("MsgBox"), _T("StdOut")

enum SingleInstanceType {SINGLE_INSTANCE_OFF, SINGLE_INSTANCE_PROMPT, SINGLE_INSTANCE_REPLACE
	, SINGLE_INSTANCE_IGNORE }; // SINGLE_INSTANCE_OFF must be zero.

enum MenuTypeType {MENU_TYPE_NONE, MENU_TYPE_POPUP, MENU_TYPE_BAR}; // NONE must be zero.

// These are used for things that can be turned on, off, or left at a
// neutral default value that is neither on nor off.  INVALID must
// be zero:
enum ToggleValueType {TOGGLE_INVALID = 0, TOGGLED_ON, TOGGLED_OFF, ALWAYS_ON, ALWAYS_OFF, TOGGLE
	, NEUTRAL, TOGGLE_SEND, TOGGLE_MOUSE, TOGGLE_SENDANDMOUSE, TOGGLE_DEFAULT
	, TOGGLE_MOUSEMOVE, TOGGLE_MOUSEMOVEOFF};

// Some things (such as ListView sorting) rely on SCS_INSENSITIVE being zero.
enum StringCaseSenseType {SCS_INSENSITIVE, SCS_SENSITIVE, SCS_INSENSITIVE_LOCALE, SCS_INSENSITIVE_LOGICAL, SCS_INVALID};

enum SymbolType // For use with ExpandExpression() and IsNumeric().
{
	// The sPrecedence array in ExpandExpression() must be kept in sync with any additions, removals,
	// or re-ordering of the below.  When reordering or adding new symbols, take care not to break
	// the range checks in the various macros defined below.
	 PURE_NOT_NUMERIC // Must be zero/false because callers rely on that.
	, PURE_INTEGER, PURE_FLOAT
	, SYM_STRING = PURE_NOT_NUMERIC, SYM_INTEGER = PURE_INTEGER, SYM_FLOAT = PURE_FLOAT // Specific operand types.
#define IS_NUMERIC(symbol) ((symbol) == SYM_INTEGER || (symbol) == SYM_FLOAT) // Ordered for short-circuit performance.
	, SYM_MISSING // Only used in parameter lists.
	, SYM_VAR // An operand that is a variable's contents.
	, SYM_OBJECT // L31: Represents an IObject interface pointer.
	, SYM_DYNAMIC // A dynamic variable reference/double-deref.  Also used in Object::Variant to identify dynamic properties.
	, SYM_SUPER // Special operand just for the "super" keyword.  Should only ever appear as the target of SYM_FUNC with callsite->func == nullptr.
#define SUPER_KEYWORD _T("Super")
	, SYM_OPERAND_END // Marks the symbol after the last operand.  This value is used below.
	, SYM_BEGIN = SYM_OPERAND_END  // SYM_BEGIN is a special marker to simplify the code.
#define IS_OPERAND(symbol) ((symbol) < SYM_OPERAND_END)
	, SYM_POST_INCREMENT, SYM_POST_DECREMENT // Kept in this position for use by YIELDS_AN_OPERAND() [helps performance].
#define IS_POSTFIX_OPERATOR(symbol) ((symbol) == SYM_POST_INCREMENT || (symbol) == SYM_POST_DECREMENT)
	, SYM_DOT // DOT must precede SYM_OPAREN so YIELDS_AN_OPERAND(SYM_GET) == TRUE, allowing auto-concat to work for it even though it is positioned after its second operand.
	, SYM_CPAREN, SYM_CBRACKET, SYM_CBRACE, SYM_OPAREN, SYM_OBRACKET, SYM_OBRACE, SYM_COMMA  // CPAREN (close-paren)/CBRACKET/CBRACE must come right before OPAREN for YIELDS_AN_OPERAND.
#define IS_OPAREN_LIKE(symbol) ((symbol) <= SYM_OBRACE && (symbol) >= SYM_OPAREN)
#define IS_CPAREN_LIKE(symbol) ((symbol) <= SYM_CBRACE && (symbol) >= SYM_CPAREN)
#define IS_OPAREN_MATCHING_CPAREN(sym_oparen, sym_cparen) ((sym_oparen - sym_cparen) == (SYM_OPAREN - SYM_CPAREN)) // Requires that (IS_OPAREN_LIKE(sym_oparen) || IS_CPAREN_LIKE(sym_cparen)) is true.
#define SYM_CPAREN_FOR_OPAREN(symbol) ((symbol) - (SYM_OPAREN - SYM_CPAREN)) // Caller must confirm it is OPAREN or OBRACKET.
#define SYM_OPAREN_FOR_CPAREN(symbol) ((symbol) + (SYM_OPAREN - SYM_CPAREN)) // Caller must confirm it is CPAREN or CBRACKET.
#define YIELDS_AN_OPERAND(symbol) ((symbol) < SYM_OPAREN) // CPAREN also covers the tail end of a function call.  Post-inc/dec yields an operand for things like Var++ + 2.  Definitely needs the parentheses around symbol.
	, SYM_ASSIGN, SYM_ASSIGN_ADD, SYM_ASSIGN_SUBTRACT, SYM_ASSIGN_MULTIPLY, SYM_ASSIGN_DIVIDE, SYM_ASSIGN_INTEGERDIVIDE
	, SYM_ASSIGN_BITOR, SYM_ASSIGN_BITXOR, SYM_ASSIGN_BITAND, SYM_ASSIGN_BITSHIFTLEFT, SYM_ASSIGN_BITSHIFTRIGHT, SYM_ASSIGN_BITSHIFTRIGHT_LOGICAL // SYM_ASSIGN_BITSHIFTLEFT_LOGICAL doesn't exist but <<<= is the same as <<=
	, SYM_ASSIGN_CONCAT // THIS MUST BE KEPT AS THE LAST (AND SYM_ASSIGN THE FIRST) BECAUSE THEY'RE USED IN A RANGE-CHECK.
#define IS_ASSIGNMENT_EXCEPT_POST_AND_PRE(symbol) (symbol <= SYM_ASSIGN_CONCAT && symbol >= SYM_ASSIGN) // Check upper bound first for short-circuit performance.
#define IS_ASSIGNMENT_OR_POST_OP(symbol) (IS_ASSIGNMENT_EXCEPT_POST_AND_PRE(symbol) || symbol == SYM_POST_INCREMENT || symbol == SYM_POST_DECREMENT)
	, SYM_IFF_ELSE, SYM_IFF_THEN // THESE TERNARY OPERATORS MUST BE KEPT IN THIS ORDER AND ADJACENT TO THE BELOW.
	, SYM_OR_MAYBE, SYM_OR, SYM_AND // MUST BE KEPT IN THIS ORDER AND ADJACENT TO THE ABOVE for the range checks below.
#define IS_SHORT_CIRCUIT_OPERATOR(symbol) ((symbol) <= SYM_AND && (symbol) >= SYM_IFF_THEN) // Excludes SYM_IFF_ELSE, which acts as a simple jump after the THEN branch is evaluated.
#define SYM_USES_CIRCUIT_TOKEN(symbol) ((symbol) <= SYM_AND && (symbol) >= SYM_IFF_ELSE)
	, SYM_IS
	, SYM_EQUAL, SYM_EQUALCASE, SYM_NOTEQUAL, SYM_NOTEQUALCASE // =, ==, !=, !==... Keep this in sync with IS_RELATIONAL_OPERATOR() below.
#define IS_EQUALITY_OPERATOR(symbol) (symbol >= SYM_EQUAL && symbol <= SYM_NOTEQUALCASE)
	, SYM_GT, SYM_LT, SYM_GTOE, SYM_LTOE  // >, <, >=, <= ... Keep this in sync with IS_RELATIONAL_OPERATOR() below.
#define IS_RELATIONAL_OPERATOR(symbol) (symbol >= SYM_EQUAL && symbol <= SYM_LTOE)
	, SYM_REGEXMATCH // ~=, equivalent to a RegExMatch call in two-parameter mode.
	, SYM_CONCAT
	, SYM_LOW_CONCAT // Zero-precedence concat, used so that "x%y=z%" is equivalent to "x%(y=z)%".
	
	// INTEGER OPERATORS START
	// Note that the below macro does not consider SYM_BITNOT or the bit-assignment operators.
	, SYM_BITOR // Seems more intuitive to have these higher in prec. than the above, unlike C and Perl, but like Python.
	, SYM_BITXOR // SYM_BITOR (ABOVE) MUST BE KEPT FIRST AMONG THE INTEGER OPERATORS BECAUSE IT'S USED IN A RANGE-CHECK (see IS_INTEGER_OPERATOR).
	, SYM_BITAND
	, SYM_BITSHIFTLEFT, SYM_BITSHIFTRIGHT, SYM_BITSHIFTRIGHT_LOGICAL // << >> >>>.  <<< is the same as SYM_BITSHIFTLEFT
	, SYM_INTEGERDIVIDE	// eg, x // y ALSO : SYM_INTEGERDIVIDE MUST BE KEPT LAST AMONG THE INTEGER OPERATORS BECAUSE IT'S USED IN A RANGE-CHECK (see IS_INTEGER_OPERATOR).
#define IS_INTEGER_OPERATOR(symbol) ((symbol) <= SYM_INTEGERDIVIDE && (symbol) >= SYM_BITOR) // Currently not considered: SYM_BITNOT and the bit-assignment operators.
	// INTEGER OPERATORS END

	, SYM_ADD, SYM_SUBTRACT
	, SYM_MULTIPLY, SYM_DIVIDE
	, SYM_POWER
	, SYM_LOWNOT  // LOWNOT is the word "not", the low precedence counterpart of !
	, SYM_NEGATIVE, SYM_POSITIVE, SYM_REF, SYM_HIGHNOT, SYM_BITNOT  // Don't change position or order of these because Infix-to-postfix converter's special handling for SYM_POWER relies on them being adjacent to each other.
#define SYM_OVERRIDES_POWER_ON_STACK(symbol) ((symbol) >= SYM_LOWNOT && (symbol) <= SYM_BITNOT) // Check lower bound first for short-circuit performance.
	, SYM_ISSET
	, SYM_PRE_INCREMENT, SYM_PRE_DECREMENT // Must be kept after the post-ops and in this order relative to each other due to a range check in the code.
#define SYM_INCREMENT_OR_DECREMENT_IS_PRE(symbol) ((symbol) >= SYM_PRE_INCREMENT) // Caller has verified symbol is an INCREMENT or DECREMENT operator.
#define IS_PREFIX_OPERATOR(symbol) ((symbol) >= SYM_LOWNOT && (symbol) <= SYM_PRE_DECREMENT)
	, SYM_FUNC     // A call to a function.
	, SYM_RESERVED_WORD, SYM_RESERVED_OPERATOR
#define SYM_IS_RESERVED(symbol) ((symbol) >= SYM_RESERVED_WORD) // No need to exclude SYM_COUNT in this case.
	, SYM_COUNT    // Must be last because it's the total symbol count for everything above.
	, SYM_INVALID = SYM_COUNT // Some callers may rely on YIELDS_AN_OPERAND(SYM_INVALID)==false.
};

// This should include all operators which can produce SYM_VAR for a subsequent assignment:
#define IS_OPERATOR_VALID_LVALUE(sym) \
	(IS_ASSIGNMENT_EXCEPT_POST_AND_PRE(sym) \
		|| sym == SYM_PRE_INCREMENT || sym == SYM_PRE_DECREMENT)


enum VarRefUsageType { VARREF_READ = 0, VARREF_ISSET, VARREF_READ_MAYBE
	, VARREF_REF, VARREF_LVALUE, VARREF_OUTPUT_VAR };
#define VARREF_IS_WRITE(var_usage) ((var_usage) >= VARREF_REF)
#define VARREF_IS_READ(var_usage) ((var_usage) == VARREF_READ || (var_usage) == VARREF_READ_MAYBE) // But not VARREF_ISSET.


struct ExprTokenType; // Forward declarations for use below.
struct ResultToken;
struct IDebugProperties;
class Object;


// Must not be smaller than INT_PTR; see "(IntKeyType)(INT_PTR)".
typedef __int64 IntKeyType;


struct DECLSPEC_NOVTABLE IObject // L31: Abstract interface for "objects".
	: public IDispatch
{
	#define IObject_Invoke_PARAMS_DECL \
		ResultToken &aResultToken, int aFlags, LPTSTR aName, ExprTokenType &aThisToken, ExprTokenType *aParam[], int aParamCount
	#define IObject_Invoke_PARAMS \
		aResultToken, aFlags, aName, aThisToken, aParam, aParamCount
	virtual ResultType Invoke(IObject_Invoke_PARAMS_DECL) = 0;
	virtual LPTSTR Type() = 0;
	#define IObject_Type_Impl(name) \
		LPTSTR Type() { return _T(name); }
	virtual Object *Base() = 0;
	virtual bool IsOfType(Object *aPrototype) = 0;
	
#ifdef CONFIG_DEBUGGER
	#define IObject_DebugWriteProperty_Def \
		void DebugWriteProperty(IDebugProperties *, int aPage, int aPageSize, int aMaxDepth)
	virtual IObject_DebugWriteProperty_Def;
#else
	#define IObject_DebugWriteProperty_Def
#endif
};


struct DECLSPEC_NOVTABLE IObjectComCompatible : public IObject
{
	STDMETHODIMP QueryInterface(REFIID riid, void **ppv);
	//STDMETHODIMP_(ULONG) AddRef() = 0;
	//STDMETHODIMP_(ULONG) Release() = 0;
	STDMETHODIMP GetTypeInfoCount(UINT *pctinfo);
	STDMETHODIMP GetTypeInfo(UINT itinfo, LCID lcid, ITypeInfo **pptinfo);
	STDMETHODIMP GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId);
	STDMETHODIMP Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr);
};


#ifdef CONFIG_DEBUGGER

typedef void *DebugCookie;

struct DECLSPEC_NOVTABLE IDebugProperties
{
	// For simplicity/code size, the debugger handles failures internally
	// rather than returning an error code and requiring caller to handle it.
	virtual void WriteProperty(LPCSTR aName, ExprTokenType &aValue) = 0;
	virtual void WriteProperty(LPCWSTR aName, ExprTokenType &aValue) = 0;
	virtual void WriteProperty(ExprTokenType &aKey, ExprTokenType &aValue) = 0;
	virtual void WriteBaseProperty(IObject *aBase) = 0;
	virtual void WriteDynamicProperty(LPTSTR aName) = 0;
	virtual void WriteEnumItems(IObject *aEnumerable, int aStart, int aEnd) = 0;
	virtual void BeginProperty(LPCSTR aName, LPCSTR aType, int aNumChildren, DebugCookie &aCookie) = 0;
	virtual void EndProperty(DebugCookie aCookie) = 0;
};

#endif


// Flags used when calling Invoke; also used by g_ObjGet etc.:
#define IT_GET				0
#define IT_SET				1
#define IT_CALL				2
#define IT_BITMASK			3 // bit-mask for the above.

#define IF_BYPASS_METAFUNC	0x000010 // Skip invocation of meta-functions, such as when calling __Init or __Delete.
#define IF_SUBSTITUTE_THIS	0x000020 // "this" is a substitute object (i.e. ValueBase()), so prohibit new value properties.
#define IF_SUPER			0x000040 // super.something invocation.
#define IF_DEFAULT			0x000100 // Invoke the default member (call a function object, array indexing, etc.).
#define IF_NEWENUM			0x000200 // Workaround for COM objects which don't resolve "_NewEnum" to DISPID_NEWENUM.

#define EIF_VARIADIC		0x010000
#define EIF_STACK_MEMBER	0x020000
#define EIF_LEAVE_PARAMS	0x040000


// Helper function for event handlers and __Delete:
ResultType CallMethod(IObject *aInvokee, IObject *aThis, LPTSTR aMethodName
	, ExprTokenType *aParamValue = NULL, int aParamCount = 0, __int64 *aRetVal = NULL // For event handlers.
	, int aExtraFlags = 0 // For Object.__Delete().
	, bool aReturnBoolean = false);


struct DerefType; // Forward declarations for use below.
struct CallSite;  //
class Var;        //
struct ExprTokenType  // Something in the compiler hates the name TokenType, so using a different name.
{
	// Due to the presence of 8-byte members (double and __int64) this entire struct is aligned on 8-byte
	// vs. 4-byte boundaries.  The compiler defaults to this because otherwise an 8-byte member might
	// sometimes not start at an even address, which would hurt performance on Pentiums, etc.
	union // Which of its members is used depends on the value of symbol, below.
	{
		__int64 value_int64; // for SYM_INTEGER
		double value_double; // for SYM_FLOAT
		struct
		{
			union // These nested structs and unions minimize the token size by overlapping data.
			{
				IObject *object;
				CallSite *callsite;   // for SYM_FUNC, and (while parsing) SYM_ASSIGN etc.
				DerefType *var_deref; // for SYM_VAR while parsing
				Var *var;             // for SYM_VAR and SYM_DYNAMIC
				LPTSTR marker;        // for SYM_STRING and (while parsing) SYM_OPAREN
				ExprTokenType *circuit_token; // for short-circuit operators
			};
			union // Due to the outermost union, this doesn't increase the total size of the struct on x86 builds (but it does on x64).
			{
				CallSite *outer_param_list; // Used by ExpressionToPostfix().
				LPTSTR error_reporting_marker; // Used by ExpressionToPostfix() for binary and unary operators.
				size_t marker_length;
				VarRefUsageType var_usage; // for SYM_DYNAMIC and SYM_VAR (at load time)
			};
		};  
	};
	SymbolType symbol;


	ExprTokenType() {}
	ExprTokenType(int aValue) { SetValue(aValue); }
	ExprTokenType(__int64 aValue) { SetValue(aValue); }
	ExprTokenType(double aValue) { SetValue(aValue); }
	ExprTokenType(IObject *aValue) { SetValue(aValue); }
	ExprTokenType(LPTSTR aValue, size_t aLength = -1) { SetValue(aValue, aLength); }
	
	void SetValue(__int64 aValue)
	{
		symbol = SYM_INTEGER;
		value_int64 = aValue;
	}
	void SetValue(int aValue) { SetValue((__int64)aValue); }
	void SetValue(UINT aValue) { SetValue((__int64)aValue); }
	void SetValue(UINT64 aValue) { SetValue((__int64)aValue); }
	void SetValue(double aValue)
	{
		symbol = SYM_FLOAT;
		value_double = aValue;
	}
	void SetValue(LPTSTR aValue, size_t aLength = -1)
	{
		ASSERT(aValue);
		symbol = SYM_STRING;
		marker = aValue;
		marker_length = aLength;
	}
	void SetValue(IObject *aValue)
	// Caller must AddRef() if appropriate.
	{
		ASSERT(aValue);
		symbol = SYM_OBJECT;
		object = aValue;
	}

	inline void CopyValueFrom(ExprTokenType &other)
	// Copies the value of a token (by reference where applicable).  Does not object->AddRef().
	{
		value_int64 = other.value_int64; // Union copy.
#ifdef _WIN64
		// For simplicity/smaller code size, don't bother checking symbol == SYM_STRING.
		marker_length = other.marker_length; // Already covered by the above on x86.
#endif
		symbol = other.symbol;
	}

	inline void CopyExprFrom(ExprTokenType &other)
	// Copies all fields typically needed in a postfix expression.
	{
		return CopyValueFrom(other); // Currently nothing needs to be done differently.
	}

	void SetVarRef(Var *aVar)
	{
		symbol = SYM_VAR;
		var = aVar;
		var_usage = VARREF_REF;
	}

	// Assignments yield a variable using this function so that it can be passed ByRef,
	// and because in some cases it avoids an extra memory allocation or string copy
	// (that would otherwise be necessary to ensure the string value is not freed or
	// overwritten by a subsequent concat/function call while still on the stack).
	// Yielding SYM_VAR means subsequent assignments may affect it, but in a safer way
	// that doesn't risk dangling pointers.
	void SetVar(Var *aVar)
	{
		symbol = SYM_VAR;
		var = aVar;
		var_usage = VARREF_READ;
	}

private: // Force code to use one of the CopyFrom() methods, for clarity.
	ExprTokenType & operator = (ExprTokenType &other)
	{
		return *this;
	}
};
#define STACK_PUSH(token_ptr) stack[stack_count++] = token_ptr
#define STACK_POP stack[--stack_count]  // To be used as the r-value for an assignment.

class Object;
class BuiltInFunc;
struct ResultToken : public ExprTokenType
{
	LPTSTR buf; // Points to a buffer of _f_retval_buf_size characters for returning short strings and misc purposes.
	LPTSTR mem_to_free; // Callee stores memory allocated for the result here.  Must be NULL or equal to marker.
#ifdef ENABLE_HALF_BAKED_NAMED_PARAMS
	IObject *named_params; // Variadic callers may pass named parameters via properties of this.
#endif

	// Utility function for initializing result tokens.
	void InitResult(LPTSTR aResultBuf)
	{
		symbol = SYM_STRING;
		marker = _T("");
		marker_length = -1; // Helps code size to do this here instead of in ReturnPtr(), which should be inlined.
		buf = aResultBuf;
		mem_to_free = nullptr;
#ifdef ENABLE_HALF_BAKED_NAMED_PARAMS
		named_params = nullptr;
#endif
		result = OK;
	}

	// Utility function for properly freeing a token's contents.
	void Free()
	{
		// If the token contains an object, release it.
		if (symbol == SYM_OBJECT)
			object->Release();
		// If the token has memory allocated for it, free it.
		if (mem_to_free)
			free(mem_to_free);
	}

	void StealMem(Var *aVar);
	
	void AcceptMem(LPTSTR aNewMem, size_t aLength)
	{
		symbol = SYM_STRING;
		marker = mem_to_free = aNewMem;
		marker_length = aLength;
	}
	
	LPTSTR Malloc(LPTSTR aValue, size_t aLength);

	ResultType Return(LPTSTR aValue, size_t aLength = -1);
	ResultType ReturnPtr(LPTSTR aValue)
	// Return a null-terminated string which is already in persistent memory.
	{
		ASSERT(aValue);
		symbol = SYM_STRING;
		marker = aValue;
		//marker_length is left at its default value, -1.  Caller will call _tcslen().
		return OK;
	}
	ResultType ReturnPtr(LPTSTR aValue, size_t aLength)
	// Return a string which is already in persistent memory.
	{
		SetValue(aValue, aLength);
		return OK;
	}
	ResultType Return(__int64 aValue)
	{
		SetValue(aValue);
		return OK;
	}
	ResultType Return(int aValue) { return Return((__int64)aValue); }
	ResultType Return(UINT aValue) { return Return((__int64)aValue); }
	ResultType Return(DWORD aValue) { return Return((__int64)aValue); }
	ResultType Return(UINT64 aValue) { return Return((__int64)aValue); }
	ResultType Return(double aValue)
	{
		SetValue(aValue);
		return OK;
	}
	ResultType Return(IObject *aValue)
	// Caller must AddRef() if appropriate and must not pass NULL.
	{
		symbol = SYM_OBJECT;
		object = aValue;
		return OK;
	}
	
	ResultType SetExitResult(ResultType aResult)
	{
		ASSERT(aResult == FAIL || aResult == EARLY_EXIT);
		return result = aResult;
	}

	ResultType SetResult(ResultType aResult) // See comments for 'result' below.
	{
		return result = aResult;
	}

	bool Exited()
	{
		return result == FAIL || result == EARLY_EXIT;
	}

	ResultType Result()
	{
		return result;
	}

	ResultType SoftFail()
	{
		SetValue(_T(""), 0);
		// Caller may rely on FAIL to unwind stack, but this->result is still OK.
		return FAIL;
	}
	
	ResultType HardFail()
	{
		return SetExitResult(FAIL);
	}
	
	// Handle the result of calling a global Error function.
	// If FAIL (abort thread), sets internal result to FAIL and returns FAIL.
	// If OK (continue thread), sets return value to "" and returns FAIL.
	ResultType Fail(ResultType aResultOfShowError)
	{
		return aResultOfShowError == OK ? SoftFail() : HardFail();
	}
	
	ResultType Error(LPCTSTR aErrorText);
	ResultType Error(LPCTSTR aErrorText, LPCTSTR aExtraInfo);
	ResultType Error(LPCTSTR aErrorText, Object *aPrototype);
	ResultType Error(LPCTSTR aErrorText, LPCTSTR aExtraInfo, Object *aPrototype);
	ResultType Error(LPCTSTR aErrorText, ExprTokenType &aExtraInfo, Object *aPrototype);
	ResultType MemoryError();
	ResultType UnknownMemberError(ExprTokenType &aObject, int aFlags, LPCTSTR aMember);
	ResultType Win32Error(DWORD aError = GetLastError());
	ResultType ValueError(LPCTSTR aErrorText);
	ResultType ValueError(LPCTSTR aErrorText, LPCTSTR aExtraInfo);
	ResultType TypeError(LPCTSTR aExpectedType, ExprTokenType &aActualValue);
	ResultType ParamError(int aIndex, ExprTokenType *aParam);
	ResultType ParamError(int aIndex, ExprTokenType *aParam, LPCTSTR aExpectedType);
	ResultType ParamError(int aIndex, ExprTokenType *aParam, LPCTSTR aExpectedType, LPCTSTR aFunction);
	
	BuiltInFunc *func; // For maintainability, this is separate from the ExprTokenType union.  Its main uses are func->mID and func->mOutputVars.

private:
	// Currently can't be included in the value union because meta-functions
	// need the EARLY_RETURN result *and* return value passed back.  However,
	// probably best to keep it separate for code size and maintainability.
	// Struct size is a non-issue since there is only one ResultToken per
	// function call on the stack.
	ResultType result;
};


#define BIF_DECL_PARAMS ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount

// The following macro is used for definitions and declarations of built-in functions:
#define BIF_DECL(name) void name(BIF_DECL_PARAMS)

typedef BIF_DECL((* BuiltInFunctionType));


// But the array that goes with these actions is in globaldata.cpp because
// otherwise it would be a little cumbersome to declare the extern version
// of the array in here (since it's only extern to modules other than
// script.cpp):
enum enum_act {
// Seems best to make ACT_INVALID zero so that it will be the ZeroMemory() default within
// any POD structures that contain an action_type field:
  ACT_INVALID = FAIL  // These should both be zero for initialization and function-return-value purposes.
, ACT_ASSIGNEXPR
, ACT_EXPRESSION
// Keep ACT_BLOCK_BEGIN as the first "control flow" action, for range checks with ACT_FIRST_CONTROL_FLOW:
, ACT_BLOCK_BEGIN, ACT_BLOCK_END
, ACT_STATIC
, ACT_HOTKEY_IF // Must be before ACT_FIRST_COMMAND.
, ACT_EXIT // Used with AddLine(), but excluded from the "named" range below so that the function is preferred.
, ACT_IF, ACT_FIRST_NAMED_ACTION = ACT_IF
, ACT_ELSE
, ACT_LOOP, ACT_LOOP_FILE, ACT_LOOP_REG, ACT_LOOP_READ, ACT_LOOP_PARSE
, ACT_FOR, ACT_WHILE, ACT_UNTIL // Keep LOOP, FOR, WHILE and UNTIL together and in this order for range checks in various places.
, ACT_BREAK, ACT_CONTINUE // Keep ACT_FOR..ACT_CONTINUE together for ACT_EXPANDS_ITS_OWN_ARGS.
, ACT_GOTO
, ACT_FIRST_JUMP = ACT_BREAK, ACT_LAST_JUMP = ACT_GOTO // Actions which accept a label name.
, ACT_RETURN
, ACT_TRY, ACT_CATCH, ACT_FINALLY, ACT_THROW // Keep TRY, CATCH and FINALLY together and in this order for range checks.
, ACT_SWITCH, ACT_CASE // Keep ACT_TRY..ACT_CASE together for ACT_EXPANDS_ITS_OWN_ARGS.
, ACT_LAST_NAMED_ACTION = ACT_CASE
// ================================================================================
// All others are not included in g_act, and are only used for misc. purposes:
// ================================================================================
, ACT_MOUSEMOVE, ACT_MOUSECLICK, ACT_MOUSECLICKDRAG // Used by PerformMouse().
};

#define ACT_IS_IF(ActionType) (ActionType == ACT_IF)
#define ACT_IS_LOOP(ActionType) (ActionType >= ACT_LOOP && ActionType <= ACT_WHILE)
#define ACT_IS_LOOP_EXCLUDING_WHILE(ActionType) (ActionType >= ACT_LOOP && ActionType <= ACT_FOR)
#define ACT_IS_LINE_PARENT(ActionType) (ACT_IS_IF(ActionType) || ActionType == ACT_ELSE \
	|| ACT_IS_LOOP(ActionType) || (ActionType >= ACT_TRY && ActionType <= ACT_FINALLY) \
	|| ActionType == ACT_SWITCH)
// The following groups of action types do not need ExpandArgs() called by ExecUntil(),
// for one of the following reasons: 1) action has no args, 2) action's args are
// always fully resolved at load time, 3) action is never executed by ExecUntil(),
// 4) action needs to be the one to call ExpandArgs(), or does so for optimization.
#define ACT_EXPANDS_ITS_OWN_ARGS(ActionType) (ActionType == ACT_ASSIGNEXPR \
	|| (ActionType >= ACT_FOR && ActionType <= ACT_CONTINUE) \
	|| (ActionType >= ACT_TRY && ActionType <= ACT_CASE))

// For convenience in many places.  Must cast to int to avoid loss of negative values.
#define BUF_SPACE_REMAINING ((int)(aBufSize - (aBuf - aBuf_orig)))

// MsgBox timeout value.  This can't be zero because that is used as a failure indicator:
// Also, this define is in this file to prevent problems with mutual
// dependency between script.h and window.h.  Update: It can't be -1 either because
// that value is used to indicate failure by DialogBox():
#define AHK_TIMEOUT -2
#define AHK_TOO_MANY_MSGBOXES -3
// And these to prevent mutual dependency problem between window.h and globaldata.h:
#define MAX_MSGBOXES 7 // Probably best not to change this because it's used by OurTimers to set the timer IDs, which should probably be kept the same for backward compatibility.
#define MAX_MSG_MONITORS 500

// IMPORTANT: Before ever changing the below, note that it will impact the IDs of menu items created
// with the MENU command, as well as the number of such menu items that are possible (currently about
// 65500-11000=54500). See comments at ID_USER_FIRST for details:
#define GUI_CONTROL_BLOCK_SIZE 1000
#define MAX_CONTROLS_PER_GUI (GUI_CONTROL_BLOCK_SIZE * 11) // Some things rely on this being less than 0xFFFF and an even multiple of GUI_CONTROL_BLOCK_SIZE.
#define NO_CONTROL_INDEX MAX_CONTROLS_PER_GUI // Must be 0xFFFF or less.
#define NO_EVENT_INFO 0 // For backward compatibility with documented contents of A_EventInfo, this should be kept as 0 vs. something more special like UINT_MAX.

#define MAX_TOOLTIPS 20
#define MAX_TOOLTIPS_STR _T("20")   // Keep this in sync with above.
#define MAX_FILEDIALOGS 4
#define MAX_FOLDERDIALOGS 4

#define MAX_NUMBER_LENGTH 255                   // Large enough to allow custom zero or space-padding via %10.2f, etc.
#define MAX_NUMBER_SIZE (MAX_NUMBER_LENGTH + 1) // But not too large because some things might rely on this being fairly small.
#define MAX_INTEGER_LENGTH 20                     // Max length of a 64-bit number when expressed as decimal or
#define MAX_INTEGER_SIZE (MAX_INTEGER_LENGTH + 1) // hex string; e.g. -9223372036854775808 or (unsigned) 18446744073709551616 or (hex) -0xFFFFFFFFFFFFFFFF.

#define VARLIST_INITIAL_SIZE 32

#define SW_NONE -1

// Hot-strings:
// memmove() and proper detection of long hotstrings rely on buf being at least this large:
#define HS_BUF_SIZE (MAX_HOTSTRING_LENGTH * 2 + 10)
#define HS_BUF_DELETE_COUNT (HS_BUF_SIZE / 2)
#define HS_MAX_END_CHARS 100

// Bitwise storage of boolean flags.  This section is kept in this file because
// of mutual dependency problems between hook.h and other header files:
typedef UCHAR HookType;
#define HOOK_KEYBD 0x01
#define HOOK_MOUSE 0x02
#define HOOK_FAIL  0xFF

#define EXTERN_G extern global_struct *g
#define EXTERN_OSVER extern OS_Version g_os
#define EXTERN_CLIPBOARD extern Clipboard g_clip
#define EXTERN_SCRIPT extern Script g_script
#define CLIPBOARD_CONTAINS_ONLY_FILES (!IsClipboardFormatAvailable(CF_NATIVETEXT) && IsClipboardFormatAvailable(CF_HDROP))


// These macros used to keep app responsive during a long operation.  In v1.0.39, the
// hooks have a dedicated thread.  However, mLastPeekTime is still compared to 5 rather
// than some higher value for the following reasons:
// 1) Want hotkeys pressed during a long operation to take effect as quickly as possible.
//    For example, in games a hotkey's response time is critical.
// 2) Benchmarking shows less than a 0.5% performance improvement from this comparing
//    to a higher value (even one as high as 500), even when the system is under heavy
//    load from other processes).
//
// mLastPeekTime is global/static so that recursive functions, such as FileSetAttrib(),
// will sleep as often as intended even if the target files require frequent recursion.
// The use of a global/static is not friendly to recursive calls to the function (i.e. calls
// made as a consequence of the current script subroutine being interrupted by another during
// this instance's MsgSleep()).  However, it doesn't seem to be that much of a consequence
// since the exact interval/period of the MsgSleep()'s isn't that important.  It's also
// pretty unlikely that the interrupting subroutine will also just happen to call the same
// function rather than some other.
//
// Older comment that applies if there is ever again no dedicated thread for the hooks:
// These macros were greatly simplified when it was discovered that PeekMessage(), when called
// directly as below, is enough to prevent keyboard and mouse lag when the hooks are installed
#define LONG_OPERATION_INIT MSG msg; DWORD tick_now;

// MsgSleep() is used rather than SLEEP_WITHOUT_INTERRUPTION to allow other hotkeys to
// launch and interrupt (suspend) the operation.  It seems best to allow that, since
// the user may want to press some fast window activation hotkeys, for example, during
// the operation.  The operation will be resumed after the interrupting subroutine finishes.
// Notes applying to the macro:
// Store tick_now for use later, in case the Peek() isn't done, though not all callers need it later.
// ...
// Since the Peek() will yield when there are no messages, it will often take 20ms or more to return
// (UPDATE: this can't be reproduced with simple tests, so either the OS has changed through service
// packs, or Peek() yields only when the OS detects that the app is calling it too often or calling
// it in certain ways [PM_REMOVE vs. PM_NOREMOVE seems to make no difference: either way it doesn't yield]).
// Therefore, must update tick_now again (its value is used by macro and possibly by its caller)
// to avoid having to Peek() immediately after the next iteration.
// ...
// The code might bench faster when "g_script.mLastPeekTime = tick_now" is a separate operation rather
// than combined in a chained assignment statement.
#define LONG_OPERATION_UPDATE \
{\
	tick_now = GetTickCount();\
	if (tick_now - g_script.mLastPeekTime > ::g->PeekFrequency)\
	{\
		if (PeekMessage(&msg, NULL, 0, 0, PM_NOREMOVE))\
			MsgSleep(-1);\
		tick_now = GetTickCount();\
		g_script.mLastPeekTime = tick_now;\
	}\
}

// Same as the above except for SendKeys() and related functions (uses SLEEP_WITHOUT_INTERRUPTION vs. MsgSleep).
#define LONG_OPERATION_UPDATE_FOR_SENDKEYS \
{\
	tick_now = GetTickCount();\
	if (tick_now - g_script.mLastPeekTime > ::g->PeekFrequency)\
	{\
		if (PeekMessage(&msg, NULL, 0, 0, PM_NOREMOVE))\
			SLEEP_WITHOUT_INTERRUPTION(-1) \
		tick_now = GetTickCount();\
		g_script.mLastPeekTime = tick_now;\
	}\
}

// Defining these here avoids awkwardness due to the fact that globaldata.cpp
// does not (for design reasons) include globaldata.h:
typedef UCHAR ActionTypeType; // If ever have more than 256 actions, will have to change this (but it would increase code size due to static data in g_act).
#pragma pack(push, 1) // v1.0.45: Reduces code size a little without impacting runtime performance because this struct is hardly ever accessed during runtime.
struct Action
{
	LPTSTR Name;
	// Changing these from ints to chars greatly reduced code size since this struct is used
	// by g_act to build static data into the code.  Testing shows that the compiler will
	// generate a warning even when not in debug mode in the unlikely event that a constant
	// larger than 127 is ever stored in one of these:
	char MinParams, MaxParams;
};
#pragma pack(pop)

// Values are hard-coded for some of the below because they must not deviate from the documented, numerical
// TitleMatchModes:
enum TitleMatchModes {MATCHMODE_INVALID = FAIL, FIND_IN_LEADING_PART = 1, FIND_ANYWHERE = 2, FIND_EXACT = 3, FIND_REGEX, FIND_FAST, FIND_SLOW};

typedef UINT GuiIndexType; // Some things rely on it being unsigned to avoid the need to check for less-than-zero.
typedef UINT GuiEventType; // Made a UINT vs. enum so that illegal/underflow/overflow values are easier to detect.

// The following array and enum must be kept in sync with each other:
#define GUI_EVENT_NAMES { _T("") \
	, _T("DropFiles"), _T("Close"), _T("Escape"), _T("Size"), _T("ContextMenu") \
	, _T("Change") \
	, _T("Click"), _T("DoubleClick"), _T("ColClick") \
	, _T("ItemCheck"), _T("ItemSelect"), _T("ItemFocus"), _T("ItemExpand") \
	, _T("ItemEdit") \
	, _T("Focus"), _T("LoseFocus") \
}
enum GuiEventTypes {GUI_EVENT_NONE  // NONE must be zero for any uses of ZeroMemory(), synonymous with false, etc.
	, GUI_EVENT_DROPFILES, GUI_EVENT_CLOSE, GUI_EVENT_ESCAPE, GUI_EVENT_RESIZE, GUI_EVENT_CONTEXTMENU
	, GUI_EVENT_WINDOW_FIRST = GUI_EVENT_DROPFILES, GUI_EVENT_WINDOW_LAST = GUI_EVENT_CONTEXTMENU
	, GUI_EVENT_CONTROL_FIRST
	, GUI_EVENT_CHANGE = GUI_EVENT_CONTROL_FIRST
	, GUI_EVENT_CLICK, GUI_EVENT_DBLCLK, GUI_EVENT_COLCLK
	, GUI_EVENT_ITEMCHECK, GUI_EVENT_ITEMSELECT, GUI_EVENT_ITEMFOCUS, GUI_EVENT_ITEMEXPAND
	, GUI_EVENT_ITEMEDIT
	, GUI_EVENT_FOCUS, GUI_EVENT_LOSEFOCUS
	, GUI_EVENT_NAMED_COUNT
	// The rest don't have explicit names in GUI_EVENT_NAMES:
	, GUI_EVENT_WM_COMMAND = GUI_EVENT_NAMED_COUNT
};

enum GuiEventKinds {GUI_EVENTKIND_EVENT = 0, GUI_EVENTKIND_NOTIFY, GUI_EVENTKIND_COMMAND, GUI_EVENTKIND_MESSAGE};

typedef USHORT CoordModeType;

#define COORD_MODE_INVALID ((CoordModeType) -1)
// Bit-field offsets:
#define COORD_MODE_PIXEL   0
#define COORD_MODE_MOUSE   2
#define COORD_MODE_TOOLTIP 4
#define COORD_MODE_CARET   6
#define COORD_MODE_MENU    8

#define COORD_MODE_CLIENT  0
#define COORD_MODE_WINDOW  1
#define COORD_MODE_SCREEN  2
#define COORD_MODE_MASK    3
#define COORD_MODES { _T("Client"), _T("Window"), _T("Screen") }

#define COORD_CENTERED (INT_MIN + 1)
#define COORD_UNSPECIFIED INT_MIN
#define COORD_UNSPECIFIED_SHORT SHRT_MIN  // This essentially makes coord -32768 "reserved", but it seems acceptable given usefulness and the rarity of a real coord like that.

typedef UINT_PTR EventInfoType;

typedef UCHAR SendLevelType;
// Setting the max level to 100 is somewhat arbitrary. It seems that typical usage would only
// require a few levels at most. We do want to keep the max somewhat small to keep the range
// for magic values that get used in dwExtraInfo to a minimum, to avoid conflicts with other
// apps that may be using the field in other ways.
const SendLevelType SendLevelMax = 100;
// Using int as the type for level so this can be used as validation before converting to SendLevelType.
inline bool SendLevelIsValid(int level) { return level >= 0 && level <= SendLevelMax; }


class Line; // Forward declaration.
typedef UCHAR HotCriterionType;
enum HotCriterionEnum {HOT_NO_CRITERION, HOT_IF_ACTIVE, HOT_IF_NOT_ACTIVE, HOT_IF_EXIST, HOT_IF_NOT_EXIST // HOT_NO_CRITERION must be zero.
	, HOT_IF_CALLBACK};
#define HOT_IF_REQUIRES_EVAL(type) ((type) == HOT_IF_CALLBACK)
struct HotkeyCriterion
{
	HotCriterionType Type;
	LPTSTR WinTitle, WinText;
	LPTSTR OriginalExpr; // For finding expr in #HotIf expr
	IObject *Callback;
	HotkeyCriterion *NextCriterion, *NextExpr;

	ResultType Eval(LPTSTR aHotkeyName); // For HOT_IF_CALLBACK.
};


class Func;                 // Forward declarations
class UserFunc;             //
class Label;                //
struct RegItemStruct;       //
struct LoopFilesStruct;     //
struct LoopReadFileStruct;  //
class GuiType;				//
class ScriptTimer;			//

struct ScriptThreadState
{
	__int64 mLoopIteration; // Signed, since script/ITOA64 aren't designed to handle unsigned.
	LoopFilesStruct *mLoopFile;  // The file of the current file-loop, if applicable.
	RegItemStruct *mLoopRegItem; // The registry subkey or value of the current registry enumeration loop.
	LoopReadFileStruct *mLoopReadFile;  // The file whose contents are currently being read by a File-Read Loop.
	LPTSTR mLoopField;  // The field of the current string-parsing loop.

	UserFunc *CurrentFunc; // v1.0.46.16: The function whose body is currently being processed at load-time, or being run at runtime (if any).
	ScriptTimer *CurrentTimer; // The timer that launched this thread (if any).
	HWND hWndLastUsed;  // In many cases, it's better to use GetValidLastUsedWindow() when referring to this.
	EventInfoType EventInfo;
	HWND DialogHWND; // MsgBox being shown by this thread.
	HWND DialogOwner; // This thread's dialog owner, if any.
#define THREAD_DIALOG_OWNER (IsWindow(::g->DialogOwner) ? ::g->DialogOwner : (::g->DialogOwner = NULL)) // Reset to NULL if invalid to mitigate the risk of errors due to HWND reuse by the OS.
	ResultToken* ThrownToken;

	int ExcptMode;
	DWORD LastError; // The result of GetLastError() after the most recent DllCall or Run.
	int Priority;  // This thread's priority relative to others.
	int UninterruptedLineCount; // Stored as a g-struct attribute in case OnExit func interrupts it while uninterruptible.
	int UninterruptibleDuration; // Must be int to preserve negative values found in g_script.mUninterruptibleTime.
	DWORD ThreadStartTime;
	DWORD CalledByIsDialogMessageOrDispatchMsg; // Detects the fact that some messages (like WM_KEYDOWN->WM_NOTIFY for UpDown controls) are translated to different message numbers by IsDialogMessage (and maybe Dispatch too).

	bool IsPaused;
	bool MsgBoxTimedOut; // Meaningful only while a MsgBox call is in progress.
	bool CalledByIsDialogMessageOrDispatch; // Helps avoid launching a monitor function twice for the same message.  This would probably be okay if it were a normal global rather than in the g-struct, but due to messaging complexity, this lends peace of mind and robustness.
	bool AllowThreadToBeInterrupted; // Whether this thread can be interrupted by custom menu items, hotkeys, or timers.  Separate from g_AllowInterruption because that's for use by ongoing operations, such as SendKeys, and should override the thread's setting.
};

struct ScriptThreadSettings
{
	HotkeyCriterion *HotCriterion;

	DWORD PeekFrequency; // DWORD vs. UCHAR might improve performance a little since it's checked so often.
	TitleMatchModes TitleMatchMode;
	int WinDelay;  // negative values may be used as special flags.
	int ControlDelay; // negative values may be used as special flags.
	int KeyDelay;     //
	int KeyDelayPlay; //
	int PressDuration;     // The delay between the up-event and down-event of each keystroke.
	int PressDurationPlay; // 
	int MouseDelay;     // negative values may be used as special flags.
	int MouseDelayPlay; //
	DWORD RegView;
	SendModes SendMode;
	UINT Encoding;

	CoordModeType CoordMode; // Bitwise collection of flags.

	// All these one-byte members are kept adjacent to make the struct smaller, which helps conserve stack space:
	bool TitleFindFast; // Whether to use the fast mode of searching window text, or the more thorough slow mode.
	bool DetectHiddenWindows; // Whether to detect the titles of hidden parent windows.
	bool DetectHiddenText;    // Whether to detect the text of hidden child windows.
	bool AllowTimers; // v1.0.40.01 Whether new timer threads are allowed to start during this thread.
	bool ThreadIsCritical; // Whether this thread has been marked (un)interruptible by the "Critical" command.
	UCHAR DefaultMouseSpeed;
	bool StoreCapslockMode;
	SendLevelType SendLevel;
	bool ListLinesIsEnabled;

	//inline bool InTryBlock() { return ExcptMode & EXCPTMODE_TRY; } // Currently unused.
	bool DetectWindow(HWND aWnd);
};

// global_struct is a combination of thread state (things specific to a thread that
// are reset for each new thread and should not be affected by the default thread)
// and thread settings (which are copied from the default thread).
// Each instance of this struct generally corresponds to a quasi-thread.
struct global_struct : public ScriptThreadState, public ScriptThreadSettings { };

inline void global_maximize_interruptibility(ScriptThreadState &g)
{
	//g.AllowThreadToBeInterrupted = true; // Not necessary since its value isn't used when !g_nThreads.
	#define PRIORITY_MINIMUM INT_MIN
	g.Priority = PRIORITY_MINIMUM; // Ensure minimum priority so that it can always be interrupted.
	// The following can't be reset because their values (as set by the auto-execute section) are
	// used by new threads.  Instead, interruption and timers are always permitted when !g_nThreads.
	//g.ThreadIsCritical = false;
	//g.AllowTimers = true;
}

inline void global_clear_state(ScriptThreadState &g)
// Reset those values that represent the condition or state created by previously executed commands
// but that shouldn't be retained for future threads (e.g. SetTitleMatchMode should be retained for
// future threads if it occurs in the auto-execute section, but A_ThisFunc shouldn't).
{
	ZeroMemory(&g, sizeof(g));
	g.AllowThreadToBeInterrupted = true;
	// All of the following are handled by zero-initialization above:
	//g.CurrentFunc = NULL;
	//g.hWndLastUsed = NULL;
	//g.IsPaused = false;
	//g.Priority = 0;
	//g.UninterruptedLineCount = 0;
	//g.DialogHWND = NULL;
	//g.DialogOwner = NULL;
	//g.CalledByIsDialogMessageOrDispatch = false; // CalledByIsDialogMessageOrDispatchMsg doesn't need to be cleared because it's value is only considered relevant when CalledByIsDialogMessageOrDispatch==true.
	// Above line is done because allowing it to be permanently changed by the auto-exec section
	// seems like it would cause more confusion that it's worth.  A change to the global default
	// or even an override/always-use-this-window-number mode can be added if there is ever a
	// demand for it.
	//g.mLoopIteration = 0; // Zero seems preferable to 1, to indicate "no loop currently running" when a thread first starts off.  This should probably be left unchanged for backward compatibility (even though script's aren't supposed to rely on it).
	//g.mLoopFile = NULL;
	//g.mLoopRegItem = NULL;
	//g.mLoopReadFile = NULL;
	//g.mLoopField = NULL;
	//g.ThrownToken = NULL;
	//g.ExcptMode = EXCPTMODE_NONE;
	//g.LastError = 0;
	//g.EventInfo = NO_EVENT_INFO;
}

inline void global_set_defaults(ScriptThreadSettings &g)
{
	// Init struct with application defaults.  They're in a struct so that it's easier
	// to save and restore their values when one hotkey interrupts another, going into
	// deeper recursion.  When the interrupting subroutine returns, the former
	// subroutine's values for these are restored prior to resuming execution:
	g.HotCriterion = NULL;
	g.SendMode = SM_INPUT;
	g.TitleMatchMode = FIND_ANYWHERE;
	g.TitleFindFast = true; // Since it's so much faster in many cases.
	g.DetectHiddenWindows = false;  // Same as AutoIt2 but unlike AutoIt3; seems like a more intuitive default.
	g.DetectHiddenText = true;  // Unlike AutoIt, which defaults to false.  This setting performs better.
	#define DEFAULT_PEEK_FREQUENCY 5
	g.PeekFrequency = DEFAULT_PEEK_FREQUENCY; // v1.0.46. See comments in Critical().
	g.AllowTimers = true;
	g.ThreadIsCritical = false;
	g.WinDelay = 100;
	g.ControlDelay = 20;
	g.KeyDelay = 10;
	g.KeyDelayPlay = -1;
	g.PressDuration = -1;
	g.PressDurationPlay = -1;
	g.MouseDelay = 10;
	g.MouseDelayPlay = -1;
	#define DEFAULT_MOUSE_SPEED 2
	#define MAX_MOUSE_SPEED 100
	g.DefaultMouseSpeed = DEFAULT_MOUSE_SPEED;
	g.CoordMode = 0;  // All the flags it contains are off by default.
	g.StoreCapslockMode = true;  // AutoIt2 (and probably 3's) default, and it makes a lot of sense.
	g.SendLevel = 0;
	g.ListLinesIsEnabled = true;
	g.Encoding = CP_ACP;
}

// Initialize g and set application defaults.  This is called only once, since new
// threads get a copy of the current settings of the auto-execute thread (g_array[0])
// combined with a cleared state.
inline void global_init(global_struct &g)
{
	global_clear_state(g);
	global_set_defaults(g);
}


#ifdef UNICODE
#define WINAPI_SUFFIX "W"
#define PROCESS_API_SUFFIX "W" // used by Process32First and Process32Next
#else
#define WINAPI_SUFFIX "A"
#define PROCESS_API_SUFFIX
#endif

#define _TSIZE(a) ((a)*sizeof(TCHAR))
#define CP_AHKNOBOM 0x80000000
#define CP_AHKCP    (~CP_AHKNOBOM)

#ifndef CP_UTF16
#define CP_UTF16 1200
#endif

// Use #pragma message(MY_WARN(nnnn) "warning messages") to generate a warning like a compiler's warning
#define __S(x) #x
#define _S(x) __S(x)
#define MY_WARN(n) __FILE__ "(" _S(__LINE__) ") : warning C" __S(n) ": "

// These will be removed when things are done.
#ifdef CONFIG_UNICODE_CHECK
#define UNICODE_CHECK __declspec(deprecated(_T("Please check what you want are bytes or characters.")))
UNICODE_CHECK inline size_t CHECK_SIZEOF(size_t n) { return n; }
#define SIZEOF(c) CHECK_SIZEOF(sizeof(c))
#pragma deprecated(memcpy, memset, memmove, malloc, realloc, _alloca, alloca, toupper, tolower)
#else
#define UNICODE_CHECK
#endif

#endif
