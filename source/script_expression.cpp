/*
AutoHotkey

Copyright 2003-2008 Chris Mallett (support@autohotkey.com)

This program is free software; you can redistribute it and/or
modify it under the terms of the GNU General Public License
as published by the Free Software Foundation; either version 2
of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.
*/

//////////////////////////////////////////////////////////////////////////////////////
// v1.0.40.02: This is now a separate file to allow its compiler optimization settings
// to be set independently of those of the other modules.  In one benchmark, this
// improved performance of expressions and function calls by 9% (that is, when the
// other modules are set to "minmize size" such as for the AutoHotkeySC.bin file).
// This gain in performance is at the cost of a 1.5 KB increase in the size of the
// compressed code, which seems well worth it given how often expressions and
// function-calls are used (such as in loops).
//
// ExpandArgs() and related functions were also put into this file because that
// further improves performance across the board -- even for AutoHotkey.exe despite
// the fact that the only thing that changed for it was the module move, not the
// compiler settings.  Apparently, the butterfly effect can cause even minor
// modifications to impact the overall performance of the generated code by as much as
// 7%.  However, this might have more to do with cache hits and misses in the CPU than
// with the nature of the code produced by the compiler.
// UPDATE 10/18/2006: There's not much difference anymore -- in fact, using min size
// for everything makes compiled scripts slightly faster in basic benchmarks, probably
// due to the recent addition of the linker optimization that physically orders
// functions in a better order inside the EXE.  Therefore, script_expression.cpp no
// longer has a separate "favor fast code" option.
//////////////////////////////////////////////////////////////////////////////////////

#include "stdafx.h" // pre-compiled headers
#include "script.h"
#include "globaldata.h" // for a lot of things
#include "qmath.h" // For ExpandExpression()

// __forceinline: Decided against it for this function because alhough it's only called by one caller,
// testing shows that it wastes stack space (room for its automatic variables would be unconditionally 
// reserved in the stack of its caller).  Also, the performance benefit of inlining this is too slight.
// Here's a simple way to verify wasted stack space in a caller that calls an inlined function:
//    DWORD stack
//    _asm mov stack, esp
//    MsgBox(stack);
char *Line::ExpandExpression(int aArgIndex, ResultType &aResult, char *&aTarget, char *&aDerefBuf
	, size_t &aDerefBufSize, char *aArgDeref[], size_t aExtraSize)
// Caller should ignore aResult unless this function returns NULL.
// Returns a pointer to this expression's result, which can be one of the following:
// 1) NULL, in which case aResult will be either FAIL or EARLY_EXIT to indicate the means by which the current
//    quasi-thread was terminated as a result of a function call.
// 2) The constant empty string (""), in which case we do not alter aTarget for our caller.
// 3) Some persistent location not in aDerefBuf, namely the mContents of a variable or a literal string/number,
//    such as a function-call that returns "abc", 123, or a variable.
// 4) At position aTarget inside aDerefBuf (note that aDerefBuf might have been reallocated by us).
// aTarget is left unchnaged except in case #4, in which case aTarget has been adjusted to the position after our
// result-string's terminator.  In addition, in case #4, aDerefBuf, aDerefBufSize, and aArgDeref[] have been adjusted
// for our caller if aDerefBuf was too small and needed to be enlarged.
//
// Thanks to Joost Mulders for providing the expression evaluation code upon which this function is based.
{
	char *target = aTarget; // "target" is used to track our usage (current position) within the aTarget buffer.

	// The following must be defined early so that mem_count is initialized and the array is guaranteed to be
	// "in scope" in case of early "goto" (goto substantially boosts performance and reduces code size here).
	#define MAX_EXPR_MEM_ITEMS 200 // v1.0.47.01: Raised from 100 because a line consisting entirely of concat operators can exceed it.  However, there's probably not much point to going much above MAX_TOKENS/2 because then it would reach the MAX_TOKENS limit first.
	char *mem[MAX_EXPR_MEM_ITEMS]; // No init necessary.  In most cases, it will never be used.
	int mem_count = 0; // The actual number of items in use in the above array.
	char *result_to_return = ""; // By contrast, NULL is used to tell the caller to abort the current thread.  That isn't done for normal syntax errors, just critical conditions such as out-of-memory.
	Var *output_var = (mActionType == ACT_ASSIGNEXPR) ? OUTPUT_VAR : NULL; // Resolve early because it's similar in usage/scope to the above.  Plus MUST be resolved prior to calling any script-functions since they could change the values in sArgVar[].

	// Having a precedence array is required at least for SYM_POWER (since the order of evaluation
	// of something like 2**1**2 does matter).  It also helps performance by avoiding unnecessary pushing
	// and popping of operators to the stack. This array must be kept in sync with "enum SymbolType".
	// Also, dimensioning explicitly by SYM_COUNT helps enforce that at compile-time:
	static UCHAR sPrecedence[SYM_COUNT] =  // Performance: UCHAR vs. INT benches a little faster, perhaps due to the slight reduction in code size it causes.
	{
		0,0,0,0,0,0,0    // SYM_STRING, SYM_INTEGER, SYM_FLOAT, SYM_VAR, SYM_OPERAND, SYM_DYNAMIC, SYM_BEGIN (SYM_BEGIN must be lowest precedence).
		, 82, 82         // SYM_POST_INCREMENT, SYM_POST_DECREMENT: Highest precedence operator so that it will work even though it comes *after* a variable name (unlike other unaries, which come before).
		, 4, 4           // SYM_CPAREN, SYM_OPAREN (to simplify the code, parentheses must be lower than all operators in precedence).
		, 6              // SYM_COMMA -- Must be just above SYM_OPAREN so it doesn't pop OPARENs off the stack.
		, 7,7,7,7,7,7,7,7,7,7,7,7  // SYM_ASSIGN_*. THESE HAVE AN ODD NUMBER to indicate right-to-left evaluation order, which is necessary for cascading assignments such as x:=y:=1 to work.
//		, 8              // THIS VALUE MUST BE LEFT UNUSED so that the one above can be promoted to it by the infix-to-postfix routine.
		, 11, 11         // SYM_IFF_ELSE, SYM_IFF_THEN (ternary conditional).  HAS AN ODD NUMBER to indicate right-to-left evaluation order, which is necessary for ternaries to perform traditionally when nested in each other without parentheses.
//		, 12             // THIS VALUE MUST BE LEFT UNUSED so that the one above can be promoted to it by the infix-to-postfix routine.
		, 16             // SYM_OR
		, 20             // SYM_AND
		, 25             // SYM_LOWNOT (the word "NOT": the low precedence version of logical-not).  HAS AN ODD NUMBER to indicate right-to-left evaluation order so that things like "not not var" are supports (which can be used to convert a variable into a pure 1/0 boolean value).
//		, 26             // THIS VALUE MUST BE LEFT UNUSED so that the one above can be promoted to it by the infix-to-postfix routine.
		, 30, 30, 30     // SYM_EQUAL, SYM_EQUALCASE, SYM_NOTEQUAL (lower prec. than the below so that "x < 5 = var" means "result of comparison is the boolean value in var".
		, 34, 34, 34, 34 // SYM_GT, SYM_LT, SYM_GTOE, SYM_LTOE
		, 38             // SYM_CONCAT
		, 42             // SYM_BITOR -- Seems more intuitive to have these three higher in prec. than the above, unlike C and Perl, but like Python.
		, 46             // SYM_BITXOR
		, 50             // SYM_BITAND
		, 54, 54         // SYM_BITSHIFTLEFT, SYM_BITSHIFTRIGHT
		, 58, 58         // SYM_ADD, SYM_SUBTRACT
		, 62, 62, 62     // SYM_MULTIPLY, SYM_DIVIDE, SYM_FLOORDIVIDE
		, 67,67,67,67,67 // SYM_NEGATIVE (unary minus), SYM_HIGHNOT (the high precedence "!" operator), SYM_BITNOT, SYM_ADDRESS, SYM_DEREF
		// NOTE: THE ABOVE MUST BE AN ODD NUMBER to indicate right-to-left evaluation order, which was added in v1.0.46 to support consecutive unary operators such as !*var !!var (!!var can be used to convert a value into a pure 1/0 boolean).
//		, 68             // THIS VALUE MUST BE LEFT UNUSED so that the one above can be promoted to it by the infix-to-postfix routine.
		, 72             // SYM_POWER (see note below).  Associativity kept as left-to-right for backward compatibility (e.g. 2**2**3 is 4**3=64 not 2**8=256).
		, 77, 77         // SYM_PRE_INCREMENT, SYM_PRE_DECREMENT (higher precedence than SYM_POWER because it doesn't make sense to evaluate power first because that would cause ++/-- to fail due to operating on a non-lvalue.
//		, 78             // THIS VALUE MUST BE LEFT UNUSED so that the one above can be promoted to it by the infix-to-postfix routine.
//		, 82, 82         // RESERVED FOR SYM_POST_INCREMENT, SYM_POST_DECREMENT (which are listed higher above for the performance of YIELDS_AN_OPERAND().
		, 86             // SYM_FUNC -- Must be of highest precedence so that it stays tightly bound together as though it's a single operand for use by other operators.
	};
	// Most programming languages give exponentiation a higher precedence than unary minus and logical-not.
	// For example, -2**2 is evaluated as -(2**2), not (-2)**2 (the latter is unsupported by qmathPow anyway).
	// However, this rule requires a small workaround in the postfix-builder to allow 2**-2 to be
	// evaluated as 2**(-2) rather than being seen as an error.  v1.0.45: A similar thing is required
	// to allow the following to work: 2**!1, 2**not 0, 2**~0xFFFFFFFE, 2**&x.
	// On a related note, the right-to-left tradition of something like 2**3**4 is not implemented (maybe in v2).
	// Instead, the expression is evaluated from left-to-right (like other operators) to simplify the code.

	#define MAX_TOKENS 512 // Max number of operators/operands.  Seems enough to handle anything realistic, while conserving call-stack space.
	ExprTokenType infix[MAX_TOKENS], *postfix[MAX_TOKENS], *stack[MAX_TOKENS + 1];  // +1 for SYM_BEGIN on the stack.
	int infix_count = 0, postfix_count = 0, stack_count = 0;
	// Above dimensions the stack to be as large as the infix/postfix arrays to cover worst-case
	// scenarios and avoid having to check for overflow.  For the infix-to-postfix conversion, the
	// stack must be large enough to hold a malformed expression consisting entirely of operators
	// (though other checks might prevent this).  It must also be large enough for use by the final
	// expression evaluation phase, the worst case of which is unknown but certainly not larger
	// than MAX_TOKENS.

	///////////////////////////////////////////////////////////////////////////////////////////////
	// TOKENIZE THE INFIX EXPRESSION INTO AN INFIX ARRAY: Avoids the performance overhead of having
	// to re-detect whether each symbol is an operand vs. operator at multiple stages.
	///////////////////////////////////////////////////////////////////////////////////////////////
	// In v1.0.46.01, this section was simplified to avoid transcribing the entire expression into the
	// deref buffer.  In addition to improving performance and reducing code size, this also solves
	// obscure timing bugs caused by functions that have side-effects, especially in comma-separated
	// sub-expressions.  In these cases, one part of an expression could change a built-in variable
	// (indirectly or in the case of Clipboard, directly), an environment variable, or a double-def.
	// For example the dynamic components of a double-deref can be changed by other parts of an
	// expression, even one without commas.  Another example is: fn(clipboard, func_that_changes_clip()).
	// So now, built-in & environment variables and double-derefs are resolve when they're actually
	// encountered during the final/evaluation phase.
	// Another benefit to deferring the resolution of these types of items is that they become eligible
	// for short-circuiting, which further helps performance (they're quite similar to built-in
	// functions in this respect).
	char *op_end, *cp;
	DerefType *deref, *this_deref, *deref_start, *deref_alloca;
	int derefs_in_this_double;
	int cp1; // int vs. char benchmarks slightly faster, and is slightly smaller in code size.

	for (cp = mArg[aArgIndex].text, deref = mArg[aArgIndex].deref // Start at the begining of this arg's text and look for the next deref.
		;; ++deref, ++infix_count) // FOR EACH DEREF IN AN ARG:
	{
		this_deref = deref && deref->marker ? deref : NULL; // A deref with a NULL marker terminates the list (i.e. the final deref isn't a deref, merely a terminator of sorts.

		// BEFORE PROCESSING "this_deref" ITSELF, MUST FIRST PROCESS ANY LITERAL/RAW TEXT THAT LIES TO ITS LEFT.
		if (this_deref && cp < this_deref->marker // There's literal/raw text to the left of the next deref.
			|| !this_deref && *cp) // ...or there's no next deref, but there's some literal raw text remaining to be processed.
		{
			for (;; ++infix_count) // FOR EACH TOKEN INSIDE THIS RAW/LITERAL TEXT SECTION.
			{
				// Because neither the postfix array nor the stack can ever wind up with more tokens than were
				// contained in the original infix array, only the infix array need be checked for overflow:
				if (infix_count > MAX_TOKENS - 1) // No room for this operator or operand to be added.
					goto abnormal_end;

				// Only spaces and tabs are considered whitespace, leaving newlines and other whitespace characters
				// for possible future use:
				cp = omit_leading_whitespace(cp);
				if (!*cp // Very end of expression...
					|| this_deref && cp >= this_deref->marker) // ...or no more literal/raw text left to process at the left side of this_deref.
					break; // Break out of inner loop so that bottom of the outer loop will process this_deref itself.

				ExprTokenType &this_infix_item = infix[infix_count]; // Might help reduce code size since it's referenced many places below.

				// CHECK IF THIS CHARACTER IS AN OPERATOR.
				cp1 = cp[1]; // Improves performance by nearly 5% and appreciably reduces code size (at the expense of being less maintainable).
				switch (*cp)
				{
				// The most common cases are kept up top to enhance performance if switch() is implemented as if-else ladder.
				case '+':
					if (cp1 == '=')
					{
						++cp; // An additional increment to have loop skip over the operator's second symbol.
						this_infix_item.symbol = SYM_ASSIGN_ADD;
					}
					else
					{
						if (infix_count && YIELDS_AN_OPERAND(infix[infix_count - 1].symbol))
						{
							if (cp1 == '+')
							{
								// For consistency, assume that since the previous item is an operand (even if it's
								// ')'), this is a post-op that applies to that operand.  For example, the following
								// are all treated the same for consistency (implicit concatention where the '.'
								// is omitted is rare anyway).
								// x++ y
								// x ++ y
								// x ++y
								// The following implicit concat is deliberately unsupported:
								//    "string" ++x
								// The ++ above is seen as applying to the string because it doesn't seem worth
								// the complexity to distinguish between expressions that can accept a post-op
								// and those that can't (operands other than variables can have a post-op;
								// e.g. (x:=y)++).
								++cp; // An additional increment to have loop skip over the operator's second symbol.
								this_infix_item.symbol = SYM_POST_INCREMENT;
							}
							else
								this_infix_item.symbol = SYM_ADD;
						}
						else if (cp1 == '+') // Pre-increment.
						{
							++cp; // An additional increment to have loop skip over the operator's second symbol.
							this_infix_item.symbol = SYM_PRE_INCREMENT;
						}
						else // Remove unary pluses from consideration since they do not change the calculation.
							--infix_count; // Counteract the loop's increment.
					}
					break;
				case '-':
					if (cp1 == '=')
					{
						++cp; // An additional increment to have loop skip over the operator's second symbol.
						this_infix_item.symbol = SYM_ASSIGN_SUBTRACT;
						break;
					}
					// Otherwise (since above didn't "break"):
					// Must allow consecutive unary minuses because otherwise, the following example
					// would not work correctly when y contains a negative value: var := 3 * -y
					if (infix_count && YIELDS_AN_OPERAND(infix[infix_count - 1].symbol))
					{
						if (cp1 == '-')
						{
							// See comments at SYM_POST_INCREMENT about this section.
							++cp; // An additional increment to have loop skip over the operator's second symbol.
							this_infix_item.symbol = SYM_POST_DECREMENT;
						}
						else
							this_infix_item.symbol = SYM_SUBTRACT;
					}
					else if (cp1 == '-') // Pre-decrement.
					{
						++cp; // An additional increment to have loop skip over the operator's second symbol.
						this_infix_item.symbol = SYM_PRE_DECREMENT;
					}
					else // Unary minus.
					{
						// Set default for cases where the processing below this line doesn't determine
						// it's a negative numeric literal:
						this_infix_item.symbol = SYM_NEGATIVE;
						// v1.0.40.06: The smallest signed 64-bit number (-0x8000000000000000) wasn't properly
						// supported in previous versions because its unary minus was being seen as an operator,
						// and thus the raw number was being passed as a positive to _atoi64() or _strtoi64(),
						// neither of which would recognize it as a valid value.  To correct this, a unary
						// minus followed by a raw numeric literal is now treated as a single literal number
						// rather than unary minus operator followed by a positive number.
						//
						// To be a valid "literal negative number", the character immediately following
						// the unary minus must not be:
						// 1) Whitespace (atoi() and such don't support it, nor is it at all conventional).
						// 2) An open-parenthesis such as the one in -(x).
						// 3) Another unary minus or operator such as --x (which is the pre-decrement operator).
						// To cover the above and possibly other unforeseen things, insist that the first
						// character be a digit (even a hex literal must start with 0).
						if ((cp1 >= '0' && cp1 <= '9') || cp1 == '.') // v1.0.46.01: Recognize dot too, to support numbers like -.5.
						{
							for (op_end = cp + 2; !strchr(EXPR_OPERAND_TERMINATORS, *op_end); ++op_end); // Find the end of this number (can be '\0').
							// 1.0.46.11: Due to obscurity, no changes have been made here to support scientific
							// notation followed by the power operator; e.g. -1.0e+1**5.
							if (!this_deref || op_end < this_deref->marker) // Detect numeric double derefs such as one created via "12%i% = value".
							{
								// Because the power operator takes precedence over unary minus, don't collapse
								// unary minus into a literal numeric literal if the number is immediately
								// followed by the power operator.  This is correct behavior even for
								// -0x8000000000000000 because -0x8000000000000000**2 would in fact be undefined
								// because ** is higher precedence than unary minus and +0x8000000000000000 is
								// beyond the signed 64-bit range.  SEE ALSO the comments higher above.
								// Use a temp variable because numeric_literal requires that op_end be set properly:
								char *pow_temp = omit_leading_whitespace(op_end);
								if (!(pow_temp[0] == '*' && pow_temp[1] == '*'))
									goto numeric_literal; // Goto is used for performance and also as a patch to minimize the chance of breaking other things via redesign.
								//else it's followed by pow.  Since pow is higher precedence than unary minus,
								// leave this unary minus as an operator so that it will take effect after the pow.
							}
							//else possible double deref, so leave this unary minus as an operator.
						}
					} // Unary minus.
					break;
				case ',':
					this_infix_item.symbol = SYM_COMMA; // Used to separate sub-statements and function parameters.
					break;
				case '/':
					if (cp1 == '=')
					{
						++cp; // An additional increment to have loop skip over the operator's second symbol.
						this_infix_item.symbol = SYM_ASSIGN_DIVIDE;
					}
					else if (cp1 == '/')
					{
						if (cp[2] == '=')
						{
							cp += 2; // An additional increment to have loop skip over the operator's 2nd & 3rd symbols.
							this_infix_item.symbol = SYM_ASSIGN_FLOORDIVIDE;
						}
						else
						{
							++cp; // An additional increment to have loop skip over the second '/' too.
							this_infix_item.symbol = SYM_FLOORDIVIDE;
						}
					}
					else
						this_infix_item.symbol = SYM_DIVIDE;
					break;
				case '*':
					if (cp1 == '=')
					{
						++cp; // An additional increment to have loop skip over the operator's second symbol.
						this_infix_item.symbol = SYM_ASSIGN_MULTIPLY;
					}
					else
					{
						if (cp1 == '*') // Python, Perl, and other languages also use ** for power.
						{
							++cp; // An additional increment to have loop skip over the second '*' too.
							this_infix_item.symbol = SYM_POWER;
						}
						else
						{
							// Differentiate between unary dereference (*) and the "multiply" operator:
							// See '-' above for more details:
							this_infix_item.symbol = (infix_count && YIELDS_AN_OPERAND(infix[infix_count - 1].symbol))
								? SYM_MULTIPLY : SYM_DEREF;
						}
					}
					break;
				case '!':
					if (cp1 == '=') // i.e. != is synonymous with <>, which is also already supported by legacy.
					{
						++cp; // An additional increment to have loop skip over the '=' too.
						this_infix_item.symbol = SYM_NOTEQUAL;
					}
					else
						// If what lies to its left is a CPARAN or OPERAND, SYM_CONCAT is not auto-inserted because:
						// 1) Allows ! and ~ to potentially be overloaded to become binary and unary operators in the future.
						// 2) Keeps the behavior consistent with unary minus, which could never auto-concat since it would
						//    always be seen as the binary subtract operator in such cases.
						// 3) Simplifies the code.
						this_infix_item.symbol = SYM_HIGHNOT; // High-precedence counterpart of the word "not".
					break;
				case '(':
					// The below should not hurt any future type-casting feature because the type-cast can be checked
					// for prior to checking the below.  For example, if what immediately follows the open-paren is
					// the string "int)", this symbol is not open-paren at all but instead the unary type-cast-to-int
					// operator.
					if (infix_count && YIELDS_AN_OPERAND(infix[infix_count - 1].symbol)) // If it's an operand, at this stage it can only be SYM_OPERAND or SYM_STRING.
					{
						if (infix_count > MAX_TOKENS - 2) // -2 to ensure room for this operator and the operand further below.
							goto abnormal_end;
						this_infix_item.symbol = SYM_CONCAT;
						++infix_count;
					}
					infix[infix_count].symbol = SYM_OPAREN; // MUST NOT REFER TO this_infix_item IN CASE ABOVE DID ++infix_count.
					break;
				case ')':
					this_infix_item.symbol = SYM_CPAREN;
					break;
				case '=':
					if (cp1 == '=')
					{
						++cp; // An additional increment to have loop skip over the other '=' too.
						this_infix_item.symbol = SYM_EQUALCASE;
					}
					else
						this_infix_item.symbol = SYM_EQUAL;
					break;
				case '>':
					switch (cp1)
					{
					case '=':
						++cp; // An additional increment to have loop skip over the '=' too.
						this_infix_item.symbol = SYM_GTOE;
						break;
					case '>':
						if (cp[2] == '=')
						{
							cp += 2; // An additional increment to have loop skip over the operator's 2nd & 3rd symbols.
							this_infix_item.symbol = SYM_ASSIGN_BITSHIFTRIGHT;
						}
						else
						{
							++cp; // An additional increment to have loop skip over the second '>' too.
							this_infix_item.symbol = SYM_BITSHIFTRIGHT;
						}
						break;
					default:
						this_infix_item.symbol = SYM_GT;
					}
					break;
				case '<':
					switch (cp1)
					{
					case '=':
						++cp; // An additional increment to have loop skip over the '=' too.
						this_infix_item.symbol = SYM_LTOE;
						break;
					case '>':
						++cp; // An additional increment to have loop skip over the '>' too.
						this_infix_item.symbol = SYM_NOTEQUAL;
						break;
					case '<':
						if (cp[2] == '=')
						{
							cp += 2; // An additional increment to have loop skip over the operator's 2nd & 3rd symbols.
							this_infix_item.symbol = SYM_ASSIGN_BITSHIFTLEFT;
						}
						else
						{
							++cp; // An additional increment to have loop skip over the second '<' too.
							this_infix_item.symbol = SYM_BITSHIFTLEFT;
						}
						break;
					default:
						this_infix_item.symbol = SYM_LT;
					}
					break;
				case '&':
					if (cp1 == '&')
					{
						++cp; // An additional increment to have loop skip over the second '&' too.
						this_infix_item.symbol = SYM_AND;
					}
					else if (cp1 == '=')
					{
						++cp; // An additional increment to have loop skip over the operator's second symbol.
						this_infix_item.symbol = SYM_ASSIGN_BITAND;
					}
					else
					{
						// Differentiate between unary "take the address of" and the "bitwise and" operator:
						// See '-' above for more details:
						this_infix_item.symbol = (infix_count && YIELDS_AN_OPERAND(infix[infix_count - 1].symbol))
							? SYM_BITAND : SYM_ADDRESS;
					}
					break;
				case '|':
					if (cp1 == '|')
					{
						++cp; // An additional increment to have loop skip over the second '|' too.
						this_infix_item.symbol = SYM_OR;
					}
					else if (cp1 == '=')
					{
						++cp; // An additional increment to have loop skip over the operator's second symbol.
						this_infix_item.symbol = SYM_ASSIGN_BITOR;
					}
					else
						this_infix_item.symbol = SYM_BITOR;
					break;
				case '^':
					if (cp1 == '=')
					{
						++cp; // An additional increment to have loop skip over the operator's second symbol.
						this_infix_item.symbol = SYM_ASSIGN_BITXOR;
					}
					else
						this_infix_item.symbol = SYM_BITXOR;
					break;
				case '~':
					// If what lies to its left is a CPARAN or OPERAND, SYM_CONCAT is not auto-inserted because:
					// 1) Allows ! and ~ to potentially be overloaded to become binary and unary operators in the future.
					// 2) Keeps the behavior consistent with unary minus, which could never auto-concat since it would
					//    always be seen as the binary subtract operator in such cases.
					// 3) Simplifies the code.
					this_infix_item.symbol = SYM_BITNOT;
					break;
				case '?':
					this_infix_item.symbol = SYM_IFF_THEN;
					break;
				case ':':
					if (cp1 == '=')
					{
						++cp; // An additional increment to have loop skip over the second '|' too.
						this_infix_item.symbol = SYM_ASSIGN;
					}
					else
						this_infix_item.symbol = SYM_IFF_ELSE;
					break;

				case '"': // QUOTED/LITERAL STRING.
					// Note that single and double-derefs are impossible inside string-literals
					// because the load-time deref parser would never detect anything inside
					// of quotes -- even non-escaped percent signs -- as derefs.
					if (infix_count && YIELDS_AN_OPERAND(infix[infix_count - 1].symbol)) // If it's an operand, at this stage it can only be SYM_OPERAND or SYM_STRING.
					{
						if (infix_count > MAX_TOKENS - 2) // -2 to ensure room for this operator and the operand further below.
							goto abnormal_end;
						this_infix_item.symbol = SYM_CONCAT;
						++infix_count;
					}
					// MUST NOT REFER TO this_infix_item IN CASE ABOVE DID ++infix_count:
					infix[infix_count].symbol = SYM_STRING; // Marked explicitly as string vs. SYM_OPERAND to prevent it from being seen as a number, e.g. if (var == "12.0") would be false if var contains "12" with no trailing ".0".
					infix[infix_count].marker = target; // Point it to its position in the buffer (populated below).
					// The following section is nearly identical to one in DefineFunc().
					// Find the end of this string literal, noting that a pair of double quotes is
					// a literal double quote inside the string:
					for (++cp;;) // Omit the starting-quote from consideration, and from the resulting/built string.
					{
						if (!*cp) // No matching end-quote. Probably impossible due to load-time validation.
							goto abnormal_end;
						if (*cp == '"') // And if it's not followed immediately by another, this is the end of it.
						{
							++cp;
							if (*cp != '"') // String terminator or some non-quote character.
								break;  // The previous char is the ending quote.
							//else a pair of quotes, which resolves to a single literal quote. So fall through
							// to the below, which will copy of quote character to the buffer. Then this pair
							// is skipped over and the loop continues until the real end-quote is found.
						}
						//else some character other than '\0' or '"'.
						*target++ = *cp++;
					}
					*target++ = '\0'; // Terminate it in the buffer.
					continue; // Continue vs. break to avoid the ++cp at the bottom. Above has already set cp to be the character after this literal string's close-quote.

				default: // NUMERIC-LITERAL, DOUBLE-DEREF, RELATIONAL OPERATOR SUCH AS "NOT", OR UNRECOGNIZED SYMBOL.
					if (*cp == '.') // This one must be done here rather than as a "case".  See comment below.
					{
						if (cp1 == '=')
						{
							++cp; // An additional increment to have loop skip over the operator's second symbol.
							this_infix_item.symbol = SYM_ASSIGN_CONCAT;
							break;
						}
						if (IS_SPACE_OR_TAB(cp1))
						{
							this_infix_item.symbol = SYM_CONCAT;
							break;
						}
						//else this is a '.' that isn't followed by a space, tab, or '='.  So it's probably
						// a number without a leading zero like .2, so continue on below to process it.
					}

					// Find the end of this operand or keyword, even if that end extended into the next deref.
					// StrChrAny() is not used because if *op_end is '\0', the strchr() below will find it too:
					for (op_end = cp + 1; !strchr(EXPR_OPERAND_TERMINATORS, *op_end); ++op_end);
					// Now op_end marks the end of this operand or keyword.  That end might be the zero terminator
					// or the next operator in the expression, or just a whitespace.
					if (this_deref && op_end >= this_deref->marker)
						goto double_deref; // This also serves to break out of the inner for(), equivalent to a break.
					// Otherwise, this operand is a normal raw numeric-literal or a word-operator (and/or/not).
					// The section below is very similar to the one used at load-time to recognize and/or/not,
					// so it should be maintained with that section.  UPDATE for v1.0.45: The load-time parser
					// now resolves "OR" to || and "AND" to && to improve runtime performance and reduce code size here.
					// However, "NOT" but still be parsed here at runtime because it's not quite the same as the "!"
					// operator (different precedence), and it seemed too much trouble to invent some special
					// operator symbol for load-time to insert as a placeholder/substitute (especially since that
					// symbol would appear in ListLines).
					if (op_end-cp == 3
						&& (cp[0] == 'n' || cp[0] == 'N')
						&& (  cp1 == 'o' ||   cp1 == 'O')
						&& (cp[2] == 't' || cp[2] == 'T')) // "NOT" was found.
					{
						this_infix_item.symbol = SYM_LOWNOT;
						cp = op_end; // Have the loop process whatever lies at op_end and beyond.
						continue; // Continue vs. break to avoid the ++cp at the bottom (though it might not matter in this case).
					}
numeric_literal:
					// Since above didn't "continue", this item is probably a raw numeric literal (either SYM_FLOAT
					// or SYM_INTEGER, to be differentiated later) because just about every other possibility has
					// been ruled out above.  For example, unrecognized symbols should be impossible at this stage
					// because load-time validation would have caught them.  And any kind of unquoted alphanumeric
					// characters (other than "NOT", which was detected above) wouldn't have reached this point
					// because load-time pre-parsing would have marked it as a deref/var, not raw/literal text.
					if (   toupper(op_end[-1]) == 'E' // v1.0.46.11: It looks like scientific notation...
						&& !(cp[0] == '0' && toupper(cp[1]) == 'X') // ...and it's not a hex number (this check avoids falsely detecting hex numbers that end in 'E' as exponents). This line fixed in v1.0.46.12.
						&& !(cp[0] == '-' && cp[1] == '0' && toupper(cp[2]) == 'X') // ...and it's not a negative hex number (this check avoids falsely detecting hex numbers that end in 'E' as exponents). This line added as a fix in v1.0.47.03.
						)
					{
						// Since op_end[-1] is the 'E' or an exponent, the only valid things for op_end[0] to be
						// are + or - (it can't be a digit because the loop above would never have stopped op_end
						// at a digit).  If it isn't + or -, it's some kind of syntax error, so doing the following
						// seems harmless in any case:
						do // Skip over the sign and its exponent; e.g. the "+1" in "1.0e+1".  There must be a sign in this particular sci-notation number or we would never have arrived here.
							++op_end;
						while (*op_end >= '0' && *op_end <= '9'); // Avoid isdigit() because it sometimes causes a debug assertion failure at: (unsigned)(c + 1) <= 256 (probably only in debug mode), and maybe only when bad data got in it due to some other bug.
					}
					if (infix_count && YIELDS_AN_OPERAND(infix[infix_count - 1].symbol)) // If it's an operand, at this stage it can only be SYM_OPERAND or SYM_STRING.
					{
						if (infix_count > MAX_TOKENS - 2) // -2 to ensure room for this operator and the operand further below.
							goto abnormal_end;
						this_infix_item.symbol = SYM_CONCAT;
						++infix_count;
					}
					// MUST NOT REFER TO this_infix_item IN CASE ABOVE DID ++infix_count:
					infix[infix_count].symbol = SYM_OPERAND;
					infix[infix_count].marker = target; // Point it to its position in the buffer (populated below).
					memcpy(target, cp, op_end - cp);
					target += op_end - cp;
					*target++ = '\0'; // Terminate it in the buffer.
					cp = op_end; // Have the loop process whatever lies at op_end and beyond.
					continue; // "Continue" to avoid the ++cp at the bottom.
				} // switch() for type of symbol/operand.
				++cp; // i.e. increment only if a "continue" wasn't encountered somewhere above. Although maintainability is reduced to do this here, it avoids dozens of ++cp in other places.
			} // for each token in this section of raw/literal text.
		} // End of processing of raw/literal text (such as operators) that lie to the left of this_deref.

		if (!this_deref) // All done because the above just processed all the raw/literal text (if any) that
			break;       // lay to the right of the last deref.

		// THE ABOVE HAS NOW PROCESSED ANY/ALL RAW/LITERAL TEXT THAT LIES TO THE LEFT OF this_deref.
		// SO NOW PROCESS THIS_DEREF ITSELF.
		if (infix_count > MAX_TOKENS - 1) // No room for the deref item below to be added.
			goto abnormal_end;
		DerefType &this_deref_ref = *this_deref; // Boosts performance slightly.
		if (this_deref_ref.is_function) // Above has ensured that at this stage, this_deref!=NULL.
		{
			if (infix_count && YIELDS_AN_OPERAND(infix[infix_count - 1].symbol)) // If it's an operand, at this stage it can only be SYM_OPERAND or SYM_STRING.
			{
				if (infix_count > MAX_TOKENS - 2) // -2 to ensure room for this operator and the operand further below.
					goto abnormal_end;
				infix[infix_count++].symbol = SYM_CONCAT;
			}
			infix[infix_count].symbol = SYM_FUNC;
			infix[infix_count].deref = deref;
		}
		else // this_deref is a variable.
		{
			if (*this_deref_ref.marker == g_DerefChar) // A double-deref because normal derefs don't start with '%'.
			{
				// Find the end of this operand, even if that end extended into the next deref.
				// StrChrAny() is not used because if *op_end is '\0', the strchr() below will find it too:
				for (op_end = this_deref_ref.marker + this_deref_ref.length; !strchr(EXPR_OPERAND_TERMINATORS, *op_end); ++op_end);
				goto double_deref;
			}
			else
			{
				if (infix_count && YIELDS_AN_OPERAND(infix[infix_count - 1].symbol)) // If it's an operand, at this stage it can only be SYM_OPERAND or SYM_STRING.
				{
					if (infix_count > MAX_TOKENS - 2) // -2 to ensure room for this operator and the operand further below.
						goto abnormal_end;
					infix[infix_count++].symbol = SYM_CONCAT;
				}
				if (this_deref_ref.var->Type() == VAR_NORMAL // VAR_ALIAS is taken into account (and resolved) by Type().
					&& (g_NoEnv || this_deref_ref.var->Length())) // v1.0.43.08: Added g_NoEnv.  Relies on short-circuit boolean order.
					// "!this_deref_ref.var->Get()" isn't checked here.  See comments in SYM_DYNAMIC evaluation.
				{
					// DllCall() and possibly others rely on this having been done to support changing the
					// value of a parameter (similar to by-ref).
					infix[infix_count].symbol = SYM_VAR; // Type() is always VAR_NORMAL as verified above. This is relied upon in several places such as built-in functions.
				}
				else // It's either a built-in variable (including clipboard) OR a possible environment variable.
				{
					infix[infix_count].symbol = SYM_DYNAMIC;
					infix[infix_count].buf = NULL; // SYM_DYNAMIC requires that buf be set to NULL for vars (since there are two different types of SYM_DYNAMIC).
				}
				infix[infix_count].var = this_deref_ref.var;
			}
		} // Handling of the var or function in this_deref.

		// Finally, jump over the dereference text. Note that in the case of an expression, there might not
		// be any percent signs within the text of the dereference, e.g. x + y, not %x% + %y% (unless they're
		// deliberately double-derefs).
		cp += this_deref_ref.length;
		// The outer loop will now do ++infix for us.

continue;     // To avoid falling into the label below. The label below is only reached by explicit goto.
double_deref: // Caller has set cp to be start and op_end to be the character after the last one of the double deref.
		if (infix_count && YIELDS_AN_OPERAND(infix[infix_count - 1].symbol)) // If it's an operand, at this stage it can only be SYM_OPERAND or SYM_STRING.
		{
			if (infix_count > MAX_TOKENS - 2) // -2 to ensure room for this operator and the operand further below.
				goto abnormal_end;
			infix[infix_count++].symbol = SYM_CONCAT;
		}

		infix[infix_count].symbol = SYM_DYNAMIC;
		infix[infix_count].buf = target; // Point it to its position in the buffer (populated below).
		memcpy(target, cp, op_end - cp); // "target" is incremented and string-terminated later below.

		// Set "deref" properly for the loop to resume processing at the item after this double deref.
		// Callers of double_deref have ensured that deref!=NULL and deref->marker!=NULL (because it
		// doesn't make sense to have a double-deref unless caller discovered the first deref that
		// belongs to this double deref, such as the "i" in Array%i%).
		for (deref_start = deref, ++deref; deref->marker && deref->marker < op_end; ++deref);
		derefs_in_this_double = (int)(deref - deref_start);
		--deref; // Compensate for the outer loop's ++deref.

		// There's insufficient room to shoehorn all the necessary data into the token (since circuit_token probably
		// can't be safely overloaded at this stage), so allocate a little bit of stack memory, just enough for the
		// number of derefs (variables) whose contents comprise the name of this double-deref variable (typically
		// there's only one; e.g. the "i" in Array%i%).
		deref_alloca = (DerefType *)_alloca((derefs_in_this_double + 1) * sizeof(DerefType)); // Provides one extra at the end as a terminator.
		memcpy(deref_alloca, deref_start, derefs_in_this_double * sizeof(DerefType));
		deref_alloca[derefs_in_this_double].marker = NULL; // Put a NULL in the last item, which terminates the array.
		for (deref_start = deref_alloca; deref_start->marker; ++deref_start)
			deref_start->marker = target + (deref_start->marker - cp); // Point each to its position in the *new* buf.
		infix[infix_count].var = (Var *)deref_alloca; // Postfix evaluation uses this to build the variable's name dynamically.

		target += op_end - cp; // Must be done only after the above, since it uses the old value of target.
		if (*op_end == '(') // i.e. dynamic function call 
		{
			if (infix_count > MAX_TOKENS - 2) // No room for the following symbol to be added (plus the ++infix done that will be done by the outer loop).
				goto abnormal_end;
			++infix_count;
			// As a result of a prior loop, deref_start = the null-marker deref which terminates the deref list. 
			deref_start->is_function = true;
			// param_count was set when the derefs were parsed. 
			deref_start->param_count = deref_alloca->param_count;
			infix[infix_count].symbol = SYM_FUNC;
			infix[infix_count].deref = deref_start;
			// postfix processing of SYM_DYNAMIC will update deref->func before SYM_FUNC is processed.
		}
		else
			deref_start->is_function = false;
		*target++ = '\0'; // Terminate the name, which looks something like "Array%i%".
		cp = op_end; // Must be done only after above is done using cp: Set things up for the next iteration.
		// The outer loop will now do ++infix for us.
	} // For each deref in this expression, and also for the final literal/raw text to the right of the last deref.

	// Terminate the array with a special item.  This allows infix-to-postfix conversion to do a faster
	// traversal of the infix array.
	if (infix_count > MAX_TOKENS - 1) // No room for the following symbol to be added.
		goto abnormal_end;
	infix[infix_count].symbol = SYM_INVALID;

	////////////////////////////
	// CONVERT INFIX TO POSTFIX.
	////////////////////////////
	#define STACK_PUSH(token_ptr) stack[stack_count++] = token_ptr
	#define STACK_POP stack[--stack_count]  // To be used as the r-value for an assignment.
	// SYM_BEGIN is the first item to go on the stack.  It's a flag to indicate that conversion to postfix has begun:
	ExprTokenType token_begin;
	token_begin.symbol = SYM_BEGIN;
	STACK_PUSH(&token_begin);

	SymbolType stack_symbol, infix_symbol, sym_prev;
	ExprTokenType *fwd_infix, *this_infix = infix;
	int functions_on_stack = 0;

	for (;;) // While SYM_BEGIN is still on the stack, continue iterating.
	{
		ExprTokenType *&this_postfix = postfix[postfix_count]; // Resolve early, especially for use by "goto". Reduces code size a bit, though it doesn't measurably help performance.
		infix_symbol = this_infix->symbol;                     //
		stack_symbol = stack[stack_count - 1]->symbol; // Frequently used, so resolve only once to help performance.

		// Put operands into the postfix array immediately, then move on to the next infix item:
		if (IS_OPERAND(infix_symbol)) // At this stage, operands consist of only SYM_OPERAND and SYM_STRING.
		{
			if (infix_symbol == SYM_DYNAMIC && SYM_DYNAMIC_IS_VAR_NORMAL_OR_CLIP(this_infix)) // Ordered for short-circuit performance.
			{
				// v1.0.46.01: If an environment variable is being used as an lvalue -- regardless
				// of whether that variable is blank in the environment -- treat it as a normal
				// variable instead.  This is because most people would want that, and also because
				// it's tranditional not to directly support assignments to environment variables
				// (only EnvSet can do that, mostly for code simplicity).  In addition, things like
				// EnvVar.="string" and EnvVar+=2 aren't supported due to obscurity/rarity (instead,
				// such expressions treat EnvVar as blank). In light of all this, convert environment
				// variables that are targets of ANY assignments into normal variables so that they
				// can be seen as a valid lvalues when the time comes to do the assignment.
				// IMPORTANT: VAR_CLIPBOARD is made into SYM_VAR here, but only for assignments.
				// This allows built-in functions and other places in the code to treat SYM_VAR
				// as though it's always VAR_NORMAL, which reduces code size and improves maintainability.
				sym_prev = this_infix[1].symbol; // Resolve to help macro's code size and performance.
				if (IS_ASSIGNMENT_OR_POST_OP(sym_prev) // Post-op must be checked for VAR_CLIPBOARD (by contrast, it seems unnecessary to check it for others; see comments below).
					|| stack_symbol == SYM_PRE_INCREMENT || stack_symbol == SYM_PRE_DECREMENT) // Stack *not* infix.
					this_infix->symbol = SYM_VAR; // Convert clipboard or environment variable into SYM_VAR.
				// POST-INC/DEC: It seems unnecessary to check for these except for VAR_CLIPBOARD because
				// those assignments (and indeed any assignment other than .= and :=) will have no effect
				// on a ON A SYM_DYNAMIC environment variable.  This is because by definition, such
				// variables have an empty Var::Contents(), and AutoHotkey v1 does not allow
				// math operations on blank variables.  Thus, the result of doing a math-assignment
				// operation on a blank lvalue is almost the same as doing it on an invalid lvalue.
				// The main difference is that with the exception of post-inc/dec, assignments
				// wouldn't produce an lvalue unless we explicitly check for them all above.
				// An lvalue should be produced so that the following features are consistent
				// even for variables whose names happen to match those of environment variables:
				// - Pass an assignment byref or takes its address; e.g. &(++x).
				// - Cascading assigments; e.g. (++var) += 4 (rare to be sure).
				// - Possibly other lvalue behaviors that rely on SYM_VAR being present.
				// Above logic might not be perfect because it doesn't check for parens such as (var):=x,
				// and possibly other obscure types of assignments.  However, it seems adequate given
				// the rarity of such things and also because env vars are being phased out (scripts can
				// use #NoEnv to avoid all such issues).
			}
			this_postfix = this_infix++;
			this_postfix->circuit_token = NULL; // Set default. It's only ever overridden after it's in the postfix array.
			++postfix_count;
			continue; // Doing a goto to a hypothetical "standard_postfix_circuit_token" (in lieu of these last 3 lines) reduced performance and didn't help code size.
		}

		// Since above didn't "continue", the current infix symbol is not an operand, but an operator or other symbol.

		switch(infix_symbol)
		{
		case SYM_CPAREN: // Listed first for performance. It occurs frequently while emptying the stack to search for the matching open-parenthesis.
			if (stack_symbol == SYM_OPAREN) // See comments near the bottom of this case.  The first open-paren on the stack must be the one that goes with this close-paren.
			{
				--stack_count; // Remove this open-paren from the stack, since it is now complete.
				++this_infix;  // Since this pair of parentheses is done, move on to the next token in the infix expression.
				// There should be no danger of stack underflow in the following because SYM_BEGIN always
				// exists at the bottom of the stack:
				if (stack[stack_count - 1]->symbol == SYM_FUNC) // i.e. topmost item on stack is SYM_FUNC.
				{
					--functions_on_stack;
					goto standard_pop_into_postfix; // Within the postfix list, a function-call should always immediately follow its params.
				}
			}
			else if (stack_symbol == SYM_BEGIN) // Paren is closed without having been opened (currently impossible due to load-time balancing, but kept for completeness).
				goto abnormal_end; 
			else // This stack item is an operator.
			{
				goto standard_pop_into_postfix;
				// By not incrementing i, the loop will continue to encounter SYM_CPAREN and thus
				// continue to pop things off the stack until the corresponding OPAREN is reached.
			}
			break;

		case SYM_FUNC:
			++functions_on_stack; // This technique performs well but prevents multi-statements from being nested inside function calls (seems too obscure to worry about); e.g. fn((x:=5, y+=3), 2)
			STACK_PUSH(this_infix++);
			// NOW FALL INTO THE OPEN-PAREN BELOW because load-time validation has ensured that each SYM_FUNC
			// is followed by a '('.
// ABOVE CASE FALLS INTO BELOW.
		case SYM_OPAREN:
			// Open-parentheses always go on the stack to await their matching close-parentheses.
			STACK_PUSH(this_infix++);
			break;

		case SYM_IFF_ELSE: // i.e. this infix symbol is ':'.
			if (stack_symbol == SYM_BEGIN) // ELSE with no matching IF/THEN (load-time currently doesn't validate/detect this).
				goto abnormal_end;  // Below relies on the above check having been done, to avoid underflow.
			// Otherwise:
			this_postfix = STACK_POP; // There should be no danger of stack underflow in the following because SYM_BEGIN always exists at the bottom of the stack.
			if (stack_symbol == SYM_IFF_THEN) // See comments near the bottom of this case. The first found "THEN" on the stack must be the one that goes with this "ELSE".
			{
				this_postfix->circuit_token = this_infix; // Point this "THEN" to its "ELSE" for use by short-circuit. This simplifies short-circuit by means such as avoiding the need to take notice of nested IFF's when discarding a branch (a different stage points the IFF's condition to its "THEN").
				STACK_PUSH(this_infix++); // Push the ELSE onto the stack so that its operands will go into the postfix array before it.
				// Above also does ++i since this ELSE found its matching IF/THEN, so it's time to move on to the next token in the infix expression.
			}
			else // This stack item is an operator INCLUDE some other THEN's ELSE (all such ELSE's should be purged from the stack so that 1 ? 1 ? 2 : 3 : 4 creates postfix 112?3:?4: not something like 112?3?4::.
			{
				this_postfix->circuit_token = NULL; // Set default. It's only ever overridden after it's in the postfix array.
				// By not incrementing i, the loop will continue to encounter SYM_IFF_ELSE and thus
				// continue to pop things off the stack until the corresponding SYM_IFF_THEN is reached.
			}
			++postfix_count;
			break;

		case SYM_INVALID:
			if (stack_symbol == SYM_BEGIN) // Stack is basically empty, so stop the loop.
			{
				--stack_count; // Remove SYM_BEGIN from the stack, leaving the stack empty for use in postfix eval.
				goto end_of_infix_to_postfix; // Both infix and stack have been fully processed, so move on to the postfix evaluation phase.
			}
			else if (stack_symbol == SYM_OPAREN) // Open paren is never closed (currently impossible due to load-time balancing, but kept for completeness).
				goto abnormal_end;
			else // Pop item off the stack, AND CONTINUE ITERATING, which will hit this line until stack is empty.
				goto standard_pop_into_postfix;
			// ALL PATHS ABOVE must continue or goto.

		default: // This infix symbol is an operator, so act according to its precedence.
			// If the symbol waiting on the stack has a lower precedence than the current symbol, push the
			// current symbol onto the stack so that it will be processed sooner than the waiting one.
			// Otherwise, pop waiting items off the stack (by means of i not being incremented) until their
			// precedence falls below the current item's precedence, or the stack is emptied.
			// Note: BEGIN and OPAREN are the lowest precedence items ever to appear on the stack (CPAREN
			// never goes on the stack, so can't be encountered there).
			if (   sPrecedence[stack_symbol] < sPrecedence[infix_symbol] + (sPrecedence[infix_symbol] % 2) // Performance: An sPrecedence2[] array could be made in lieu of the extra add+indexing+modulo, but it benched only 0.3% faster, so the extra code size it caused didn't seem worth it.
				|| IS_ASSIGNMENT_EXCEPT_POST_AND_PRE(infix_symbol) && stack_symbol != SYM_DEREF // See note 1 below. Ordered for short-circuit performance.
				|| stack_symbol == SYM_POWER && (infix_symbol >= SYM_NEGATIVE && infix_symbol <= SYM_DEREF // See note 2 below. Check lower bound first for short-circuit performance.
					|| infix_symbol == SYM_LOWNOT)   )
			{
				// NOTE 1: v1.0.46: The IS_ASSIGNMENT_EXCEPT_POST_AND_PRE line above was added in conjunction with
				// the new assignment operators (e.g. := and +=). Here's what it does: Normally, the assignment
				// operators have the lowest precedence of all (except for commas) because things that lie
				// to the *right* of them in the infix expression should be evaluated first to be stored
				// as the assignment's result.  However, if what lies to the *left* of the assignment
				// operator isn't a valid lvalue/variable (and not even a unary like -x can produce an lvalue
				// because they're not supposed to alter the contents of the variable), obeying the normal
				// precedence rules would be produce a syntax error due to "assigning to non-lvalue".
				// So instead, apply any pending operator on the stack (which lies to the left of the lvalue
				// in the infix expression) *after* the assignment by leaving it on the stack.  For example,
				// C++ and probably other langauges (but not the older ANSI C) evaluate "true ? x:=1 : y:=1"
				// as a pair of assignments rather than as who-knows-what (probably a syntax error if you
				// strictly followed precedence).  Similarly, C++ evaluates "true ? var1 : var2 := 3" not as
				// "(true ? var1 : var2) := 3" (like ANSI C) but as "true ? var1 : (var2 := 3)".  Other examples:
				// -> not var:=5 ; It's evaluated contrary to precedence as: not (var:=5) [PHP does this too,
				//    and probably others]
				// -> 5 + var+=5 ; It's evaluated contrary to precedence as: 5 + (var+=5) [not sure if other
				//    languages do ones like this]
				// -> ++i := 5 ; Silly since increment has no lasting effect; so assign the 5 then do the pre-inc.
				// -> ++i /= 5 ; Valid, but maybe too obscure and inconsistent to treat it differently than
				//    the others (especially since not many people will remember that unlike i++, ++i produces
				//    an lvalue); so divide by 5 then do the increment.
				// -> i++ := 5 (and i++ /= 5) ; Postfix operator can't produce an lvalue, so do the assignment
				//    first and then the postfix op.
				// SYM_DEREF is the only exception to the above because there's a slight chance that
				// *Var:=X (evaluated strictly according to precedence as (*Var):=X) will be used for someday.
				// Also, SYM_FUNC seems unaffected by any of this due to its enclosing parentheses (i.e. even
				// if a function-call can someday generate an lvalue [SYM_VAR], the current rules probably
				// already support it.
				// Performance: Adding the above behavior reduced the expressions benchmark by only 0.6%; so
				// it seems worth it.
				//
				// NOTE 2: The SYM_POWER line above is a workaround to allow 2**-2 (and others in v1.0.45) to be
				// evaluated as 2**(-2) rather than being seen as an error.  However, as of v1.0.46, consecutive
				// unary operators are supported via the right-to-left evaluation flag above (formerly, consecutive
				// unaries produced a failure [blank value]).  For example:
				// !!x  ; Useful to convert a blank value into a zero for use with unitialized variables.
				// not not x  ; Same as above.
				// Other examples: !-x, -!x, !&x, -&Var, ~&Var
				// And these deref ones (which worked even before v1.0.46 by different means: giving
				// '*' a higher precedence than the other unaries): !*Var, -*Var and ~*Var
				// !x  ; Supported even if X contains a negative number, since x is recognized as an isolated operand and not something containing unary minus.
				//
				// To facilitate short-circuit boolean evaluation, right before an AND/OR/IFF is pushed onto the
				// stack, point the end of it's left branch to it.  Note that the following postfix token
				// can itself be of type AND/OR/IFF, a simple example of which is "if (true and true and true)",
				// in which the first and's parent (in an imaginary tree) is the second "and".
				// But how is it certain that this is the final operator or operand of and AND/OR/IFF's left branch?
				// Here is the explanation:
				// Everything higher precedence than the AND/OR/IFF came off the stack right before it, resulting in
				// what must be a balanced/complete sub-postfix-expression in and of itself (unless the expression
				// has a syntax error, which is caught in various places).  Because it's complete, during the
				// postfix evaluation phase, that sub-expression will result in a new operand for the stack,
				// which must then be the left side of the AND/OR/IFF because the right side immediately follows it
				// within the postfix array, which in turn is immediately followed its operator (namely AND/OR/IFF).
				// Also, the final result of an IFF's condition-branch must point to the IFF/THEN symbol itself
				// because that's the means by which the condition is merely "checked" rather than becoming an
				// operand itself.
				if (infix_symbol <= SYM_AND && infix_symbol >= SYM_IFF_THEN && postfix_count) // Check upper bound first for short-circuit performance.
					postfix[postfix_count - 1]->circuit_token = this_infix; // In the case of IFF, this points the final result of the IFF's condition to its SYM_IFF_THEN (a different stage points the THEN to its ELSE).
				if (infix_symbol != SYM_COMMA)
					STACK_PUSH(this_infix);
				else // infix_symbol == SYM_COMMA, but which type of comma (function vs. statement-separator).
				{
					// KNOWN LIMITATION: Although the functions_on_stack method is simple and efficient, it isn't
					// capable of detecting commas that separate statements inside a function call such as:
					//    fn(x, (y:=2, fn2()))
					// Thus, such attempts will cause the expression as a whole to fail and evaluate to ""
					// (though individual parts of the expression may execute before it fails).
					// C++ and possibly other C-like languages seem to allow such expressions as shown by the
					// following simple example: MsgBox((1, 2)); // In which MsgBox sees (1, 2) as a single arg.
					// Perhaps this could be solved someday by checking/tracking whether there is a non-function
					// open-paren on the stack above/prior to the first function-call-open-paren on the stack.
					// That rule seems flexible enough to work even for things like f1((f2(), X)).  Perhaps a
					// simple stack traversal could be done to find the first OPAREN.  If it's a function's OPAREN,
					// this is a function-comma.  Otherwise, this comma is a statement-separator nested inside a
					// function call.  But the performance impact of that doesn't seem worth it given rarity of use.
					if (!functions_on_stack) // This comma separates statements rather than function parameters.
					{
						STACK_PUSH(this_infix);
						// v1.0.46.01: Treat ", var = expr" as though the "=" is ":=", even if there's a ternary
						// on the right side (for consistency and since such a ternary would be stand-alone,
						// which is a rare use for ternary).  Also cascade to the right to treat things like
						// x=y=z as assignments because its intuitiveness seems to outweigh other considerations.
						// In a future version, these transformations could be done at loadtime to improve runtime
						// performance; but currently that seems more complex than it's worth (and loadtime
						// performance and code size shouldn't be entirely ignored).
						for (fwd_infix = this_infix + 1;; fwd_infix += 2)
						{
							// The following is checked first to simplify things and avoid any chance of reading
							// beyond the last item in the array. This relies on the fact that a SYM_INVALID token
							// exists at the end of the array as a terminator.
							if (fwd_infix->symbol == SYM_INVALID || fwd_infix[1].symbol != SYM_EQUAL) // Relies on short-circuit boolean order.
								break; // No further checking needed because there's no qualified equal-sign.
							// Otherwise, check what lies to the left of the equal-sign.
							if (fwd_infix->symbol == SYM_VAR)
							{
								fwd_infix[1].symbol = SYM_ASSIGN;
								continue; // Cascade to the right until the last qualified '=' operator is found.
							}
							// Otherwise, it's not a pure/normal variable.  But check if it's an environment var.
							if (fwd_infix->symbol != SYM_DYNAMIC || !SYM_DYNAMIC_IS_VAR_NORMAL_OR_CLIP(fwd_infix))
								break; // It qualifies as neither SYM_DYNAMIC nor SYM_VAR.
							// Otherwise, this is an environment variable being assigned something, so treat
							// it as a normal variable rather than an environment variable. This is because
							// by tradition (and due to the fact that not many people would want it),
							// direct assignment to environment variables isn't supported by anything other
							// than EnvSet.
							fwd_infix->symbol = SYM_VAR; // Convert dynamic to a normal variable, see above.
							fwd_infix[1].symbol = SYM_ASSIGN;
							// And now cascade to the right until the last qualified '=' operator is found.
						}
					}
					//else it's a function comma, so don't put it in stack because function commas aren't
					// needed and they would probably prevent proper evaluation.  Only statement-separator
					// commas need to go onto the stack (see SYM_COMMA further below for comments).
				}
				++this_infix; // Regardless of the outcome above, move rightward to the next infix item.
			}
			else // Stack item's precedence >= infix's (if equal, left-to-right evaluation order is in effect).
				goto standard_pop_into_postfix;
		} // switch(infix_symbol)

		continue; // Avoid falling into the label below except via explicit jump.  Performance: Doing it this way rather than replacing break with continue everywhere above generates slightly smaller and slightly faster code.
standard_pop_into_postfix: // Use of a goto slightly reduces code size.
		this_postfix = STACK_POP;
		this_postfix->circuit_token = NULL; // Set default. It's only ever overridden after it's in the postfix array.
		++postfix_count;
	} // End of loop that builds postfix array from the infix array.
end_of_infix_to_postfix:

	///////////////////////////////////////////////////
	// EVALUATE POSTFIX EXPRESSION (constructed above).
	///////////////////////////////////////////////////
	int i, j, s, actual_param_count, delta;
	SymbolType right_is_number, left_is_number, result_symbol;
	double right_double, left_double;
	__int64 right_int64, left_int64;
	char *right_string, *left_string;
	char *right_contents, *left_contents;
	size_t right_length, left_length;
	char left_buf[MAX_FORMATTED_NUMBER_LENGTH + 1];  // BIF_OnMessage and SYM_DYNAMIC rely on this one being large enough to hold MAX_VAR_NAME_LENGTH.
	char right_buf[MAX_FORMATTED_NUMBER_LENGTH + 1]; // Only needed for holding numbers
	char *result; // "result" is used for return values and also the final result.
	VarSizeType result_length;
	size_t result_size, alloca_usage = 0; // v1.0.45: Track amount of alloca mem to avoid stress on stack from extreme expressions (mostly theoretical).
	BOOL done, done_and_have_an_output_var, make_result_persistent, left_branch_is_true
		, left_was_negative, is_pre_op; // BOOL vs. bool benchmarks slightly faster, and is slightly smaller in code size (or maybe it's cp1's int vs. char that shrunk it).
	ExprTokenType *circuit_token;
	Var *sym_assign_var, *temp_var;
	VarBkp *var_backup = NULL;  // If needed, it will hold an array of VarBkp objects. v1.0.40.07: Initialized to NULL to facilitate an approach that's more maintainable.
	int var_backup_count; // The number of items in the above array (when it's non-NULL).

	// v1.0.44.06: EXPR_SMALL_MEM_LIMIT is the means by which _alloca() is used to boost performance a
	// little by avoiding the overhead of malloc+free for small strings.  The limit should be something
	// small enough that the chance that even 10 of them would cause stack overflow is vanishingly small
	// (the program is currently compiled to allow stack to expand anyway).  Even in a worst-case
	// scenario where an expression is composed entirely of functions and they all need to use this
	// limit of stack space, there's a practical limit on how many functions you can call in an
	// expression due to MAX_TOKENS (probably around MAX_TOKENS / 3).
	#define EXPR_SMALL_MEM_LIMIT 4097
	#define EXPR_ALLOCA_LIMIT 40000  // v1.0.45: Just as an extra precaution against stack stress in extreme/theoretical cases.

	// For each item in the postfix array: if it's an operand, push it onto stack; if it's an operator or
	// function call, evaluate it and push its result onto the stack.
	for (i = 0; i < postfix_count; ++i) // Performance: Using a handle to traverse the postfix array rather than array indexing unexpectedly benchmarked 3% slower (perhaps not statistically significant due to being caused by CPU cache hits or compiler's use of registers).  Because of that, there's not enough reason to switch to that method -- though it does generate smaller code (perhaps a savings of 200 bytes).
	{
		ExprTokenType &this_token = *postfix[i];  // For performance and convenience.

		// At this stage, operands in the postfix array should be SYM_OPERAND, SYM_STRING, or SYM_DYNAMIC.
		// But all are checked since that operation is just as fast:
		if (IS_OPERAND(this_token.symbol)) // If it's an operand, just push it onto stack for use by an operator in a future iteration.
		{
			if (this_token.symbol == SYM_DYNAMIC) // CONVERTED HERE/EARLY TO SOMETHING *OTHER* THAN SYM_DYNAMIC so that no later stages need any handling for them as operands. SYM_DYNAMIC is quite similar to SYM_FUNC/BIF in this respect.
			{
				if (!SYM_DYNAMIC_IS_DOUBLE_DEREF(this_token)) // It's a built-in variable or potential environment variable.
				{
					result_size = this_token.var->Get() + 1; // Get() is used even for environment vars because it has a cache that improves their performance.
					if (result_size == 1)
					{
						if (this_token.var->Type() == VAR_NORMAL) // It's an empty variable, so treated as a non-environment (normal) var.
						{
							// The following is done here rather than during infix creation/tokenizing because
							// 1) It's more correct because it's conceivable that some part of the expression
							//    that has already been evaluated before this_token has newly made an environment
							//    variable blank or non-blank, which should be detected here (i.e. only at the
							//    last possible moment).  For example, a function might have the side-effect of
							//    altering an environment variable.
							// 2) It performs better because Get()'s environment variable cache is most effective
							//    when each size-Get() is followed immediately by a contents-Get() for the same
							//    variable.
							// Must make empty variables that aren't environment variables into SYM_VAR so that
							// they can be passed by reference into functions, their address can be taken with
							// the '&' operator, and so that they can be the lvalue for an assignment.
							// Environment variables aren't supported for any of that because it would be silly
							// in most cases, and would probably complicate the code far more than its worth.
							this_token.symbol = SYM_VAR; // The fact that a SYM_VAR operand is always VAR_NORMAL (with one limited exception) is relied upon in several places such as built-in functions.
						}
						else // It's a built-in variable that's blank.
						{
							this_token.marker = "";
							this_token.symbol = SYM_STRING;
						}
						goto push_this_token;
					}
					// Otherwise, it's not an empty string.  But there's a slight chance it could be a normal
					// variable rather than a built-in or environment variable.  This happens when another
					// part of this same expression (such as a UDF via ByRef) has put contents into this
					// variable since the time this item was made SYM_DYNAMIC.  In other words, we wouldn't
					// have made it SYM_DYNAMIC in the first place if we'd known that was going to happen.
					if (this_token.var->Type() == VAR_NORMAL && this_token.var->Length()) // v1.0.46.07: It's not a built-in or environment variable.
					{
						this_token.symbol = SYM_VAR; // The fact that a SYM_VAR operand is always VAR_NORMAL (with one limited exception) is relied upon in several places such as built-in functions.
						goto push_this_token;
					}
					// Otherwise, it's an environment variable or built-in variable. Need some memory to store it.
					// The following section is similar to that in the make_result_persistent section further
					// below.  So maintain them together and see it for more comments.
					// Must cast to int to avoid loss of negative values:
					if (result_size <= (int)(aDerefBufSize - (target - aDerefBuf))) // There is room at the end of our deref buf, so use it.
					{
						// Point result to its new, more persistent location:
						result = target;
						target += result_size; // Point it to the location where the next string would be written.
					}
					else if (result_size < EXPR_SMALL_MEM_LIMIT && alloca_usage < EXPR_ALLOCA_LIMIT) // See comments at EXPR_SMALL_MEM_LIMIT.
					{
						result = (char *)_alloca(result_size);
						alloca_usage += result_size; // This might put alloca_usage over the limit by as much as EXPR_SMALL_MEM_LIMIT, but that is fine because it's more of a guideline than a limit.
					}
					else // Need to create some new persistent memory for our temporary use.
					{
						if (mem_count == MAX_EXPR_MEM_ITEMS // No more slots left (should be nearly impossible).
							|| !(mem[mem_count] = (char *)malloc(result_size)))
						{
							LineError(ERR_OUTOFMEM ERR_ABORT, FAIL, this_token.var->mName);
							goto abort;
						}
						// Point result to its new, more persistent location:
						result = mem[mem_count];
						++mem_count; // Must be done last.
					}
					this_token.var->Get(result); // MUST USE "result" TO AVOID OVERWRITING MARKER/VAR UNION.
					this_token.marker = result;  // Must be done last because marker and var overlap in union.
					this_token.symbol = SYM_OPERAND; // Generic operand so that it can later be interpreted as a number (if it's numeric).
				}
				else // Double-deref such as Array%i%.
				{
					// Start off by looking for the first deref.
					deref = (DerefType *)this_token.var; // MUST BE DONE PRIOR TO OVERWRITING MARKER/UNION BELOW.
					cp = this_token.buf; // Start at the begining of this arg's text.
					int var_name_length = 0;

					this_token.marker = "";         // Set default in case of early goto.  Must be done after above.
					this_token.symbol = SYM_STRING; //

					// Lexikos: Changed "goto push_this_token" to "goto double_deref_fail" so expression evaluation can be aborted if a dynamic function reference fails to resolve.

					// Loadtime validation has ensured that none of these derefs are function-calls
					// (i.e. deref->is_function is alway false).  Loadtime logic seems incapable of
					// producing function-derefs inside something that would later be interpreted
					// as a double-deref.
					for (; deref->marker; ++deref)  // A deref with a NULL marker terminates the list. "deref" was initialized higher above.
					{
						// FOR EACH DEREF IN AN ARG (if we're here, there's at least one):
						// Copy the chars that occur prior to deref->marker into the buffer:
						for (; cp < deref->marker && var_name_length < MAX_VAR_NAME_LENGTH; left_buf[var_name_length++] = *cp++);
						if (var_name_length >= MAX_VAR_NAME_LENGTH && cp < deref->marker) // The variable name would be too long!
							goto double_deref_fail; // For simplicity and in keeping with the tradition that expressions generally don't display runtime errors, just treat it as a blank (or abort if this is a dynamic function call.)
						// Now copy the contents of the dereferenced var.  For all cases, aBuf has already
						// been verified to be large enough, assuming the value hasn't changed between the
						// time we were called and the time the caller calculated the space needed.
						if (deref->var->Get() > (VarSizeType)(MAX_VAR_NAME_LENGTH - var_name_length)) // The variable name would be too long!
							goto double_deref_fail; // For simplicity and in keeping with the tradition that expressions generally don't display runtime errors, just treat it as a blank (or abort if this is a dynamic function call.)
						var_name_length += deref->var->Get(left_buf + var_name_length);
						// Finally, jump over the dereference text. Note that in the case of an expression, there might not
						// be any percent signs within the text of the dereference, e.g. x + y, not %x% + %y%.
						cp += deref->length;
					}

					// Copy any chars that occur after the final deref into the buffer:
					for (; *cp && var_name_length < MAX_VAR_NAME_LENGTH; left_buf[var_name_length++] = *cp++);
					if (var_name_length >= MAX_VAR_NAME_LENGTH && *cp // The variable name would be too long!
						|| !var_name_length) // It resolves to an empty string (e.g. a simple dynamic var like %Var% where Var is blank).
						goto double_deref_fail; // For simplicity and in keeping with the tradition that expressions generally don't display runtime errors, just treat it as a blank (or abort if this is a dynamic function call.)

					// Terminate the buffer, even if nothing was written into it:
					left_buf[var_name_length] = '\0';

					// As a result of a prior loop, deref = the null-marker deref which terminates the deref list.
					// is_function is set by the infix processing code.
					if (deref->is_function)
					{
						// Traditionally, expressions don't display any runtime errors.  So if the function is being
						// called incorrectly by the script, the expression is aborted like it would be for other
						// syntax errors.
						if (   !(deref->func = g_script.FindFunc(left_buf, var_name_length)) // Below relies on short-circuit boolean order, with this line being executed first.
							// Lexikos: Disabled these checks so dynamic function calls may be less strict, in combination with some other modifications.
							/*|| deref->param_count > deref->func->mParamCount    // param_count was set by the
							|| deref->param_count < deref->func->mMinParams*/   ) // infix processing code.
							goto abnormal_end;
						// The SYM_FUNC following this SYM_DYNAMIC uses the same deref as above.
						continue;
					}

					// In v1.0.31, FindOrAddVar() vs. FindVar() is called below to support the passing of non-existent
					// array elements ByRef, e.g. Var:=MyFunc(Array%i%) where the MyFunc function's parameter is
					// defined as ByRef, would effectively create the new element Array%i% if it doesn't already exist.
					// Since at this stage we don't know whether this particular double deref is to be sent as a param
					// to a function, or whether it will be byref, this is done unconditionally for all double derefs
					// since it seems relatively harmless to create a blank variable in something like var := Array%i%
					// (though it will produce a runtime error if the double resolves to an illegal variable name such
					// as one containing spaces).
					// The use of ALWAYS_PREFER_LOCAL below improves flexibility of assume-global functions
					// by allowing this command to resolve to a local first if such a local exists:
					if (   !(temp_var = g_script.FindOrAddVar(left_buf, var_name_length, ALWAYS_PREFER_LOCAL))   )
					{
						// Above already displayed the error.  As of v1.0.31, this type of error is displayed and
						// causes the current thread to terminate, which seems more useful than the old behavior
						// that tolerated anything in expressions.
						goto abort;
					}
					// Otherwise, var was found or created.
					if (temp_var->Type() != VAR_NORMAL)
					{
						// Non-normal variables such as Clipboard and A_ScriptFullPath are not allowed to be
						// generated from a double-deref such as A_Script%VarContainingFullPath% because:
						// Update: The only good reason now is code simplicity.  Reason #1 below could probably
						// be solved via SYM_DYNAMIC.
						// 1) Anything that needed their contents would have to find memory in which to store
						//    the result of Var::Get(), which would complicate the code since such handling would have
						//    to be added.
						// 2) It doesn't appear to have much use, not even for passing them as a ByRef parameter to
						//    a function (since they're read-only [except Clipboard, but temporary memory would be
						//    needed somewhere if the clipboard contains files that need to be expanded to text] and
						//    essentially global by their very nature), and the value of catching unintended usages
						//    seems more important than any flexibilty that might add.
						goto push_this_token; // For simplicity and in keeping with the tradition that expressions generally don't display runtime errors, just treat it as a blank.
					}
					// Otherwise:
					// Even if it's an environment variable, it gets added as SYM_VAR.  However, unlike other
					// aspects of the program, double-derefs that resolve to environment variables will be seen
					// as always-blank due to the use of Var::Contents() vs. Var::Get() in various places.
					// This seems okay due to the extreme rarity of anyone intentionally wanting a double
					// reference such as Array%i% to resolve to the name of an environment variable.
					this_token.symbol = SYM_VAR; // The fact that a SYM_VAR operand is always VAR_NORMAL (with one limited exception) is relied upon in several places such as built-in functions.
					this_token.var = temp_var;
				} // Double-deref.
			} // if (this_token.symbol == SYM_DYNAMIC)
			goto push_this_token;
		} // if (IS_OPERAND(this_token.symbol))

		if (this_token.symbol == SYM_FUNC) // A call to a function (either built-in or defined by the script).
		{
			Func &func = *this_token.deref->func; // For performance.
			actual_param_count = this_token.deref->param_count; // For performance.
			if (actual_param_count > stack_count) // Prevent stack underflow (probably impossible if actual_param_count is accurate).
				goto abnormal_end;
			if (func.mIsBuiltIn)
			{
				// Adjust the stack early to simplify.  Above already confirmed that this won't underflow.
				// Pop the actual number of params involved in this function-call off the stack.  Load-time
				// validation has ensured that this number is always less than or equal to the number of
				// parameters formally defined by the function.  Therefore, there should never be any leftover
				// function-params on the stack after this is done:
				stack_count -= actual_param_count; // The function called below will see this portion of the stack as an array of its parameters.
				this_token.symbol = SYM_INTEGER; // Set default return type so that functions don't have to do it if they return INTs.
				this_token.marker = func.mName;  // Inform function of which built-in function called it (allows code sharing/reduction). Can't use circuit_token because it's value is still needed later below.
				this_token.buf = left_buf;       // mBIF() can use this to store a string result, and for other purposes.

				// BACK UP THE CIRCUIT TOKEN (it's saved because it can be non-NULL at this point (verified
				// through code review).
				circuit_token = this_token.circuit_token;
				this_token.circuit_token = NULL; // Init to detect whether the called function allocates it (i.e. we're overloading it with a new purpose).
				// RESIST TEMPTATIONS TO OPTIMIZE CIRCUIT_TOKEN by passing output_var as circuit_token
				// when done==true (i.e. the built-in function could then assign directly to output_var).
				// It doesn't help performance at all except for a mere 10% or less in certain fairly rare cases.
				// More importantly, it hurts maintainability because it makes RegExReplace() more complicated
				// than it already is, and worse: each BIF would have to check that output_var doesn't overlap
				// with its input/source strings because if it does, the function must not initialize a default
				// in output_var before starting (and avoiding this would further complicate the code).
				// Here is the crux of the abandoned approach: Any function that wishes to pass memory back to
				// us via circuit_token: When circuit_token!=NULL, that function MUST INSTEAD: 1) Turn that
				// memory over to output_var via AcceptNewMem(); 2) Set circuit_token to NULL to indicate to
				// us that it is a user of circuit_token.

				// CALL THE FUNCTION:
				func.mBIF(this_token, stack + stack_count, actual_param_count);

				// RESTORE THE CIRCUIT TOKEN (after handling what came back inside it):
				#define EXPR_IS_DONE (!stack_count && i == postfix_count-1) // True if we've used up the last of the operators & operands.
				done = EXPR_IS_DONE;
				done_and_have_an_output_var = done && output_var; // i.e. this is ACT_ASSIGNEXPR and we've now produced the final result.
				make_result_persistent = true; // Set default.
				if (this_token.circuit_token) // The called function allocated some memory here (to facilitate returning long strings) and turned it over to us.
				{
					// In most cases, the string stored in circuit_token is the same address as this_token.marker
					// (i.e. what is named "result" further below), because that's what the built-in functions
					// are normally using the memory for.
					if ((char *)this_token.circuit_token == this_token.marker) // circuit_token is checked in case caller alloc'd mem but didn't use it as its actual result.
					{
						// v1.0.45: If possible, take a shortcut for performance.  Doing it this way saves at least
						// two memcpy's (one into deref buffer and then another back into the output_var by
						// ACT_ASSIGNEXPR itself).  In some cases is also saves from having to expand the deref
						// buffer as well as the output_var (since it's current memory might be too small to hold
						// the new memory block). Thus we give it a new block directly to avoid all of that.
						// This should be a big boost to performance when long strings are involved.
						// So now, turn over responsibility for this memory to the variable. The called function
						// is responsible for having stored the length of what's in the memory as an overload of
						// this_token.buf, but only when that memory is the result (currently might always be true).
						if (done_and_have_an_output_var)
						{
							// AcceptNewMem() will shrink the memory for us, via _expand(), if there's a lot of
							// extra/unused space in it.
							output_var->AcceptNewMem((char *)this_token.circuit_token, (VarSizeType)(size_t)this_token.buf); // "buf" is the length. See comment higher above.
							goto normal_end_skip_output_var; // No need to restore circuit_token because the expression is finished.
						}
						if (i < postfix_count-1 && postfix[i+1]->symbol == SYM_ASSIGN // Next operation is ":=".
							&& stack_count && stack[stack_count-1]->symbol == SYM_VAR // i.e. let the next iteration handle it instead of doing it here.  Further below relies on this having been checked.
							&& stack[stack_count-1]->var->Type() == VAR_NORMAL) // Don't do clipboard here because: 1) AcceptNewMem() doesn't support it; 2) Could probably use Assign() and then make its result be a newly added mem_count item, but the code complexity doesn't seem worth it given the rarity.
						{
							// This section is an optimization that avoids memory allocation and an extra memcpy()
							// whenever this result is going to be assigned to a variable as the very next step.
							// See the comment section higher above for examples.
							ExprTokenType &left = *STACK_POP; // Above has already confirmed that it's SYM_VAR and VAR_NORMAL.
							// AcceptNewMem() will shrink the memory for us, via _expand(), if there's a lot of
							// extra/unused space in it.
							left.var->AcceptNewMem((char *)this_token.circuit_token, (VarSizeType)(size_t)this_token.buf);
							this_token.circuit_token = postfix[++i]->circuit_token; // Must be done AFTER above. this_token.circuit_token should have been NULL prior to this because the final right-side result of an assignment shouldn't be the last item of an AND/OR/IFF's left branch. The assignment itself would be that.
							this_token.var = left.var;   // Make the result a variable rather than a normal operand so that its
							this_token.symbol = SYM_VAR; // address can be taken, and it can be passed ByRef. e.g. &(x:=1)
							goto push_this_token;
						}
						make_result_persistent = false; // Override the default set higher above.
					} // if (this_token.circuit_token == this_token.marker)
					// Since above didn't goto, we're not done yet; so handle this memory the normal way: Mark it
					// to be freed at the time we return.
					if (mem_count == MAX_EXPR_MEM_ITEMS) // No more slots left (should be nearly impossible).
					{
						LineError(ERR_OUTOFMEM ERR_ABORT, FAIL, func.mName);
						goto abort;
					}
					mem[mem_count++] = (char *)this_token.circuit_token;
				}
				//else this_token.circuit_token==NULL, so the BIF just called didn't allocate memory to give to us.
				this_token.circuit_token = circuit_token; // Restore it to its original value.

				// HANDLE THE RESULT (unless it was already handled above due to an optimization):
				if (IS_NUMERIC(this_token.symbol)) // No need for make_result_persistent or early Assign(). Any numeric result can be considered final because it's already stored in permanent memory (namely the token itself).
					goto push_this_token; // For code simplicity, the optimization for numeric results is done at a later stage.
				//else it's a string, which might need to be moved to persistent memory further below.
				if (done_and_have_an_output_var) // RELIES ON THE IS_NUMERIC() CHECK above having been done first.
				{
					// v1.0.45: This mode improves performance by avoiding the need to copy the result into
					// more persistent memory, then avoiding the need to copy it into the defer buffer (which
					// also avoids the possibility of needing to expand that buffer).
					output_var->Assign(this_token.marker); // Marker can be used because symbol will never be SYM_VAR
					goto normal_end_skip_output_var;       // in this case. ALSO: Assign() contains an optimization that avoids actually doing the mem-copying if output_var is being assigned to itself (which can happen in cases like RegExMatch()).
				}
				// Otherwise, there's no output_var or the expression isn't finished yet, so do normal processing.
				if (!*this_token.marker) // Various make-persistent sections further below may rely on this check.
				{
					this_token.marker = ""; // Ensure it's a constant memory area, not a buf that might get overwritten soon.
					goto push_this_token; // For code simplicity, the optimization for numeric results is done at a later stage.
				}
				// Since above didn't goto, "result" is not SYM_INTEGER/FLOAT/VAR, and not "".  Therefore, it's
				// either a pointer to static memory (such as a constant string), or more likely the small buf
				// we gave to the BIF for storing small strings.  For simplicity assume its the buf, which is
				// volatile and must be made persistent if called for below.
				result = this_token.marker; // Marker can be used because symbol will never be SYM_VAR in this case.
				if (make_result_persistent) // At this stage, this means that the above wasn't able to determine its correct value yet.
					make_result_persistent = !done;
			}
			else // It's not a built-in function, or it's a built-in that was overridden with a custom function.
			{
				// If there are other instances of this function already running, either via recursion or
				// an interrupted quasi-thread, back up the local variables of the instance that lies immediately
				// beneath ours (in turn, that instance is responsible for backing up any instance that lies
				// beneath it, and so on, since when recursion collapses or threads resume, they always do so
				// in the reverse order in which they were created.
				//
				// I think the backup-and-restore approach to local variables might enhance performance over
				// other approaches, perhaps a lot.  This is because most of the time there will be no other
				// instances of a given function on the call stack, thus no backup/restore is needed, and thus
				// the function's existing local variables can be reused as though they're globals (i.e.
				// memory allocation/deallocation overhead is often completely avoided for non-recursive calls
				// to a function after the first).
				if (func.mInstances > 0) // i.e. treat negatives as zero to help catch any bugs in the way mInstances is maintained.
				{
					// Backup/restore of function's variables is needed.
					// Only when a backup is needed is it possible for this function to be calling itself recursively,
					// either directly or indirectly by means of an intermediate function.  As a consequence, it's
					// possible for this function to be passing one or more of its own params or locals to itself.
					// The following section compensates for that to handle parameters passed by-value, but it
					// doesn't correctly handle passing its own locals/params to itself ByRef, which is in the
					// help file as a known limitation.  Also, the below doesn't indicate a failure when stack
					// underflow would occur because the loop after this one needs to do that (since this
					// one will never execute if a backup isn't needed).  Note that this loop that reviews all
					// actual parameters is necessary as a separate loop from the one further below because this
					// first one's conversion must occur prior to calling BackupFunctionVars().  In addition, there
					// might be other interdependencies between formals and actuals if a function is calling itself
					// recursively.
					for (j = (actual_param_count < func.mParamCount ? actual_param_count : func.mParamCount) - 1
						, s = stack_count // Above line starts at the first formal parameter that has an actual.
						; j > -1; --j) // For each formal parameter (reverse order to mirror the nature of the stack).
					{
						// --s below moves on to the next item in the stack (without popping):  A check higher
						// above has already ensured that this won't cause stack underflow:
						ExprTokenType &this_stack_token = *stack[--s]; // Traditional, but doesn't measurably reduce code size and it's unlikely to help performance due to actual flow of control in this case.
						if (this_stack_token.symbol == SYM_VAR && !func.mParam[j].is_byref)
						{
							// Since this formal parameter is passed by value, if it's SYM_VAR, convert it to
							// SYM_OPERAND to allow the variables to be backed up and reset further below without
							// corrupting any SYM_VARs that happen to be locals or params of this very same
							// function.
							// DllCall() relies on the fact that this transformation is only done for user
							// functions, not built-in ones such as DllCall().  This is because DllCall()
							// sometimes needs the variable of a parameter for use as an output parameter.
							this_stack_token.marker = this_stack_token.var->Contents();
							this_stack_token.symbol = SYM_OPERAND;
						}
					}
					// BackupFunctionVars() will also clear each local variable and formal parameter so that
					// if that parameter or local var or is assigned a value by any other means during our call
					// to it, new memory will be allocated to hold that value rather than overwriting the
					// underlying recursed/interrupted instance's memory, which it will need intact when it's resumed.
					if (!Var::BackupFunctionVars(func, var_backup, var_backup_count)) // Out of memory.
					{
						LineError(ERR_OUTOFMEM ERR_ABORT, FAIL, func.mName);
						goto abort;
					}
				}
				//else backup is not needed because there are no other instances of this function on the call-stack.
				// So by definition, this function is not calling itself directly or indirectly, therefore there's no
				// need to do the conversion of SYM_VAR because those SYM_VARs can't be ones that were blanked out
				// due to a function exiting.  In other words, it seems impossible for a there to be no other
				// instances of this function on the call-stack and yet SYM_VAR to be one of this function's own
				// locals or formal params because it would have no legitimate origin.

				// Lexikos: Discard surplus parameters for dynamic function calls.
				if (actual_param_count > func.mParamCount)
					stack_count -= actual_param_count - func.mParamCount;

				j = func.mParamCount - 1; // The index of the last formal parameter. Relied upon by BOTH loops below.
				// The following loop will have zero iterations unless at least one formal parameter lacks an actual,
				// which should be possible only if the parameter is optional (i.e. has a default value).
				for (; j >= actual_param_count; --j) // For each formal parameter that lacks an actual (reverse order to mirror the nature of the stack).
				{
					// The following worsens performance by 7% under UPX 2.0 but is the faster method on UPX 3.0.
					// This could merely be due to unpredictable cache hits/misses in a my particular CPU.
					// In addition to being fastest, it also reduces code size by 16 bytes:
					FuncParam &this_formal_param = func.mParam[j]; // For performance and convenience.
					if (this_formal_param.is_byref) // v1.0.46.13: Allow ByRef parameters to by optional by converting an omitted-actual into a non-alias formal/local.
						this_formal_param.var->ConvertToNonAliasIfNecessary(); // Convert from alias-to-normal, if necessary.
					switch(this_formal_param.default_type)
					{
					case PARAM_DEFAULT_STR:   this_formal_param.var->Assign(this_formal_param.default_str);    break;
					case PARAM_DEFAULT_INT:   this_formal_param.var->Assign(this_formal_param.default_int64);  break;
					case PARAM_DEFAULT_FLOAT: this_formal_param.var->Assign(this_formal_param.default_double); break;
					// Lexikos: Allow dynamic function calls to pass fewer parameters than are declared.
					case PARAM_DEFAULT_NONE:  this_formal_param.var->Assign(); break;
					}
				}
				// Pop the actual number of params involved in this function-call off the stack.  Load-time
				// validation has ensured that this number is always less than or equal to the number of
				// parameters formally defined by the function.  Therefore, there should never be any leftover
				// params on the stack after this is done.  Relies upon the value of j established above:
				for (; j > -1; --j) // For each formal parameter that has a matching actual (reverse order to mirror the nature of the stack).
				{
					ExprTokenType &token = *STACK_POP; // A check higher above has already ensured that this won't cause stack underflow.
					// Below uses IS_OPERAND rather than checking for only SYM_OPERAND because the stack can contain
					// both generic and specific operands.  Specific operands were evaluated by a previous iteration
					// of this section.  Generic ones were pushed as-is onto the stack by a previous iteration.
					if (!IS_OPERAND(token.symbol)) // Haven't found a way to produce this situation yet, but safe to assume it's possible.
					{
						Var::FreeAndRestoreFunctionVars(func, var_backup, var_backup_count);
						goto abort;
					}
					// Seems to worsen performance in this case:
					//FuncParam &this_formal_param = func.mParam[j]; // For performance and convenience.
					if (func.mParam[j].is_byref)
					{
						// Note that the previous loop might not have checked things like the following because that
						// loop never ran unless a backup was needed:
						if (token.symbol != SYM_VAR)
						{
							// In most cases this condition would have been caught by load-time validation.
							// However, in the case of badly constructed double derefs, that won't be true
							// (though currently, only a double deref that resolves to a built-in variable
							// would be able to get this far to trigger this error, because something like
							// func(Array%VarContainingSpaces%) would have been caught at an earlier stage above.
							// UPDATE: In v1.0.47.06, a dynamic function call won't have validated at loadtime
							// that its parameters are the right type (since the formal definition of the function
							// won't be known until runtime).  So it seems best to treat this as a failed expression
							// rather than aborting the current thread.
							//LineError(ERR_BYREF ERR_ABORT, FAIL, func.mParam[j].var->mName);
							//Var::FreeAndRestoreFunctionVars(func, var_backup, var_backup_count);
							//goto abort;
							Var::FreeAndRestoreFunctionVars(func, var_backup, var_backup_count);
							goto abnormal_end;
						}
						func.mParam[j].var->UpdateAlias(token.var); // Make the formal parameter point directly to the actual parameter's contents.
					}
					else // This parameter is passed "by value".
						// Assign actual parameter's value to the formal parameter (which is itself a
						// local variable in the function).  
						// If token.var's Type() is always VAR_NORMAL (e.g. never the clipboard).
						// A SYM_VAR token can still happen because the previous loop's conversion of all
						// by-value SYM_VAR operands into SYM_OPERAND would not have happened if no
						// backup was needed for this function.
						func.mParam[j].var->Assign(token);
				} // for()

				aResult = func.Call(result); // Call the UDF.
				if (aResult == EARLY_EXIT || aResult == FAIL) // "Early return". See comment below.
				{
					// Take a shortcut because for backward compatibility, ACT_ASSIGNEXPR (and anything else
					// for that matter) is being aborted by this type of early return (i.e. if there's an
					// output_var, its contents are left as-is).  In other words, this expression will have
					// no result storable by the outside world.
					Var::FreeAndRestoreFunctionVars(func, var_backup, var_backup_count);
					result_to_return = NULL; // Use NULL to inform our caller that this thread is finished (whether through normal means such as Exit or a critical error).
					// Above: The callers of this function know that the value of aResult (which already contains
					// the reason for early exit) should be considered valid/meaningful only if result_to_return
					// is NULL.  aResult has already been set higher above for our caller.
					goto normal_end_skip_output_var; // output_var is left unchanged in these cases.
				}
				// Since above didn't goto, this isn't an early return, so proceed normally.
				if (done = EXPR_IS_DONE) // Resolve macro only once for use in more than one place below.
				{
					if (output_var // i.e. this is ACT_ASSIGNEXPR and we've now produced the final result.
						&& !(var_backup && g.CurrentFunc == &func && output_var->IsNonStaticLocal())) // Ordered for short-circuit performance.
						// Above line is a fix for v1.0.45.03: It detects whether output_var is among the variables
						// that are about to be restored from backup.  If it is, we can't assign to it now
						// because it's currently a local that belongs to the instance we're in the middle of
						// calling; i.e. it doesn't belong to our instance (which is beneath it on the call stack
						// until after the restore-from-backup is done later below).  And we can't assign "result"
						// to it *after* the restore because by then result may have been freed (if it happens to be
						// a local variable too).  Therefore, continue on to the normal method, which will check
						// whether "result" needs to be stored in more persistent memory.
					{
						// v1.0.45: Take a shortcut for performance.  Doing it this way saves up to two memcpy's
						// (make_result_persistent then copy into deref buffer).  In some cases, it also saves
						// from having to make_result_persistent and prevents the need to expand the deref buffer.
						// HOWEVER, the optimization described next isn't done because not only does it complicate
						// the code a lot (such as verifying that the variable isn't static, isn't ALLOC_SIMPLE,
						// isn't a ByRef to a global or some other function's local, etc.), there's also currently
						// no way to find out which function owns a particular local variable (a name lookup
						// via binary search is a possibility, but its performance probably isn't worth it)
						// Abandoned idea: When a user-defined function returns one of its local variables,
						// the contents of that local variable can be salvaged if it's about to be destroyed
						// anyway in conjunction with normal function-call cleanup. In other words, we can take
						// that local variable's memory and directly hang it onto the output_var.
						result_length = (VarSizeType)strlen(result);
						// 1.0.46.06: If the UDF has stored its result in its deref buffer, take possession
						// of that buffer, which saves a memcpy of a potentially huge string.  The cost
						// of this is that if there are any other UDF-calls pending after this one, the
						// code in their bodies will have to create another deref buffer if they need one.
						if (result == sDerefBuf && result_length >= MAX_ALLOC_SIMPLE) // Result is in their buffer and it's longer than what can fit in a SimpleHeap variable (avoids wasting SimpleHeap memory).
						{
							// AcceptNewMem() will shrink the memory for us, via _expand(), if there's a lot of
							// extra/unused space in it.
							output_var->AcceptNewMem(result, result_length);
							NULLIFY_S_DEREF_BUF // Force any UDFs called subsequently by us to create a new deref buffer because this one was just taken over by a variable.
						}
						else
							output_var->Assign(result, result_length);
						Var::FreeAndRestoreFunctionVars(func, var_backup, var_backup_count); // Do end-of-function-call cleanup (see comment above). No need to do make_result_persistent section.
						goto normal_end_skip_output_var; // Nothing more to do because it has even taken care of output_var already.
					}
					if (mActionType == ACT_EXPRESSION) // Isolated expression: Outermost function call's result will be ignored, so no need to store it.
					{
						Var::FreeAndRestoreFunctionVars(func, var_backup, var_backup_count); // Do end-of-function-call cleanup (see comment above). No need to do make_result_persistent section.
						goto normal_end_skip_output_var; // No output_var is possible for ACT_EXPRESSION.
					}
				} // if (done)
				// Otherwise (since above didn't goto), the following statement is true:
				//    !output_var || !EXPR_IS_DONE || var_backup
				// ...so do normal handling of "result". Also, no more optimizations of output_var should be
				// attempted because either we're not "done" (i.e. it isn't valid to assign to output_var yet)
				// or it isn't safe to store in output_var yet for the reason mentioned earlier.
				if (!*result) // RELIED UPON by the make-persistent check further below.
				{
					// Empty strings are returned pretty often by UDFs, such as when they don't use "return"
					// at all.  Therefore, handle them fully now, which should improve performance (since it
					// avoids all the other checking later on).  It also doesn't hurt code size because this
					// check avoids having to check for empty string in other sections later on.
					this_token.marker = ""; // Ensure it's a non-volatile address instead (read-only mem is okay for expression results).
					this_token.symbol = SYM_OPERAND; // SYM_OPERAND vs. SYM_STRING probably doesn't matter in the case of empty string, but it's used for consistency with what the other UDF handling further below does.
					Var::FreeAndRestoreFunctionVars(func, var_backup, var_backup_count);
					goto push_this_token;
				}
				// The following section is done only for UDFs (i.e. here) rather than for BIFs too because
				// the only BIFs remaining that haven't yet been fully handled by earlier optimizations are
				// those whose results are almost always tiny (small strings, since floats/integers were already
				// handled earlier). So performing the following optimization for them would probably reduce
				// average-case performance (since both performing and checking for the optimization is costly).
				// It's fairly rare that the following optimization is even be applicable because it requires
				// an assignment *internal* to an expression, such as "if not var:=func()", or "a:=b, c:=func()".
				// But it seems best to optimize these cases so that commas aren't penalized.
				if (i < postfix_count-1 && postfix[i+1]->symbol == SYM_ASSIGN  // Next operation is ":=".
					&& stack_count && stack[stack_count-1]->symbol == SYM_VAR) // i.e. let the next iteration handle it instead of doing it here.  Further below relies on this having been checked.
				{
					// This section is an optimization that avoids memory allocation and an extra memcpy()
					// whenever this result is going to be assigned to a variable as the very next step.
					// See the comment section higher above for examples.
					Var &output_var_internal = *stack[stack_count-1]->var; // Above has already confirmed that it's SYM_VAR and VAR_NORMAL.
					if (output_var_internal.Type() == VAR_NORMAL // Don't do clipboard here because don't want its result to be SYM_VAR, yet there's no place to store that result (i.e. need to continue on to make-persistent further below to get some memory for it).
						&& !(var_backup && g.CurrentFunc == &func && output_var_internal.IsNonStaticLocal())) // Ordered for short-circuit performance.
						// v1.0.46.09: The above line is a fix for a bug caused by 1.0.46.06's optimization below.
						// For details, see comments in a similar line higher above.
					{
						--stack_count; // This officially pops the lvalue off the stack (now that we know we will be handling this operation here).
						// The following section is similar to one higher above, so maintain them together.
						result_length = (VarSizeType)strlen(result);
						// 1.0.46.06: If the UDF has stored its result in its deref buffer, take possession
						// of that buffer, which saves a memcpy of a potentially huge string.  The cost
						// of this is that if there are any other UDF-calls pending after this one, the
						// code in their bodies will have to create another deref buffer if they need one.
						if (result == sDerefBuf && result_length >= MAX_ALLOC_SIMPLE) // Result is in their buffer and it's longer than what can fit in a SimpleHeap variable (avoids wasting SimpleHeap memory).
						{
							// AcceptNewMem() will shrink the memory for us, via _expand(), if there's a lot of
							// extra/unused space in it.
							output_var_internal.AcceptNewMem(result, result_length);
							NULLIFY_S_DEREF_BUF // Force any UDFs called subsequently by us to get a new deref buffer because this one was just hung onto a variable.
						}
						else
							output_var_internal.Assign(result, result_length);
						this_token.circuit_token = postfix[++i]->circuit_token; // this_token.circuit_token should have been NULL prior to this because the final right-side result of an assignment shouldn't be the last item of an AND/OR/IFF's left branch. The assignment itself would be that.
						this_token.var = &output_var_internal;   // Make the result a variable rather than a normal operand so that its
						this_token.symbol = SYM_VAR; // address can be taken, and it can be passed ByRef. e.g. &(x:=1)
						Var::FreeAndRestoreFunctionVars(func, var_backup, var_backup_count); // Do end-of-function-call cleanup (see comment above). No need to do make_result_persistent section.
						goto push_this_token;
					}
				}
				// Since above didn't goto:
				// The result just returned may need to be copied to a more persistent location.  This is done right
				// away if the result is the contents of a local variable (since all locals are about to be freed
				// and overwritten), which is assumed to be the case if it's not in the new deref buf because it's
				// difficult to distinguish between when the function returned one of its own local variables
				// rather than a global or a string/numeric literal).  The only exceptions are covered below.
				// Old method, not necessary to be so thorough because "return" always puts its result as the
				// very first item in its deref buf.  So this is commneted out in favor of the line below it:
				//if (result < sDerefBuf || result >= sDerefBuf + sDerefBufSize)
				if (result != sDerefBuf) // Not in their deref buffer (yields correct result even if sDerefBuf is NULL; also, see above.)
					// In this case, the result must be assumed to be one of their local variables (since there's
					// no way to distinguish between that and a literal string such as "abc"?). So it should be
					// immediately copied since if it's a local, it's about to be freed.
					make_result_persistent = true;
				else // The result must be in their deref buffer, perhaps due to something like "return x+3" or "return bif()" on their part.
				{
					make_result_persistent = false; // Set default to be possibly overridden below.
					if (!done) // There are more operators/operands to be evaluated, but if there are no more function calls, we don't have to make it persistent since their deref buf won't be overwritten by anything during the time we need it.
					{
						// Since there's more in the stack or postfix array to be evaluated, and since the return
						// value is in the new deref buffer, must copy result to somewhere non-volatile whenever
						// there's another function-call pending by us.  Note that an empty-string result was
						// already checked and fully handled higher above.
						// If we don't have have any more user-defined function calls pending, we can skip the
						// make-persistent section since this deref buffer will not be overwritten during the
						// period we need it.
						for (j = i + 1; j < postfix_count; ++j)
							if (postfix[j]->symbol == SYM_FUNC)
							{
								make_result_persistent = true;
								break;
							}
					}
					//else done==true, so don't have to make it persistent here because the final stage will
					// copy it from their deref buf into ours (since theirs is only deleted later, by our caller).
					// In this case, leave make_result_persistent set to false.
				} // This is the end of the section that determines the value of "make_result_persistent" for UDFs.
			} // Call to a user-defined function (UDF).

			this_token.symbol = SYM_OPERAND; // Set default. Use generic, not string, so that any operator of function call that uses this result is free to reinterpret it as an integer or float.
			if (make_result_persistent) // Both UDFs and built-in functions have ensured make_result_persistent is set.
			{
				// BELOW RELIES ON THE ABOVE ALWAYS HAVING VERFIED FULLY HANDLED RESULT BEING AN EMPTY STRING.
				// So now we know result isn't an empty string, which in turn ensures that size > 1 and length > 0,
				// which might be relied upon by things further below.
				result_size = strlen(result) + 1; // No easy way to avoid strlen currently. Maybe some future revisions to architecture will provide a length.
				// Must cast to int to avoid loss of negative values:
				if (result_size <= (int)(aDerefBufSize - (target - aDerefBuf))) // There is room at the end of our deref buf, so use it.
				{
					// Make the token's result the new, more persistent location:
					this_token.marker = (char *)memcpy(target, result, result_size); // Benches slightly faster than strcpy().
					target += result_size; // Point it to the location where the next string would be written.
				}
				else if (result_size < EXPR_SMALL_MEM_LIMIT && alloca_usage < EXPR_ALLOCA_LIMIT) // See comments at EXPR_SMALL_MEM_LIMIT.
				{
					this_token.marker = (char *)memcpy(_alloca(result_size), result, result_size); // Benches slightly faster than strcpy().
					alloca_usage += result_size; // This might put alloca_usage over the limit by as much as EXPR_SMALL_MEM_LIMIT, but that is fine because it's more of a guideline than a limit.
				}
				else // Need to create some new persistent memory for our temporary use.
				{
					// In real-world scripts the need for additonal memory allocation should be quite
					// rare because it requires a combination of worst-case situations:
					// - Called-function's return value is in their new deref buf (rare because return
					//   values are more often literal numbers, true/false, or variables).
					// - We still have more functions to call here (which is somewhat atypical).
					// - There's insufficient room at the end of the deref buf to store the return value
					//   (unusual because the deref buf expands in block-increments, and also because
					//   return values are usually small, such as numbers).
					if (mem_count == MAX_EXPR_MEM_ITEMS // No more slots left (should be nearly impossible).
						|| !(mem[mem_count] = (char *)malloc(result_size)))
					{
						LineError(ERR_OUTOFMEM ERR_ABORT, FAIL, func.mName);
						goto abort;
					}
					// Make the token's result the new, more persistent location:
					this_token.marker = (char *)memcpy(mem[mem_count], result, result_size); // Benches slightly faster than strcpy().
					++mem_count; // Must be done last.
				}
			}
			else // make_result_persistent==false
				this_token.marker = result;

			if (!func.mIsBuiltIn)
			{
				// Free the memory of all the just-completed function's local variables.  This is done in
				// both of the following cases:
				// 1) There are other instances of this function beneath us on the call-stack: Must free
				//    the memory to prevent a memory leak for any variable that existed prior to the call
				//    we just did.  Although any local variables newly created as a result of our call
				//    technically don't need to be freed, they are freed for simplicity of code and also
				//    because not doing so might result in side-effects for instances of this function that
				//    lie beneath ours that would expect such nonexistent variables to have blank contents
				//    when *they* create it.
				// 2) No other instances of this function exist on the call stack: The memory is freed and
				//    the contents made blank for these reasons:
				//    a) Prevents locals from all being static in duration, and users coming to rely on that,
				//       since in the future local variables might be implemented using a non-persistent method
				//       such as hashing (rather than maintaining a permanent list of Var*'s for each function).
				//    b) To conserve memory between calls (in case the function's locals use a lot of memory).
				//    c) To yield results consistent with when the same function is called while other instances
				//       of itself exist on the call stack.  In other words, it would be inconsistent to make
				//       all variables blank for case #1 above but not do it here in case #2.
				Var::FreeAndRestoreFunctionVars(func, var_backup, var_backup_count);
			} // if (!func.mIsBuiltIn)
			goto push_this_token;
		} // if (this_token.symbol == SYM_FUNC)

		if (this_token.symbol == SYM_IFF_ELSE) // This is encountered when a ternary's condition was found to be false by a prior iteration.
		{
			if (this_token.circuit_token // This ternary's result is some other ternary's condition (somewhat rare, so the simple method used here isn't much a concern for performance optimization).
				&& stack_count) // Prevent underflow (this check might not be necessary; so it's just in case there's a way it can happen).
			{
				// To support *cascading* short-circuit when ternary/IFF's are nested inside each other, pop the
				// topmost operand off the stack to modify its circuit_token.  The routine below will then
				// use this as the parent IFF's *condition*, which is an non-operand of sorts because it's
				// used only to determine which branch of an IFF will become the operand/result of this IFF.
				circuit_token = this_token.circuit_token; // Temp copy to avoid overwrite by the next line.
				this_token = *STACK_POP; // Struct copy.  Doing it this way is more maintainable than other methods, and is unlikely to perform much worse.
				this_token.circuit_token = circuit_token;
				goto non_null_circuit_token; // Must do this so that it properly evaluates this_token as the next ternary's condition.
			}
			// Otherwise, ignore it because its final result has already been evaluated and pushed onto the
			// stack via prior iterations.  In other words, this ELSE branch was the IFF's final result, which
			// is now topmost on the stack for use as an operand by a future operator.
			continue;
		}

		// Since the above didn't "goto" or continue, this token must be a unary or binary operator.
		// Get the first operand for this operator (for non-unary operators, this is the right-side operand):
		if (!stack_count) // Prevent stack underflow.  An expression such as -*3 causes this.
			goto abnormal_end;
		ExprTokenType &right = *STACK_POP;
		// Below uses IS_OPERAND rather than checking for only SYM_OPERAND because the stack can contain
		// both generic and specific operands.  Specific operands were evaluated by a previous iteration
		// of this section.  Generic ones were pushed as-is onto the stack by a previous iteration.
		if (!IS_OPERAND(right.symbol)) // Haven't found a way to produce this situation yet, but safe to assume it's possible.
			goto abnormal_end;

		// The following check is done after popping "right" off the stack because a prior iteration has set up
		// SYM_IFF_THEN to be a unary operator of sorts.
		if (this_token.symbol == SYM_IFF_THEN) // This is encountered when a ternary's condition was found to be true by a prior iteration.
		{
			if (!this_token.circuit_token) // This check is needed for syntax errors such as "1 ? 2" (no matching else) and perhaps other unusual circumstances.
				goto abnormal_end; // Seems best to consider it a syntax error rather than supporting partial functionality (hard to imagine much legitimate need to omit an ELSE).
			// SYM_IFF_THEN is encountered only when a previous iteration has determined that the ternary's condition
			// is true.  At this stage, the ternary's "THEN" branch has already been evaluated and stored in
			// "right".  So skip over its "else" branch (short-circuit) because that doesn't need to be evaluated.
			for (++i; postfix[i] != this_token.circuit_token; ++i); // Should always be found, so no need to check postfix_count.
			// And very soon, the outer loop's ++i will skip over the SYM_IFF_ELSE just found above.
			right.circuit_token = this_token.circuit_token->circuit_token; // Can be NULL (in fact, it usually is).
			this_token = right;   // Struct copy to set things up for push_this_token, which in turn is needed
			goto push_this_token; // (rather than a simple STACK_PUSH(right)) because it checks for *cascading* short circuit in cases where this ternary's result is the boolean condition of another ternary.
		}

		if (this_token.symbol == SYM_COMMA) // This can only be a statement-separator comma, not a function comma, since function commas weren't put into the postfix array.
			// Do nothing other than discarding the right-side operand that was just popped off the stack.
			// This collapses the two sub-statements delimated by a given comma into a single result for
			// subequent uses by another operator.  Unlike C++, the leftmost operand is preserved, not the
			// rightmost.  This is because it's faster to just discard the topmost item on the stack, but
			// more importantly it allows ACT_ASSIGNEXPR, ACT_ADD, and others to work properly.  For example:
			//    Var:=5, Var1:=(Var2:=1, Var3:=2)
			// Without the behavior implemented here, the above would wrongly put Var3's rvalue into Var2.
			continue;

		if (this_token.symbol != SYM_ASSIGN) // SYM_ASSIGN doesn't need "right" to be resolved.
		{
			// If the operand is still generic/undetermined, find out whether it is a string, integer, or float:
			switch(right.symbol)
			{
			case SYM_VAR:
				right_contents = right.var->Contents(); // Can be the clipboard in this case, but it works even then.
				right_is_number = IsPureNumeric(right_contents, true, false, true);
				break;
			case SYM_OPERAND:
				right_contents = right.marker;
				right_is_number = IsPureNumeric(right_contents, true, false, true);
				break;
			case SYM_STRING:
				right_contents = right.marker;
				right_is_number = PURE_NOT_NUMERIC; // Explicitly-marked strings are not numeric, which allows numeric strings to be compared as strings rather than as numbers.
			default: // INTEGER or FLOAT
				// right_contents is left uninitialized for performance and to catch bugs.
				right_is_number = right.symbol;
			}
		}

		// IF THIS IS A UNARY OPERATOR, we now have the single operand needed to perform the operation.
		// The cases in the switch() below are all unary operators.  The other operators are handled
		// in the switch()'s default section:
		sym_assign_var = NULL; // Set default for use at the bottom of the following switch().
		switch (this_token.symbol)
		{
		case SYM_AND: // These are now unary operators because short-circuit has made them so.  If the AND/OR
		case SYM_OR:  // had short-circuited, we would never be here, so this is the right branch of a non-short-circuit AND/OR.
			if (right_is_number == PURE_INTEGER)
				this_token.value_int64 = (right.symbol == SYM_INTEGER ? right.value_int64 : ATOI64(right_contents)) != 0;
			else if (right_is_number == PURE_FLOAT)
				this_token.value_int64 = (right.symbol == SYM_FLOAT ? right.value_double : atof(right_contents)) != 0.0;
			else // This is either a non-numeric string or a numeric raw literal string such as "123".
				// All non-numeric strings are considered TRUE here.  In addition, any raw literal string,
				// even "0", is considered to be TRUE.  This relies on the fact that right.symbol will be
				// SYM_OPERAND/generic (and thus handled higher above) for all pure-numeric strings except
				// explicit raw literal strings.  Thus, if something like !"0" ever appears in an expression,
				// it evaluates to !true.  EXCEPTION: Because "if x" evaluates to false when X is blank,
				// it seems best to have "if !x" evaluate to TRUE.
				this_token.value_int64 = *right_contents != '\0';
			this_token.symbol = SYM_INTEGER; // Result of AND or OR is always a boolean integer (one or zero).
			break;

		case SYM_NEGATIVE:  // Unary-minus.
			if (right_is_number == PURE_INTEGER)
				this_token.value_int64 = -(right.symbol == SYM_INTEGER ? right.value_int64 : ATOI64(right_contents));
			else if (right_is_number == PURE_FLOAT)
				// Overwrite this_token's union with a float. No need to have the overhead of ATOF() since PURE_FLOAT is never hex.
				this_token.value_double = -(right.symbol == SYM_FLOAT ? right.value_double : atof(right_contents));
			else // String.
			{
				// Seems best to consider the application of unary minus to a string, even a quoted string
				// literal such as "15", to be a failure.  UPDATE: For v1.0.25.06, invalid operations like
				// this instead treat the operand as an empty string.  This avoids aborting a long, complex
				// expression entirely just because on of its operands is invalid.  However, the net effect
				// in most cases might be the same, since the empty string is a non-numeric result and thus
				// will cause any operator it is involved with to treat its other operand as a string too.
				// And the result of a math operation on two strings is typically an empty string.
				this_token.marker = "";
				this_token.symbol = SYM_STRING;
				break;
			}
			// Since above didn't "break":
			this_token.symbol = right_is_number; // Convert generic SYM_OPERAND into a specific type: float or int.
			break;

		// Both nots are equivalent at this stage because precedence was already acted upon by infix-to-postfix:
		case SYM_LOWNOT:  // The operator-word "not".
		case SYM_HIGHNOT: // The symbol !
			if (right_is_number == PURE_INTEGER)
				this_token.value_int64 = !(right.symbol == SYM_INTEGER ? right.value_int64 : ATOI64(right_contents));
			else if (right_is_number == PURE_FLOAT) // Convert to float, not int, so that a number between 0.0001 and 0.9999 is considered "true".
				// Using ! vs. comparing explicitly to 0.0 might generate faster code, and K&R implies it's okay:
				this_token.value_int64 = !(right.symbol == SYM_FLOAT ? right.value_double : atof(right_contents));
			else // This is either a non-numeric string or a numeric raw literal string such as "123".
				// All non-numeric strings are considered TRUE here.  In addition, any raw literal string,
				// even "0", is considered to be TRUE.  This relies on the fact that right.symbol will be
				// SYM_OPERAND/generic (and thus handled higher above) for all pure-numeric strings except
				// explicit raw literal strings.  Thus, if something like !"0" ever appears in an expression,
				// it evaluates to !true.  EXCEPTION: Because "if x" evaluates to false when X is blank,
				// it seems best to have "if !x" evaluate to TRUE.
				this_token.value_int64 = !*right_contents; // i.e. result is false except for empty string because !"string" is false.
			this_token.symbol = SYM_INTEGER; // Result of above is always a boolean integer (one or zero).
			break;

		case SYM_POST_INCREMENT: // These were added in v1.0.46.  It doesn't seem worth translating them into
		case SYM_POST_DECREMENT: // += and -= at load-time or during the tokenizing phase higher above because 
		case SYM_PRE_INCREMENT:  // it might introduce precedence problems, plus the post-inc/dec's nature is
		case SYM_PRE_DECREMENT:  // unique among all the operators in that it pushes an operand before the evaluation.
			is_pre_op = (this_token.symbol >= SYM_PRE_INCREMENT); // Store this early because its symbol will soon be overwritten.
			if (right.symbol != SYM_VAR || right_is_number == PURE_NOT_NUMERIC) // Invalid operation.
			{
				if (right.symbol == SYM_VAR) // Thus due to the above check, it's a non-numeric target such as ++i when "i" is blank or contains text. This line was fixed in v1.0.46.16.
				{
					right.var->Assign(); // If target var contains "" or "non-numeric text", make it blank. Clipboard is also supported here.
					if (is_pre_op)
					{
						// v1.0.46.01: For consistency, it seems best to make the result of a pre-op be a
						// variable whenever a variable came in.  This allows its address to be taken, and it
						// to be passed byreference, and other SYM_VAR behaviors, even if the operation itself
						// produces a blank value.
						// KNOWN LIMITATION: Although this behavior is convenient to have have, I realize now
						// that it produces at least one weird effect: whenever a binary operator's operands
						// both use a pre-op on the same variable, or whenever two or more of a function-call's
						// parameters both do a pre-op on the same variable, that variable will have the same
						// value at the time the binary operator or function-call is evaluated.  For example:
						//    y = 1
						//    x = ++y + ++y  ; Yields 6 not 5.
						// However, if you think about the situations anyone would intentionally want to do
						// the above or a function-call with two or more pre-ops in its parameters, it seems
						// so extremely rare that retaining the existing behavior might be superior because of:
						// 1) Convenience: It allows ++x to be passed ByRef, it's address taken.  Less importantly,
						//    it also allows ++++x to work.
						// 2) Backward compatibility: Some existing scripts probably already rely on the fact that
						//    ++x and --x produce an lvalue (though it's undocumented).
						if (right.var->Type() == VAR_NORMAL)
						{
							this_token.var = right.var;  // Make the result a variable rather than a normal operand so that its
							this_token.symbol = SYM_VAR; // address can be taken, and it can be passed ByRef. e.g. &(++x)
							break;
						}
						//else VAR_CLIPBOARD, which is allowed in only when it's the lvalue of an assignent or
						// inc/dec.  So fall through to make the result blank because clipboard isn't allowed as
						// SYM_VAR beyond this point (to simplify the code and improve maintainability).
					}
					//else post_op against non-numeric target-var.  Fall through to below to yield blank result.
				}
				//else target isn't a var.  Fall through to below to yield blank result.
				this_token.marker = "";          // Make the result blank to indicate invalid operation
				this_token.symbol = SYM_STRING;  // (assign to non-lvalue or increment/decrement a non-number).
				break;
			} // end of "invalid operation" block.

			// DUE TO CODE SIZE AND PERFORMANCE decided not to support things like the following:
			// -> ++++i ; This one actually works because pre-ops produce a variable (usable by future pre-ops).
			// -> i++++ ; Fails because the first ++ produces an operand that isn't a variable.  It could be
			//    supported via a cascade loop here to pull all remaining consective post/pre ops out of
			//    the postfix array and apply them to "delta", but it just doesn't seem worth it.
			// -> --Var++ ; Fails because ++ has higher precedence than --, but it produces an operand that isn't
			//    a variable, so the "--" fails.  Things like --Var++ seem pointless anyway because they seem
			//    nearly identical to the sub-expression (Var+1)? Anyway, --Var++ could probably be supported
			//    using the loop described in the previous example.
			delta = (this_token.symbol == SYM_POST_INCREMENT || this_token.symbol == SYM_PRE_INCREMENT) ? 1 : -1;
			if (right_is_number == PURE_INTEGER)
			{
				this_token.value_int64 = (right.symbol == SYM_INTEGER) ? right.value_int64 : ATOI64(right_contents);
				right.var->Assign(this_token.value_int64 + delta);
			}
			else // right_is_number must be PURE_FLOAT because it's the only remaining alternative.
			{
				// Uses atof() because no need to have the overhead of ATOF() since PURE_FLOAT is never hex.
				this_token.value_double = (right.symbol == SYM_FLOAT) ? right.value_double : atof(right_contents);
				right.var->Assign(this_token.value_double + delta);
			}
			if (is_pre_op)
			{
				// Push the variable itself so that the operation will have already taken effect for whoever
				// uses this operand/result in the future (i.e. pre-op vs. post-op).
				if (right.var->Type() == VAR_NORMAL)
				{
					this_token.var = right.var;  // Make the result a variable rather than a normal operand so that its
					this_token.symbol = SYM_VAR; // address can be taken, and it can be passed ByRef. e.g. &(++x)
				}
				else // VAR_CLIPBOARD, which is allowed in only when it's the lvalue of an assignent or inc/dec.
				{
					// Clipboard isn't allowed as SYM_VAR beyond this point (to simplify the code and
					// improve maintainability).  So use the new contents of the clipboard as the result,
					// rather than the clipboard itself.
					if (right_is_number == PURE_INTEGER)
						this_token.value_int64 += delta;
					else // right_is_number must be PURE_FLOAT because it's the only alternative remaining.
						this_token.value_double += delta;
					this_token.symbol = right_is_number; // Set the symbol type to match the double or int64 that was already stored higher above.
				}
			}
			else // Post-inc/dec, so the non-delta version, which was already stored in this_token, should get pushed.
				this_token.symbol = right_is_number; // Set the symbol type to match the double or int64 that was already stored higher above.
			break;

		case SYM_ADDRESS: // Take the address of a variable.
			if (right.symbol == SYM_VAR) // At this stage, SYM_VAR is always a normal variable, never a built-in one, so taking its address should be safe.
			{
				this_token.symbol = SYM_INTEGER;
				this_token.value_int64 = (__int64)right_contents;
			}
			else // Invalid, so make it a localized blank value.
			{
				this_token.marker = "";
				this_token.symbol = SYM_STRING;
			}
			break;

		case SYM_DEREF:   // Dereference an address to retrieve a single byte.
		case SYM_BITNOT:           // The tilde (~) operator.
			if (right_is_number == PURE_INTEGER) // But in this case, it can be hex, so use ATOI64().
				right_int64 = right.symbol == SYM_INTEGER ? right.value_int64 : ATOI64(right_contents);
			else if (right_is_number == PURE_FLOAT)
				// No need to have the overhead of ATOI64() since PURE_FLOAT can't be hex:
				right_int64 = right.symbol == SYM_FLOAT ? (__int64)right.value_double : _atoi64(right_contents);
			else // String.  Seems best to consider the application of unary minus to a string, even a quoted string literal such as "15", to be a failure.
			{
				this_token.marker = "";
				this_token.symbol = SYM_STRING;
				break;
			}
			// Since above didn't "break":
			if (this_token.symbol == SYM_BITNOT)
			{
				// Note that it is not legal to perform ~, &, |, or ^ on doubles.  Because of this, and also to
				// conform to the behavior of the Transform command, any floating point operand is truncated to
				// an integer above.
				if (right_int64 < 0 || right_int64 > UINT_MAX)
					this_token.value_int64 = ~right_int64;
				else // See comments at TRANS_CMD_BITNOT for why it's done this way:
					this_token.value_int64 = (size_t)~(DWORD)right_int64; // Casting this way avoids compiler warning.
			}
			else // SYM_DEREF
			{
				// Reasons for resolving *Var to a number rather than a single-char string:
				// 1) More consistent with future uses of * that might operate on the address of 2-byte,
				//    4-byte, and 8-byte targets.
				// 2) Performs better in things like ExtractInteger() that would otherwise have to call Asc().
				// 3) Converting it to a one-char string would add no value beyond convenience because script
				//    could do "if (*var = 65)" if it's concerned with avoiding a Chr() call for performance
				//    reasons.  Also, it seems somewhat rare that a script will access a string's characters
				//    one-by-one via the * method because that a parsing loop can already do that more easily.
				// 4) Reduces code size and improves performance (however, the single-char string method would
				//    use _alloca(2) to get some temporary memory, so it wouldn't be too bad in performance).
				//
				// The following does a basic bounds check to prevent crashes due to dereferencing addresses
				// that are obviously bad.  In terms of percentage impact on performance, this seems quite
				// justified.  In the future, could also put a __try/__except block around this (like DllCall
				// uses) to prevent buggy scripts from crashing.  In addition to ruling out the dereferencing of
				// a NULL address, the >255 check also rules out common-bug addresses (I don't think addresses
				// this low can realistically ever be legitimate, but it would be nice to get confirmation).
				// For simplicity and due to rarity, a zero is yielded in such cases rather than an empty string.
				this_token.value_int64 = (right_int64 < 256 || right_int64 > 0xFFFFFFFF)
					? 0 : this_token.value_int64 = *(UCHAR *)right_int64; // Dereference to extract one unsigned character, just like Asc().
			}
			this_token.symbol = SYM_INTEGER; // Must be done only after its old value was used above. v1.0.36.07: Fixed to be SYM_INTEGER vs. right_is_number for SYM_BITNOT.
			break;

		default: // NON-UNARY OPERATOR.
			// GET THE SECOND (LEFT-SIDE) OPERAND FOR THIS OPERATOR:
			if (!stack_count) // Prevent stack underflow.
				goto abnormal_end;
			ExprTokenType &left = *STACK_POP; // i.e. the right operand always comes off the stack before the left.
			if (!IS_OPERAND(left.symbol)) // Haven't found a way to produce this situation yet, but safe to assume it's possible.
				goto abnormal_end;

			if (IS_ASSIGNMENT_EXCEPT_POST_AND_PRE(this_token.symbol)) // v1.0.46: Added support for various assignment operators.
			{
				if (left.symbol != SYM_VAR)
				{
					this_token.marker = "";          // Make the result blank to indicate invalid operation
					this_token.symbol = SYM_STRING;  // (assign to non-lvalue).
					break; // Equivalent to "goto push_this_token" in this case.
				}
				switch(this_token.symbol)
				{
				case SYM_ASSIGN: // Listed first for performance (it's probably the most common because things like ++ and += aren't expressions when they're by themselves on a line).
					left.var->Assign(right); // left.var can be VAR_CLIPBOARD in this case.
					if (left.var->Type() == VAR_CLIPBOARD) // v1.0.46.01: Clipboard is present as SYM_VAR, but only for assign-to-clipboard so that built-in functions and other code sections don't need handling for VAR_CLIPBOARD.
					{
						circuit_token = this_token.circuit_token; // Temp copy to avoid overwrite by the next line.
						this_token = right; // Struct copy.  Doing it this way is more maintainable than other methods, and is unlikely to perform much worse.
						this_token.circuit_token = circuit_token;
					}
					else
					{
						this_token.var = left.var;   // Make the result a variable rather than a normal operand so that its
						this_token.symbol = SYM_VAR; // address can be taken, and it can be passed ByRef. e.g. &(x:=1)
					}
					goto push_this_token;
				case SYM_ASSIGN_ADD:           this_token.symbol = SYM_ADD; break;
				case SYM_ASSIGN_SUBTRACT:      this_token.symbol = SYM_SUBTRACT; break;
				case SYM_ASSIGN_MULTIPLY:      this_token.symbol = SYM_MULTIPLY; break;
				case SYM_ASSIGN_DIVIDE:        this_token.symbol = SYM_DIVIDE; break;
				case SYM_ASSIGN_FLOORDIVIDE:   this_token.symbol = SYM_FLOORDIVIDE; break;
				case SYM_ASSIGN_BITOR:         this_token.symbol = SYM_BITOR; break;
				case SYM_ASSIGN_BITXOR:        this_token.symbol = SYM_BITXOR; break;
				case SYM_ASSIGN_BITAND:        this_token.symbol = SYM_BITAND; break;
				case SYM_ASSIGN_BITSHIFTLEFT:  this_token.symbol = SYM_BITSHIFTLEFT; break;
				case SYM_ASSIGN_BITSHIFTRIGHT: this_token.symbol = SYM_BITSHIFTRIGHT; break;
				case SYM_ASSIGN_CONCAT:        this_token.symbol = SYM_CONCAT; break;
				}
				// Since above didn't goto/break, this is an assignment other than SYM_ASSIGN, so it needs further
				// evaluation later below before the assignment will actually be made.
				sym_assign_var = left.var; // This tells the bottom of this switch() to do extra steps for this assignment.
			}

			// The following section needs done even for assignments such as += because the type of value
			// inside the target variable (integer vs. float vs. string) must be known to determine how
			// the operation should proceed.
			// Since above didn't goto/break, this is a non-unary operator that needs further processing.
			// If the operand is still generic/undetermined, find out whether it is a string, integer, or float:
			switch(left.symbol)
			{
			case SYM_VAR:
				left_contents = left.var->Contents();
				left_is_number = IsPureNumeric(left_contents, true, false, true);
				break;
			case SYM_OPERAND:
				left_contents = left.marker;
				left_is_number = IsPureNumeric(left_contents, true, false, true);
				break;
			case SYM_STRING:
				left_contents = left.marker;
				left_is_number = PURE_NOT_NUMERIC;
			default:
				// left_contents is left uninitialized for performance and to catch bugs.
				left_is_number = left.symbol;
			}

			if (!right_is_number || !left_is_number || this_token.symbol == SYM_CONCAT)
			{
				// Above check has ensured that at least one of them is a string.  But the other
				// one might be a number such as in 5+10="15", in which 5+10 would be a numerical
				// result being compared to the raw string literal "15".
				switch (right.symbol)
				{
				// Seems best to obey SetFormat for these two, though it's debatable:
				case SYM_INTEGER: right_string = ITOA64(right.value_int64, right_buf); break;
				case SYM_FLOAT: snprintf(right_buf, sizeof(right_buf), g.FormatFloat, right.value_double); right_string = right_buf; break;
				default: right_string = right_contents; // SYM_STRING/SYM_OPERAND/SYM_VAR, which is already in the right format.
				}

				switch (left.symbol)
				{
				// Seems best to obey SetFormat for these two, though it's debatable:
				case SYM_INTEGER: left_string = ITOA64(left.value_int64, left_buf); break;
				case SYM_FLOAT: snprintf(left_buf, sizeof(left_buf), g.FormatFloat, left.value_double); left_string = left_buf; break;
				default: left_string = left_contents; // SYM_STRING/SYM_OPERAND/SYM_VAR, which is already in the right format.
				}

				result_symbol = SYM_INTEGER; // Set default.  Boolean results are treated as integers.
				switch(this_token.symbol)
				{
				case SYM_EQUAL:     this_token.value_int64 = !((g.StringCaseSense == SCS_INSENSITIVE)
										? stricmp(left_string, right_string)
										: lstrcmpi(left_string, right_string)); break; // i.e. use the "more correct mode" except when explicitly told to use the fast mode (v1.0.43.03).
				case SYM_EQUALCASE: this_token.value_int64 = !strcmp(left_string, right_string); break; // Case sensitive.
				// The rest all obey g.StringCaseSense since they have no case sensitive counterparts:
				case SYM_NOTEQUAL:  this_token.value_int64 = g_strcmp(left_string, right_string) ? 1 : 0; break;
				case SYM_GT:        this_token.value_int64 = g_strcmp(left_string, right_string) > 0; break;
				case SYM_LT:        this_token.value_int64 = g_strcmp(left_string, right_string) < 0; break;
				case SYM_GTOE:      this_token.value_int64 = g_strcmp(left_string, right_string) > -1; break;
				case SYM_LTOE:      this_token.value_int64 = g_strcmp(left_string, right_string) < 1; break;

				case SYM_CONCAT:
					// Even if the left or right is "", must copy the result to temporary memory, at least
					// when integers and floats had to be converted to temporary strings above.
					// Binary clipboard is ignored because it's documented that except for certain features,
					// binary clipboard variables are seen only up to the first binary zero (mostly to
					// simplify the code).
					right_length = (right.symbol == SYM_VAR) ? right.var->LengthIgnoreBinaryClip() : strlen(right_string);
					if (sym_assign_var // Since "right" is being appended onto a variable ("left"), an optimization is possible.
						&& sym_assign_var->AppendIfRoom(right_string, (VarSizeType)right_length)) // But only if the target variable has enough remaining capacity.
					{
						// AppendIfRoom() always fails for VAR_CLIPBOARD, so below won't execute for it (which is
						// good because don't want clipboard to stay as SYM_VAR after the assignment. This is
						// because it simplifies the code not to have to worry about VAR_CLIPBOARD in BIFs, etc.)
						this_token.var = sym_assign_var; // Make the result a variable rather than a normal operand so that its
						this_token.symbol = SYM_VAR;     // address can be taken, and it can be passed ByRef. e.g. &(x+=1)
						goto push_this_token; // Skip over all other sections such as subsequent checks of sym_assign_var because it was all taken care of here.
					}
					// Otherwise, fall back to the other concat methods:
					left_length = (left.symbol == SYM_VAR) ? left.var->LengthIgnoreBinaryClip() : strlen(left_string);
					result_size = right_length + left_length + 1;

					if (output_var && EXPR_IS_DONE) // i.e. this is ACT_ASSIGNEXPR and we're at the final operator, a concat.
						temp_var = output_var;
					else if (i < postfix_count-1 && postfix[i+1]->symbol == SYM_ASSIGN // Next operation is ":=".
						&& stack_count && stack[stack_count-1]->symbol == SYM_VAR // i.e. let the next iteration handle it instead of doing it here.  Further below relies on this having been checked.
						&& stack[stack_count-1]->var->Type() == VAR_NORMAL) // Don't do clipboard here because: 1) AcceptNewMem() doesn't support it; 2) Could probably use Assign() and then make its result be a newly added mem_count item, but the code complexity doesn't seem worth it given the rarity.
						temp_var = stack[stack_count-1]->var;
					else
						temp_var = NULL;

					if (temp_var)
					{
						result = temp_var->Contents();
						if (result == left_string) // This is something like x := x . y, so simplify it to x .= y
						{
							// MUST DO THE ABOVE CHECK because the next section further below might free the
							// destination memory before doing the operation. Thus, if the destination is the
							// same as one of the sources, freeing it beforehand would obviously be a problem.
							if (temp_var->AppendIfRoom(right_string, (VarSizeType)right_length))
							{
								if (temp_var == output_var)
									goto normal_end_skip_output_var; // Nothing more to do because it has even taken care of output_var already.
								else // temp_var is from look-ahead to a future assignment.
								{
									this_token.circuit_token = postfix[++i]->circuit_token; // this_token.circuit_token should have been NULL prior to this because the final right-side result of an assignment shouldn't be the last item of an AND/OR/IFF's left branch. The assignment itself would be that.
									this_token.var = STACK_POP->var; // Make the result a variable rather than a normal operand so that its
									this_token.symbol = SYM_VAR;     // address can be taken, and it can be passed ByRef. e.g. &(x:=1)
									goto push_this_token;
								}
							}
							//else no optimizations are possible because: 1) No room; 2) The overlap between the
							// source and dest requires temporary memory.  So fall through to the slower method.
						}
						else if (result != right_string) // No overlap between the two sources and dest.
						{
							// The check above assumes that only a complete equality/overlap is possible,
							// not a partial overlap.  A partial overlap between the memory of two variables
							// seems impossible for a script to produce.  But if it ever does happen, the
							// Assign() below would free part or all of one of the sources before doing
							// the concat, which would corrupt the result.
							// Optimize by copying directly into the target variable rather than the intermediate
							// step of putting into temporary memory.
							if (!temp_var->Assign(NULL, (VarSizeType)result_size - 1)) // Resize the destination, if necessary.
								goto abort; // Above should have already reported the error.
							result = temp_var->Contents(); // Call Contents() AGAIN because Assign() may have changed it.
							if (left_length)
								memcpy(result, left_string, left_length);  // Not +1 because don't need the zero terminator.
							memcpy(result + left_length, right_string, right_length + 1); // +1 to include its zero terminator.
							temp_var->Close(); // Mostly just to reset the VAR_ATTRIB_BINARY_CLIP attribute and for maintainability.
							if (temp_var == output_var)
								goto normal_end_skip_output_var; // Nothing more to do because it has even taken care of output_var already.
							else // temp_var is from look-ahead to a future assignment.
							{
								this_token.circuit_token = postfix[++i]->circuit_token; // this_token.circuit_token should have been NULL prior to this because the final right-side result of an assignment shouldn't be the last item of an AND/OR/IFF's left branch. The assignment itself would be that.
								this_token.var = STACK_POP->var; // Make the result a variable rather than a normal operand so that its
								this_token.symbol = SYM_VAR;     // address can be taken, and it can be passed ByRef. e.g. &(x:=1)
								goto push_this_token;
							}
						}
						//else result==right_string (e.g. x := y . x).  Although this could be optimized by 
						// moving memory around inside output_var (if it has enough capacity), it seems more
						// complicated than it's worth given the rarity of this.  It probably wouldn't save
						// much time anyway due to the memory-moves inside output_var.  So just fall through
						// to the normal method.
					} // if (temp_var)

					// Since above didn't "goto", it didn't find a way to optimize this concat.
					// So fall back to the standard method.
					// The following section is similar to the one for "symbol == SYM_FUNC", so they
					// should be maintained together.
					// The following isn't done because there's a memcpy() further below which would also
					// have to check it, which hurts maintainability.  This doesn't seem worth it since
					// it's unlikely to be the empty string in the case of concat.
					//if (result_size == 1)
					//	this_token.marker = "";
					//else
					// Must cast to int to avoid loss of negative values:
					if (result_size <= (int)(aDerefBufSize - (target - aDerefBuf))) // There is room at the end of our deref buf, so use it.
					{
						this_token.marker = target;
						target += result_size;  // Adjust target for potential future use by another concat or functionc call.
					}
					else if (result_size < EXPR_SMALL_MEM_LIMIT && alloca_usage < EXPR_ALLOCA_LIMIT) // See comments at EXPR_SMALL_MEM_LIMIT.
					{
						this_token.marker = (char *)_alloca(result_size);
						alloca_usage += result_size; // This might put alloca_usage over the limit by as much as EXPR_SMALL_MEM_LIMIT, but that is fine because it's more of a guideline than a limit.
					}
					else // Need to create some new persistent memory for our temporary use.
					{
						// See the nearly identical section higher above for comments:
						if (mem_count == MAX_EXPR_MEM_ITEMS // No more slots left (should be nearly impossible).
							|| !(this_token.marker = mem[mem_count] = (char *)malloc(result_size)))
						{
							LineError(ERR_OUTOFMEM ERR_ABORT);
							goto abort;
						}
						++mem_count;
					}
					if (left_length)
						memcpy(this_token.marker, left_string, left_length);  // Not +1 because don't need the zero terminator.
					memcpy(this_token.marker + left_length, right_string, right_length + 1); // +1 to include its zero terminator.

					// For this new concat operator introduced in v1.0.31, it seems best to treat the
					// result as a SYM_STRING if either operand is a SYM_STRING.  That way, when the
					// result of the operation is later used, it will be a real string even if pure numeric,
					// which allows an exact string match to be specified even when the inputs are
					// technically numeric; e.g. the following should be true only if (Var . 33 = "1133") 
					result_symbol = (left.symbol == SYM_STRING || right.symbol == SYM_STRING) ? SYM_STRING: SYM_OPERAND;
					break;

				default:
					// Other operators do not support string operands, so the result is an empty string.
					this_token.marker = "";
					result_symbol = SYM_STRING;
				}
				this_token.symbol = result_symbol; // Must be done only after the switch() above.
			}

			else if (right_is_number == PURE_INTEGER && left_is_number == PURE_INTEGER && this_token.symbol != SYM_DIVIDE
				|| this_token.symbol <= SYM_BITSHIFTRIGHT && this_token.symbol >= SYM_BITOR) // Check upper bound first for short-circuit performance (because operators like +-*/ are much more frequently used).
			{
				// Because both are integers and the operation isn't division, the result is integer.
				// The result is also an integer for the bitwise operations listed in the if-statement
				// above.  This is because it is not legal to perform ~, &, |, or ^ on doubles, and also
				// because this behavior conforms to that of the Transform command.  Any floating point
				// operands are truncated to integers prior to doing the bitwise operation.

				switch (right.symbol)
				{
				case SYM_INTEGER: right_int64 = right.value_int64; break;
				case SYM_FLOAT: right_int64 = (__int64)right.value_double; break;
				default: right_int64 = ATOI64(right_contents); // SYM_OPERAND or SYM_VAR
				// It can't be SYM_STRING because in here, both right and left are known to be numbers
				// (otherwise an earlier "else if" would have executed instead of this one).
				}

				switch (left.symbol)
				{
				case SYM_INTEGER: left_int64 = left.value_int64; break;
				case SYM_FLOAT: left_int64 = (__int64)left.value_double; break;
				default: left_int64 = ATOI64(left_contents); // SYM_OPERAND or SYM_VAR
				// It can't be SYM_STRING because in here, both right and left are known to be numbers
				// (otherwise an earlier "else if" would have executed instead of this one).
				}

				result_symbol = SYM_INTEGER; // Set default.
				switch(this_token.symbol)
				{
				// The most common cases are kept up top to enhance performance if switch() is implemented as if-else ladder.
				case SYM_ADD:      this_token.value_int64 = left_int64 + right_int64; break;
				case SYM_SUBTRACT: this_token.value_int64 = left_int64 - right_int64; break;
				case SYM_MULTIPLY: this_token.value_int64 = left_int64 * right_int64; break;
				// A look at K&R confirms that relational/comparison operations and logical-AND/OR/NOT
				// always yield a one or a zero rather than arbitrary non-zero values:
				case SYM_EQUALCASE: // Same behavior as SYM_EQUAL for numeric operands.
				case SYM_EQUAL:    this_token.value_int64 = left_int64 == right_int64; break;
				case SYM_NOTEQUAL: this_token.value_int64 = left_int64 != right_int64; break;
				case SYM_GT:       this_token.value_int64 = left_int64 > right_int64; break;
				case SYM_LT:       this_token.value_int64 = left_int64 < right_int64; break;
				case SYM_GTOE:     this_token.value_int64 = left_int64 >= right_int64; break;
				case SYM_LTOE:     this_token.value_int64 = left_int64 <= right_int64; break;
				case SYM_BITAND:   this_token.value_int64 = left_int64 & right_int64; break;
				case SYM_BITOR:    this_token.value_int64 = left_int64 | right_int64; break;
				case SYM_BITXOR:   this_token.value_int64 = left_int64 ^ right_int64; break;
				case SYM_BITSHIFTLEFT:  this_token.value_int64 = left_int64 << right_int64; break;
				case SYM_BITSHIFTRIGHT: this_token.value_int64 = left_int64 >> right_int64; break;
				case SYM_FLOORDIVIDE:
					// Since it's integer division, no need for explicit floor() of the result.
					// Also, performance is much higher for integer vs. float division, which is part
					// of the justification for a separate operator.
					if (right_int64 == 0) // Divide by zero produces blank result (perhaps will produce exception if script's ever support exception handlers).
					{
						this_token.marker = "";
						result_symbol = SYM_STRING;
					}
					else
						this_token.value_int64 = left_int64 / right_int64;
					break;
				case SYM_POWER:
					// Note: The function pow() in math.h adds about 28 KB of code size (uncompressed)!
					// Even assuming pow() supports negative bases such as (-2)**2, its size is why it's not used.
					// v1.0.44.11: With Laszlo's help, negative integer bases are now supported.
					if (!left_int64 && right_int64 < 0) // In essense, this is divide-by-zero.
					{
						// Return a consistent result rather than something that varies:
						this_token.marker = "";
						result_symbol = SYM_STRING;
					}
					else // We have a valid base and exponent and both are integers, so the calculation will always have a defined result.
					{
						if (left_was_negative = (left_int64 < 0))
							left_int64 = -left_int64; // Force a positive due to the limitiations of qmathPow().
						this_token.value_double = qmathPow((double)left_int64, (double)right_int64);
						if (left_was_negative && right_int64 % 2) // Negative base and odd exponent (not zero or even).
							this_token.value_double = -this_token.value_double;
						if (right_int64 < 0)
							result_symbol = SYM_FLOAT; // Due to negative exponent, override to float like TRANS_CMD_POW.
						else
							this_token.value_int64 = (__int64)this_token.value_double;
					}
					break;
				}
				this_token.symbol = result_symbol; // Must be done only after the switch() above.
			}

			else // Since one or both operands are floating point (or this is the division of two integers), the result will be floating point.
			{
				// For these two, use ATOF vs. atof so that if one of them is an integer to be converted
				// to a float for the purpose of this calculation, hex will be supported:
				switch (right.symbol)
				{
				case SYM_INTEGER: right_double = (double)right.value_int64; break;
				case SYM_FLOAT: right_double = right.value_double; break;
				default: right_double = ATOF(right_contents); // SYM_OPERAND or SYM_VAR
				// It can't be SYM_STRING because in here, both right and left are known to be numbers.
				}

				switch (left.symbol)
				{
				case SYM_INTEGER: left_double = (double)left.value_int64; break;
				case SYM_FLOAT: left_double = left.value_double; break;
				default: left_double = ATOF(left_contents); // SYM_OPERAND or SYM_VAR
				// It can't be SYM_STRING because in here, both right and left are known to be numbers.
				}

				result_symbol = IS_RELATIONAL_OPERATOR(this_token.symbol) ? SYM_INTEGER : SYM_FLOAT; // Set default. v1.0.47.01: Changed relational operators to yield integers vs. floats because it's more intuitive and traditional (might also make relational operators perform better).
				switch(this_token.symbol)
				{
				case SYM_ADD:      this_token.value_double = left_double + right_double; break;
				case SYM_SUBTRACT: this_token.value_double = left_double - right_double; break;
				case SYM_MULTIPLY: this_token.value_double = left_double * right_double; break;
				case SYM_DIVIDE:
				case SYM_FLOORDIVIDE:
					if (right_double == 0.0) // Divide by zero produces blank result (perhaps will produce exception if script's ever support exception handlers).
					{
						this_token.marker = "";
						result_symbol = SYM_STRING;
					}
					else
					{
						this_token.value_double = left_double / right_double;
						if (this_token.symbol == SYM_FLOORDIVIDE) // Like Python, the result is floor()'d, moving to the nearest integer to the left on the number line.
							this_token.value_double = qmathFloor(this_token.value_double); // Result is always a double when at least one of the inputs was a double.
					}
					break;
				case SYM_EQUALCASE: // Same behavior as SYM_EQUAL for numeric operands.
				case SYM_EQUAL:    this_token.value_int64 = left_double == right_double; break;
				case SYM_NOTEQUAL: this_token.value_int64 = left_double != right_double; break;
				case SYM_GT:       this_token.value_int64 = left_double > right_double; break;
				case SYM_LT:       this_token.value_int64 = left_double < right_double; break;
				case SYM_GTOE:     this_token.value_int64 = left_double >= right_double; break;
				case SYM_LTOE:     this_token.value_int64 = left_double <= right_double; break;
				case SYM_POWER:
					// v1.0.44.11: With Laszlo's help, negative bases are now supported as long as the exponent is not fractional.
					// See the other SYM_POWER higher above for more details about below.
					left_was_negative = (left_double < 0);
					if (left_double == 0.0 && right_double < 0  // In essense, this is divide-by-zero.
						|| left_was_negative && qmathFmod(right_double, 1.0) != 0.0) // Negative base, but exponent isn't close enough to being an integer: unsupported (to simplify code).
					{
						this_token.marker = "";
						result_symbol = SYM_STRING;
					}
					else
					{
						if (left_was_negative)
							left_double = -left_double; // Force a positive due to the limitiations of qmathPow().
						this_token.value_double = qmathPow(left_double, right_double);
						if (left_was_negative && qmathFabs(qmathFmod(right_double, 2.0)) == 1.0) // Negative base and exactly-odd exponent (otherwise, it can only be zero or even because if not it would have returned higher above).
							this_token.value_double = -this_token.value_double;
					}
					break;
				} // switch(this_token.symbol)
				this_token.symbol = result_symbol; // Must be done only after the switch() above.
			} // Result is floating point.
		} // switch() operator type

		if (sym_assign_var) // Added in v1.0.46. There are some places higher above that handle sym_assign_var themselves and skip this section via goto.
		{
			sym_assign_var->Assign(this_token); // Assign the result (based on its type) to the target variable.
			if (sym_assign_var->Type() != VAR_CLIPBOARD)
			{
				this_token.var = sym_assign_var;    // Make the result a variable rather than a normal operand so that its
				this_token.symbol = SYM_VAR;        // address can be taken, and it can be passed ByRef. e.g. &(x+=1)
			}
			//else its the clipboard, so just push this_token as-is because after its assignment is done,
			// VAR_CLIPBOARD should no longer be a SYM_VAR.  This is done to simplify the code, such as BIFs.
			//
			// Now fall through and push this_token onto the stack as an operand for use by future operators.
			// This is because by convention, an assignment like "x+=1" produces a usable operand.
		}
		goto push_this_token;
double_deref_fail:
		for (; deref->marker; ++deref);  // A deref with a NULL marker terminates the list, and also indicates whether this is a dynamic function call. deref has been set by the caller, and may or may not be the NULL marker deref.
		// Lexikos: Abort expression evaluation if a dynamic function reference fails to resolve.
		if (deref->is_function)
			goto abnormal_end;
		// FALL THROUGH TO push_this_token.
push_this_token:
		if (!this_token.circuit_token) // It's not capable of short-circuit.
			STACK_PUSH(&this_token);   // Push the result onto the stack for use as an operand by a future operator.
		else // This is the final result of an IFF's condition or a AND or OR's left branch.  Apply short-circuit boolean method to it.
		{
non_null_circuit_token:
			// Cast this left-branch result to true/false, then determine whether it should cause its
			// parent AND/OR/IFF to short-circuit.

			// If its a function result or raw numeric literal such as "if (123 or false)", its type might
			// still be SYM_OPERAND, so resolve that to distinguish between the any SYM_STRING "0"
			// (considered "true") and something that is allowed to be the number zero (which is
			// considered "false").  In other words, the only literal string (or operand made a
			// SYM_STRING via a previous operation) that is considered "false" is the empty string
			// (i.e. "0" doesn't qualify but 0 does):
			switch(this_token.symbol)
			{
			case SYM_VAR:
				// "right" vs. "left" is used even though this is technically the left branch because
				// right is used more often (for unary operators) and sometimes the compiler generates
				// faster code for the most frequently accessed variables.
				right_contents = this_token.var->Contents();
				right_is_number = IsPureNumeric(right_contents, true, false, true);
				break;
			case SYM_OPERAND:
				right_contents = this_token.marker;
				right_is_number = IsPureNumeric(right_contents, true, false, true);
				break;
			case SYM_STRING:
				right_contents = this_token.marker;
				right_is_number = PURE_NOT_NUMERIC;
			default:
				// right_contents is left uninitialized for performance and to catch bugs.
				right_is_number = this_token.symbol;
			}

			switch (right_is_number)
			{
			case PURE_INTEGER: // Probably the most common, e.g. both sides of "if (x>3 and x<6)" are the number 1/0.
				// Force it to be purely 1 or 0 if it isn't already.
				left_branch_is_true = (this_token.symbol == SYM_INTEGER ? this_token.value_int64
					: ATOI64(right_contents)) != 0;
				break;
			case PURE_FLOAT: // Convert to float, not int, so that a number between 0.0001 and 0.9999 is is considered "true".
				left_branch_is_true = (this_token.symbol == SYM_FLOAT ? this_token.value_double
					: atof(right_contents)) != 0.0;
				break;
			default:  // string.
				// Since "if x" evaluates to false when x is blank, it seems best to also have blank
				// strings resolve to false when used in more complex ways. In other words "if x or y"
				// should be false if both x and y are blank.  Logical-not also follows this convention.
				left_branch_is_true = (*right_contents != '\0');
			}

			if (this_token.circuit_token->symbol == SYM_IFF_THEN)
			{
				if (!left_branch_is_true) // The ternary's condition is false.
				{
					// Discard the entire "then" branch of this ternary operator, leaving only the
					// "else" branch to be evaluated later as the result.
					// Ternaries nested inside each other don't need to be considered special for the purpose
					// of discarding ternary branches due to the very nature of postfix (i.e. it's already put
					// nesting in the right postfix order to support this method of discarding a branch).
					for (++i; postfix[i] != this_token.circuit_token; ++i); // Should always be found, so no need to check postfix_count.
					// The outer loop will now do a ++i to discard the SYM_IFF_THEN itself.
				}
				//else the ternary's condition is true.  Do nothing; just let the next iteration evaluate the
				// THEN portion and then treat the SYM_IFF_THEN it encounters as a unary operator (after that,
				// it will discard the ELSE branch).
				continue;
			}
			// Since above didn't "continue", this_token is the left branch of an AND/OR.  Check for short-circuit.
			// The following loop exists to support cascading short-circuiting such as the following example:
			// 2>3 and 2>3 and 2>3
			// In postfix notation, the above looks like:
			// 2 3 > 2 3 > and 2 3 > and
			// When the first '>' operator is evaluated to false, it sees that its parent is an AND and
			// thus it short-circuits, discarding everything between the first '>' and the "and".
			// But since the first and's parent is the second "and", that false result just produced is now
			// the left branch of the second "and", so the loop conducts a second iteration to discard
			// everything between the first "and" and the second.  By contrast, if the second "and" were
			// an "or", the second iteration would never occur because the loop's condition would be false
			// on the second iteration, which would then cause the first and's false value to be discarded
			// (due to the loop ending without having PUSHed) because solely the right side of the "or" should
			// determine the final result of the "or".
			//
			// The following code is probably equivalent to the loop below it.  However, it's only slightly
			// smaller in code size when you examine what it actually does, and it almost certainly performs
			// slightly worse because the "goto" incurs unnecessary steps such as recalculating left_branch_is_true.
			// Therefore, it doesn't seem worth changing it:
			//if (left_branch_is_true == (this_token.circuit_token->symbol == SYM_OR)) // If true, this AND/OR causes a short-circuit
			//{
			//	for (++i; postfix[i] != this_token.circuit_token; ++i); // Should always be found, so no need to check postfix_count.
			//	this_token.symbol = SYM_INTEGER;
			//	this_token.value_int64 = left_branch_is_true; // Assign a pure 1 (for SYM_OR) or 0 (for SYM_AND).
			//	this_token.circuit_token = postfix[i]->circuit_token; // In case circuit_token == SYM_IFF_THEN.
			//	goto push_this_token; // In lieu of STACK_PUSH(this_token) in case circuit_token == SYM_IFF_THEN.
			//}
			for (circuit_token = this_token.circuit_token
				; left_branch_is_true == (circuit_token->symbol == SYM_OR);) // If true, this AND/OR causes a short-circuit
			{
				// Discard the entire right branch of this AND/OR:
				for (++i; postfix[i] != circuit_token; ++i); // Should always be found, so no need to check postfix_count.
				// Above loop is self-contained.
				if (   !(circuit_token = postfix[i]->circuit_token) // This value is also used by our loop's condition. Relies on short-circuit boolean order with the below.
					|| circuit_token->symbol == SYM_IFF_THEN   ) // Don't cascade from AND/OR into IFF because IFF requires a different cascade approach that's implemented only after its winning branch is evaluated.  Otherwise, things like "0 and 1 ? 3 : 4" wouldn't work.
				{
					// No more cascading is needed because this AND/OR isn't the left branch of another.
					// This will be the final result of this AND/OR because it's right branch was discarded
					// above without having been evaluated nor any of its functions called.  It's safe to use
					// this_token vs. postfix[i] below, for performance, because the value in its circuit_token
					// member no longer matters:
					this_token.symbol = SYM_INTEGER;
					this_token.value_int64 = left_branch_is_true; // Assign a pure 1 (for SYM_OR) or 0 (for SYM_AND).
					this_token.circuit_token = circuit_token; // In case circuit_token == SYM_IFF_THEN.
					goto push_this_token; // In lieu of STACK_PUSH(this_token) in case circuit_token == SYM_IFF_THEN.
				}
				//else there is more cascading to be checked, so continue looping.
			}
			// If the loop ends normally (not via "break"), postfix[i] is now the left branch of an
			// AND/OR that should not short-circuit.  As a result, this left branch is simply discarded
			// (by means of the outer loop's ++i) because its right branch will be the sole determination
			// of whether this AND/OR is true or false.
		} // Short-circuit (left branch of an AND/OR).
	} // For each item in the postfix array.

	// Although ACT_EXPRESSION was already checked higher above for function calls, there are other ways besides
	// an isolated function call to have ACT_EXPRESSION.  For example: var&=3 (where &= is an operator that lacks
	// a corresponding command).  Another example: true ? fn1() : fn2()
	// Also, there might be ways the function-call section didn't return for ACT_EXPRESSION, such as when somehow
	// there was more than one token on the stack even for the final function call, or maybe other unforeseen ways.
	// It seems best to avoid any chance of looking at the result since it might be invalid due to the above
	// having taken shortcuts (since it knew the result would be discarded).
	if (mActionType == ACT_EXPRESSION)   // A stand-alone expression whose end result doesn't matter.
		goto normal_end_skip_output_var; // Can't be any output_var for this action type. Also, leave result_to_return at its default of "".

	if (stack_count != 1)  // Even for multi-statement expressions, the stack should have only one item left on it:
		goto abnormal_end; // the overall result. Examples of errors include: () ... x y ... (x + y) (x + z) ... etc. (some of these might no longer produce this issue due to auto-concat).

	ExprTokenType &result_token = *stack[0];  // For performance and convenience.  Even for multi-statement, the bottommost item on the stack is the final result so that things like var1:=1,var2:=2 work.

	// Store the result of the expression in the deref buffer for the caller.  It is stored in the current
	// format in effect via SetFormat because:
	// 1) The := operator then doesn't have to convert to int/double then back to string to put the right format into effect.
	// 2) It might add a little bit of flexibility in places parameters where floating point values are expected
	//    (i.e. it allows a way to do automatic rounding), without giving up too much.  Changing floating point
	//    precision from the default of 6 decimal places is rare anyway, so as long as this behavior is documented,
	//    it seems okay for the moment.
	if (output_var)
	{
		// v1.0.45: Take a shortcut, which in the case of SYM_STRING/OPERAND/VAR avoids one memcpy
		// (into the deref buffer).  In some cases, this also saves from having to expand the deref buffer.
		output_var->Assign(result_token);
		goto normal_end_skip_output_var; // result_to_return is left at its default of "", though its value doesn't matter as long as it isn't NULL.
	}

	result_to_return = aTarget; // Set default.
	switch (result_token.symbol)
	{
	case SYM_INTEGER:
		// SYM_INTEGER and SYM_FLOAT will fit into our deref buffer because an earlier stage has already ensured
		// that the buffer is large enough to hold at least one number.  But a string/generic might not fit if it's
		// a concatenation and/or a large string returned from a called function.
		aTarget += strlen(ITOA64(result_token.value_int64, aTarget)) + 1; // Store in hex or decimal format, as appropriate.
		// Above: +1 because that's what callers want; i.e. the position after the terminator.
		goto normal_end_skip_output_var; // output_var was already checked higher above, so no need to consider it again.
	case SYM_FLOAT:
		// In case of float formats that are too long to be supported, use snprint() to restrict the length.
		 // %f probably defaults to %0.6f.  %f can handle doubles in MSVC++.
		aTarget += snprintf(aTarget, MAX_FORMATTED_NUMBER_LENGTH + 1, g.FormatFloat, result_token.value_double) + 1; // +1 because that's what callers want; i.e. the position after the terminator.
		goto normal_end_skip_output_var; // output_var was already checked higher above, so no need to consider it again.
	case SYM_STRING:
	case SYM_OPERAND:
	case SYM_VAR: // SYM_VAR is somewhat unusual at this late a stage.
		// At this stage, we know the result has to go into our deref buffer because if a way existed to
		// avoid that, we would already have goto/returned higher above.  Also, at this stage,
		// the pending result can exist in one several places:
		// 1) Our deref buf (due to being a single-deref, a function's return value that was copied to the
		//    end of our buf because there was enough room, etc.)
		// 2) In a called function's deref buffer, namely sDerefBuf, which will be deleted by our caller
		//    shortly after we return to it.
		// 3) In an area of memory we alloc'd for lack of any better place to put it.
		if (result_token.symbol == SYM_VAR)
		{
			result = result_token.var->Contents();
            result_size = result_token.var->LengthIgnoreBinaryClip() + 1; // Ignore binary clipboard for anything other than ACT_ASSIGNEXPR (i.e. output_var!=NULL) because it's documented that except for certain features, binary clipboard variables are seen only up to the first binary zero (mostly to simplify the code).
		}
		else
		{
			result = result_token.marker;
			result_size = strlen(result) + 1;
		}

		// Notes about the macro below:
		// Space is needed for whichever of the following is greater (since only one of the following is in
		// the deref buf at any given time; i.e. they can share the space by being in it at different times):
		// 1) All the expression's literal strings/numbers and double-derefs (e.g. "Array%i%" as a string).
		//    Allowing room for this_arg.length plus a terminator seems enough for any conceivable
		//    expression, even worst-cases and malformatted syntax-error expressions. This is because
		//    every numeric literal or double-deref needs to have some kind of symbol or character
		//    between it and the next one or it would never have been recognized as a separate operand
		//    in the first place.  And the final item uses the final terminator provided via +1 below.
		// 2) Any numeric result (i.e. MAX_FORMATTED_NUMBER_LENGTH).  If the expression needs to store a
		//    string result, it will take care of expanding the deref buffer.
		#define EXPR_BUF_SIZE(raw_expr_len) (raw_expr_len < MAX_FORMATTED_NUMBER_LENGTH \
			? MAX_FORMATTED_NUMBER_LENGTH : raw_expr_len) + 1 // +1 for the overall terminator.

		// If result is the empty string or a number, it should always fit because the size estimation
		// phase has ensured that capacity_of_our_buf_portion is large enough to hold those.
		// In addition, it doesn't matter if we already used target/aTarget for things higher above
		// because anything in there we're now done with, and memmove() vs. memcpy() later below
		// will allow overlap of the final result with intermediate results already in the buffer.
		size_t capacity_of_our_buf_portion;
		capacity_of_our_buf_portion = EXPR_BUF_SIZE(mArg[aArgIndex].length) + aExtraSize; // The initial amount of size available to write our final result.
		if (result_size > capacity_of_our_buf_portion)
		{
			// Do a simple expansion of our deref buffer to handle the fact that our actual result is bigger
			// than the size estimator could have calculated (due to a concatenation or a large string returned
			// from a called function).  This performs poorly but seems justified by the fact that it typically
			// happens only in extreme cases.
			size_t new_buf_size = aDerefBufSize + (result_size - capacity_of_our_buf_portion);

			// malloc() and free() are used instead of realloc() because in many cases, the overhead of
			// realloc()'s internal memcpy(entire contents) can be avoided because only part or
			// none of the contents needs to be copied (realloc's ability to do an in-place resize might
			// be unlikely for anything other than small blocks; see compiler's realloc.c):
			char *new_buf;
			if (   !(new_buf = (char *)malloc(new_buf_size))   )
			{
				LineError(ERR_OUTOFMEM ERR_ABORT);
				goto abort;
			}
			if (new_buf_size > LARGE_DEREF_BUF_SIZE)
				++sLargeDerefBufs; // And if the old deref buf was larger too, this value is decremented later below.

			// Copy only that portion of the old buffer that is in front of our portion of the buffer
			// because we no longer need our portion (except for result.marker if it happens to be
			// in the old buffer, but that is handled after this):
			size_t aTarget_offset = aTarget - aDerefBuf;
			if (aTarget_offset) // aDerefBuf has contents that must be preserved.
				memcpy(new_buf, aDerefBuf, aTarget_offset); // This will also copy the empty string if the buffer first and only character is that.
			aTarget = new_buf + aTarget_offset;
			result_to_return = aTarget; // Update to reflect new value above.
			// NOTE: result_token.marker might extend too far to the right in our deref buffer and thus be
			// larger than capacity_of_our_buf_portion because other arg(s) exist in this line after ours
			// that will be using a larger total portion of the buffer than ours.  Thus, the following must be
			// done prior to free(), but memcpy() vs. memmove() is safe in any case:
			memcpy(aTarget, result, result_size); // Copy from old location to the newly allocated one.

			free(aDerefBuf); // Free our original buffer since it's contents are no longer needed.
			if (aDerefBufSize > LARGE_DEREF_BUF_SIZE)
				--sLargeDerefBufs;

			// Now that the buffer has been enlarged, need to adjust any other pointers that pointed into
			// the old buffer:
			char *aDerefBuf_end = aDerefBuf + aDerefBufSize; // Point it to the character after the end of the old buf.
			for (i = 0; i < aArgIndex; ++i) // Adjust each item beneath ours (if any). Our own is not adjusted because we'll be returning the right address to our caller.
				if (aArgDeref[i] >= aDerefBuf && aArgDeref[i] < aDerefBuf_end)
					aArgDeref[i] = new_buf + (aArgDeref[i] - aDerefBuf); // Set for our caller.
			// The following isn't done because target isn't used anymore at this late a stage:
			//target = new_buf + (target - aDerefBuf);
			aDerefBuf = new_buf; // Must be the last step, since the old address is used above.  Set for our caller.
			aDerefBufSize = new_buf_size; // Set for our caller.
		}
		else // Deref buf is already large enough to fit the string.
			if (aTarget != result) // Currently, might be always true.
				memmove(aTarget, result, result_size); // memmove() vs. memcpy() in this case, since source and dest might overlap (i.e. "target" may have been used to put temporary things into aTarget, but those things are no longer needed and now safe to overwrite).
		aTarget += result_size;
		goto normal_end_skip_output_var; // output_var was already checked higher above, so no need to consider it again.

	default: // Result contains a non-operand symbol such as an operator.
		goto abnormal_end;
	} // switch (result_token.symbol)

// ALL PATHS ABOVE SHOULD "GOTO".  TO CATCH BUGS, ANY THAT DON'T FALL INTO "ABORT" BELOW.
abort:
	// The callers of this function know that the value of aResult (which contains the reason
	// for early exit) should be considered valid/meaningful only if result_to_return is NULL.
	result_to_return = NULL; // Use NULL to inform our caller that this entire thread is to be terminated.
	aResult = FAIL; // Indicate reason to caller.
	goto normal_end_skip_output_var; // output_var is skipped as part of standard abort behavior.

abnormal_end: // Currently the same as normal_end; it's separate to improve readability.  When this happens, result_to_return is typically "" (unless the caller overrode that default).
//normal_end: // This isn't currently used, but is available for future-use and readability.
	// v1.0.45: ACT_ASSIGNEXPR relies on us to set the output_var (i.e. whenever it's ARG1's is_expression==true).
	// Our taking charge of output_var allows certain performance optimizations in other parts of this function,
	// such as avoiding excess memcpy's and malloc's during intermediate stages.
	if (output_var && result_to_return) // i.e. don't assign if NULL to preserve backward compatibility with scripts that rely on the old value being changed in cases where an expression fails (unlikely).
		output_var->Assign(result_to_return);

normal_end_skip_output_var:
	for (i = mem_count; i--;) // Free any temporary memory blocks that were used.  Using reverse order might reduce memory fragmentation a little (depending on implementation of malloc).
		free(mem[i]);

	return result_to_return;
}



ResultType Line::ExpandArgs(VarSizeType aSpaceNeeded, Var *aArgVar[])
// Caller should either provide both or omit both of the parameters.  If provided, it means
// caller already called GetExpandedArgSize for us.
// Returns OK, FAIL, or EARLY_EXIT.  EARLY_EXIT occurs when a function-call inside an expression
// used the EXIT command to terminate the thread.
{
	// The counterparts of sArgDeref and sArgVar kept on our stack to protect them from recursion caused by
	// the calling of functions in the script:
	char *arg_deref[MAX_ARGS];
	Var *arg_var[MAX_ARGS];
	int i;

	// Make two passes through this line's arg list.  This is done because the performance of
	// realloc() is worse than doing a free() and malloc() because the former often does a memcpy()
	// in addition to the latter's steps.  In addition, realloc() as much as doubles the memory
	// load on the system during the brief time that both the old and the new blocks of memory exist.
	// First pass: determine how much space will be needed to do all the args and allocate
	// more memory if needed.  Second pass: dereference the args into the buffer.

	// First pass. It takes into account the same things as 2nd pass.
	size_t space_needed;
	if (aSpaceNeeded == VARSIZE_ERROR)
	{
		space_needed = GetExpandedArgSize(true, arg_var);
		if (space_needed == VARSIZE_ERROR)
			return FAIL;  // It will have already displayed the error.
	}
	else // Caller already determined it.
	{
		space_needed = aSpaceNeeded;
		for (i = 0; i < mArgc; ++i) // Copying only the actual/used elements is probably faster than using memcpy to copy both entire arrays.
			arg_var[i] = aArgVar[i]; // Init to values determined by caller, which helps performance if any of the args are dynamic variables.
	}

	if (space_needed > g_MaxVarCapacity)
		// Dereferencing the variables in this line's parameters would exceed the allowed size of the temp buffer:
		return LineError(ERR_MEM_LIMIT_REACHED);

	// Only allocate the buf at the last possible moment,
	// when it's sure the buffer will be used (improves performance when only a short
	// script with no derefs is being run):
	if (space_needed > sDerefBufSize)
	{
		// KNOWN LIMITATION: The memory utilization of *recursive* user-defined functions is rather high because
		// of the size of DEREF_BUF_EXPAND_INCREMENT, which is used to create a new deref buffer for each
		// layer of recursion.  So if a UDF recurses deeply, say 100 layers, about 1600 MB (16KB*100) of
		// memory would be temporarily allocated, which in a worst-case scenario would cause swapping and
		// kill performance.  Perhaps the best solution to this is to dynamically change the size of
		// DEREF_BUF_EXPAND_INCREMENT (via a new global variable) in the expression evaluation section that
		// detects that a UDF has another instance of itself on the call stack.  To ensure proper collapse-back
		// out of nested udfs and threads, the old value should be backed up, the new smaller increment set,
		// then the old size should be passed to FreeAndRestoreFunctionVars() so that it can restore it.
		// However, given the rarity of deep recursion, this doesn't seem worth the extra code size and loss of
		// performance.
		size_t increments_needed = space_needed / DEREF_BUF_EXPAND_INCREMENT;
		if (space_needed % DEREF_BUF_EXPAND_INCREMENT)  // Need one more if above division truncated it.
			++increments_needed;
		size_t new_buf_size = increments_needed * DEREF_BUF_EXPAND_INCREMENT;
		if (sDerefBuf)
		{
			// Do a free() and malloc(), which should be far more efficient than realloc(), especially if
			// there is a large amount of memory involved here (realloc's ability to do an in-place resize
			// might be unlikely for anything other than small blocks; see compiler's realloc.c):
			free(sDerefBuf);
			if (sDerefBufSize > LARGE_DEREF_BUF_SIZE)
				--sLargeDerefBufs;
		}
		if (   !(sDerefBuf = (char *)malloc(new_buf_size))   )
		{
			// Error msg was formerly: "Ran out of memory while attempting to dereference this line's parameters."
			sDerefBufSize = 0;  // Reset so that it can make another attempt, possibly smaller, next time.
			return LineError(ERR_OUTOFMEM ERR_ABORT); // Short msg since so rare.
		}
		sDerefBufSize = new_buf_size;
		if (sDerefBufSize > LARGE_DEREF_BUF_SIZE)
			++sLargeDerefBufs;
	}

	// Always init our_buf_marker even if zero iterations, because we want to enforce
	// the fact that its prior contents become invalid once we're called.
	// It's also necessary due to the fact that all the old memory is discarded by
	// the above if more space was needed to accommodate this line.
	char *our_buf_marker = sDerefBuf;  // Prior contents of buffer will be overwritten in any case.

	// From this point forward, must not refer to sDerefBuf as our buffer since it might have been
	// given a new memory area by an expression's function-call within this line.  In other words,
	// our_buf_marker is our recursion layer's buffer, but not necessarily sDerefBuf.  To enforce
	// that, and keep responsibility here rather than in ExpandExpression(), set sDerefBuf to NULL
	// so that the zero or more calls to ExpandExpression() made in the loop below (each of which will
	// in turn call zero or more user-defined functions) will allocate and use a single new deref
	// buffer if any of them need it (they all share a single deref buffer because each UDF-call
	// in a particular expression of the current line creates a buf only if necessary, and it won't
	// be necessary if some prior UDF of this same expression or line already created a deref buffer
	// "above" ours because our layer here is the only one who ever frees that upper/extra buffer).
	// Note that it is not possible for a new quasi-thread to directly interrupt ExpandArgs() because
	// ExpandArgs() never calls MsgSleep().  Therefore, each ExpandArgs() layer on the call-stack
	// is safe from interrupting threads overwriting its deref buffer.  It's true that a call to a
	// script function will usually result in MsgSleep(), and thus allow interruptions, but those
	// interruptions would hit some other deref buffer, not that of our layer.
	char *our_deref_buf = sDerefBuf; // For detecting whether ExpandExpression() caused a new buffer to be created.
	size_t our_deref_buf_size = sDerefBufSize;
	SET_S_DEREF_BUF(NULL, 0);

	ResultType result, result_to_return = OK;  // Set default return value.
	Var *the_only_var_of_this_arg;

	if (!mArgc)            // v1.0.45: Required by some commands that can have zero parameters (such as Random and
		sArgVar[0] = NULL; // PixelSearch), even if it's just to allow their output-var(s) to be omitted.  This allows OUTPUT_VAR to be used without any need to check mArgC.
	else
	{
		size_t extra_size = our_deref_buf_size - space_needed;
		for (i = 0; i < mArgc; ++i) // Second pass.  For each arg:
		{
			ArgStruct &this_arg = mArg[i]; // For performance and convenience.

			// Load-time routines have already ensured that an arg can be an expression only if
			// it's not an input or output var.
			if (this_arg.is_expression)
			{
				// v1.0.45:
				// Make ARGVAR1 (OUTPUT_VAR) temporarily valid (the entire array is made valid only later, near the
				// bottom of this function).  This helps the performance of ACT_ASSIGNEXPR by avoiding the need
				// resolve a dynamic output variable like "Array%i% := (Expr)" twice: once in GetExpandedArgSize
				// and again in ExpandExpression()).
				*sArgVar = *arg_var; // Shouldn't need to be backed up or restored because no one beneath us on the call stack should be using it; only things that go on top of us might overwrite it, so ExpandExpr() must be sure to copy this out before it launches any script-functions.
				// In addition to producing its return value, ExpandExpression() will alter our_buf_marker
				// to point to the place in our_deref_buf where the next arg should be written.
				// In addition, in some cases it will alter some of the other parameters that are arrays or
				// that are passed by-ref.  Finally, it might temporarily use parts of the buffer beyond
				// extra_size plus what the size estimator provided for it, so we should be sure here that
				// everything in our_deref_buf to the right of our_buf_marker is available to it as temporary memory.
				// Note: It doesn't seem worthwhile to enhance ExpandExpression to give us back a variable
				// for use in arg_var[] (for performance) because only rarely does an expression yield
				// a variable other than some function's local variable (and a local's contents are no
				// longer valid due to having been freed after the call [unless it's static]).
				arg_deref[i] = ExpandExpression(i, result, our_buf_marker, our_deref_buf
					, our_deref_buf_size, arg_deref, extra_size);
				extra_size = 0; // See comment below.
				// v1.0.46.01: The whole point of passing extra_size is to allow an expression to write
				// a large string to the deref buffer without having to expand it (i.e. if there happens to
				// be extra room in it that won't be used by ANY arg, including ones after THIS expression).
				// Since the expression just called above might have used some/all of the extra size,
				// the line above prevents subsequent expressions in this line from getting any extra size.
				// It's pretty rare to have more than one expression in a line anyway, and even when there
				// is there's hardly ever a need for the extra_size.  As an alternative to setting it to
				// zero, above could check how much the expression wrote to the buffer (by comparing our_buf_marker
				// before and after the call above), and compare that to how much space was reserved for this
				// particular arg/expression (which is currently a standard formula for expressions).
				if (!arg_deref[i])
				{
					// A script-function-call inside the expression returned EARLY_EXIT or FAIL.  Report "result"
					// to our caller (otherwise, the contents of "result" should be ignored since they're undefined).
					result_to_return = result;
					goto end;
				}
				continue;
			}

			if (this_arg.type == ARG_TYPE_OUTPUT_VAR)  // Don't bother wasting the mem to deref output var.
			{
				// In case its "dereferenced" contents are ever directly examined, set it to be
				// the empty string.  This also allows the ARG to be passed a dummy param, which
				// makes things more convenient and maintainable in other places:
				arg_deref[i] = "";
				continue;
			}

			// arg_var[i] was previously set by GetExpandedArgSize() so that we don't have to determine its
			// value again:
			if (   !(the_only_var_of_this_arg = arg_var[i])   ) // Arg isn't an input var or singled isolated deref.
			{
				#define NO_DEREF (!ArgHasDeref(i + 1))
				if (NO_DEREF)
				{
					arg_deref[i] = this_arg.text;  // Point the dereferenced arg to the arg text itself.
					continue;  // Don't need to use the deref buffer in this case.
				}
			}

			// Check the value of the_only_var_of_this_arg again in case the above changed it:
			if (the_only_var_of_this_arg) // This arg resolves to only a single, naked var.
			{
				switch(ArgMustBeDereferenced(the_only_var_of_this_arg, i, arg_var)) // Yes, it was called by GetExpandedArgSize() too, but a review shows it's difficult to avoid this without being worse than the disease (10/22/2006).
				{
				case CONDITION_FALSE:
					// This arg contains only a single dereference variable, and no
					// other text at all.  So rather than copy the contents into the
					// temp buffer, it's much better for performance (especially for
					// potentially huge variables like %clipboard%) to simply set
					// the pointer to be the variable itself.  However, this can only
					// be done if the var is the clipboard or a non-environment
					// normal var (since zero-length normal vars need to be fetched via
					// GetEnvironmentVariable() when g_NoEnv==false).
					// Update: Changed it so that it will deref the clipboard if it contains only
					// files and no text, so that the files will be transcribed into the deref buffer.
					// This is because the clipboard object needs a memory area into which to write
					// the filespecs it translated:
					arg_deref[i] = the_only_var_of_this_arg->Contents();
					break;
				case CONDITION_TRUE:
					// the_only_var_of_this_arg is either a reserved var or a normal var of that is also
					// an environment var (for which GetEnvironmentVariable() is called for), or is used
					// again in this line as an output variable.  In all these cases, it must
					// be expanded into the buffer rather than accessed directly:
					arg_deref[i] = our_buf_marker; // Point it to its location in the buffer.
					our_buf_marker += the_only_var_of_this_arg->Get(our_buf_marker) + 1; // +1 for terminator.
					break;
				default: // FAIL should be the only other possibility.
					result_to_return = FAIL; // ArgMustBeDereferenced() will already have displayed the error.
					goto end;
				}
			}
			else // The arg must be expanded in the normal, lower-performance way.
			{
				arg_deref[i] = our_buf_marker; // Point it to its location in the buffer.
				if (   !(our_buf_marker = ExpandArg(our_buf_marker, i))   ) // Expand the arg into that location.
				{
					result_to_return = FAIL; // ExpandArg() will have already displayed the error.
					goto end;
				}
			}
		} // for each arg.

		// IT'S NOT SAFE to do the following until the above loop FULLY completes because any calls made above to
		// ExpandExpression() might call functions, which in turn might result in a recursive call to ExpandArgs(),
		// which in turn might change the values in the static arrays sArgDeref and sArgVar.
		// Also, only when the loop ends normally is the following needed, since otherwise it's a failure condition.
		// Now that any recursive calls to ExpandArgs() above us on the stack have collapsed back to us, it's
		// safe to set the args of this command for use by our caller, to whom we're about to return.
		for (i = 0; i < mArgc; ++i) // Copying actual/used elements is probably faster than using memcpy to copy both entire arrays.
		{
			sArgDeref[i] = arg_deref[i];
			sArgVar[i] = arg_var[i];
		}
	} // mArgc > 0

	// v1.0.40.02: The following loop was added to avoid the need for the ARGn macros to provide an empty
	// string when mArgc was too small (indicating that the parameter is absent).  This saves quite a bit
	// of code size.  Also, the slight performance loss caused by it is partially made up for by the fact
	// that all the other sections don't need to check mArgc anymore.
	// Benchmarks show that it doesn't help performance to try to tweak this with a pre-check such as
	// "if (mArgc < max_params)":
	int max_params = g_act[mActionType].MaxParams; // Resolve once for performance.
	for (i = mArgc; i < max_params; ++i) // For performance, this only does the actual max args for THIS command, not MAX_ARGS.
		sArgDeref[i] = "";
		// But sArgVar isn't done (since it's more rarely used) except sArgVar[0] = NULL higher above.
		// Therefore, users of sArgVar must check mArgC if they have any doubt how many args are present in
		// the script line (this is now enforced via macros).

	// When the main/large loop above ends normally, it falls into the label below and uses the original/default
	// value of "result_to_return".

end:
	// As of v1.0.31, there can be multiple deref buffers simultaneously if one or more called functions
	// requires a deref buffer of its own (separate from ours).  In addition, if a called function is
	// interrupted by a new thread before it finishes, the interrupting thread will also use the
	// new/separate deref buffer.  To minimize the amount of memory used in such cases cases,
	// each line containing one or more expression with one or more function call (rather than each
	// function call) will get up to one deref buffer of its own (i.e. only if its function body contains
	// commands that actually require a second deref buffer).  This is achieved by saving sDerefBuf's
	// pointer and setting sDerefBuf to NULL, which effectively makes the original deref buffer private
	// until the line that contains the function-calling expressions finishes completely.
	// Description of recursion and usage of multiple deref buffers:
	// 1) ExpandArgs() receives a line with one or more expressions containing one or more calls to user functions.
	// 2) Worst-case: those function-calls create a new sDerefBuf automatically via us having set sDerefBuf to NULL.
	// 3) Even worse, the bodies of those functions call other functions, which ExpandArgs() receives, resulting in
	//    a recursive leap back to step #1.
	// So the above shows how any number of new deref buffers can be created.  But that's okay as long as the
	// recursion collapses in an orderly manner (or the program exits, in which case the OS frees all its memory
	// automatically).  This is because prior to returning, each recursion layer properly frees any extra deref
	// buffer it was responsible for creating.  It only has to free at most one such buffer because each layer of
	// ExpandArgs() on the call-stack can never be blamed for creating more than one extra buffer.
	if (our_deref_buf)
	{
		// Must always restore the original buffer, not keep the new one, because our caller needs
		// the arg_deref addresses, which point into the original buffer.
		if (sDerefBuf)
		{
			free(sDerefBuf);
			if (sDerefBufSize > LARGE_DEREF_BUF_SIZE)
				--sLargeDerefBufs;
		}
		SET_S_DEREF_BUF(our_deref_buf, our_deref_buf_size);
	}
	//else the original buffer is NULL, so keep any new sDerefBuf that might have been created (should
	// help avg-case performance).

	// For v1.0.31, this is no done right before returning so that any script function calls
	// made by our calls to ExpandExpression() will now be done.  There might still be layers
	// of ExpandArgs() beneath us on the call-stack, which is okay since they will keep the
	// largest of the two available deref bufs (as described above) and thus they should
	// reset the timer below right before they collapse/return.  
	// (Re)set the timer unconditionally so that it starts counting again from time zero.
	// In other words, we only want the timer to fire when the large deref buffer has been
	// unused/idle for a straight 10 seconds.  There is no danger of this timer freeing
	// the deref buffer at a critical moment because:
	// 1) The timer is reset with each call to ExpandArgs (this function);
	// 2) If our ExpandArgs() recursion layer takes a long time to finish, messages
	//    won't be checked and thus the timer can't fire because it relies on the msg loop.
	// 3) If our ExpandArgs() recursion layer launches function-calls in ExpandExpression(),
	//    those calls will call ExpandArgs() recursively and reset the timer if its
	//    buffer (not necessarily the original buffer somewhere on the call-stack) is large
	//    enough.  In light of this, there is a chance that the timer might execute and free
	//    a deref buffer other than the one it was originally intended for.  But in real world
	//    scenarios, that seems rare.  In addition, the consequences seem to be limited to
	//    some slight memory inefficiency.
	// It could be aruged that the timer should only be activated when a hypothetical static
	// var sLayersthat we maintain here indicates that we're the only layer.  However, if that
	// were done and the launch of a script function creates (directly or through thread
	// interruption, indirectly) a large deref buffer, and that thread is waiting for something
	// such as WinWait, that large deref buffer would never get freed.
	#define SET_DEREF_TIMER(aTimeoutValue) g_DerefTimerExists = SetTimer(g_hWnd, TIMER_ID_DEREF, aTimeoutValue, DerefTimeout);
	if (sDerefBufSize > LARGE_DEREF_BUF_SIZE)
		SET_DEREF_TIMER(10000) // Reset the timer right before the deref buf is possibly about to become idle.

	return result_to_return;
}

	

VarSizeType Line::GetExpandedArgSize(bool aCalcDerefBufSize, Var *aArgVar[])
// Returns the size, or VARSIZE_ERROR if there was a problem.
// This function can return a size larger than what winds up actually being needed
// (e.g. caused by ScriptGetCursor()), so our callers should be aware that that can happen.
{
	int i;
	VarSizeType space_needed;
	Var *the_only_var_of_this_arg;
	ResultType result;

	// Note: the below loop is similar to the one in ExpandArgs(), so the two should be maintained together:
	for (i = 0, space_needed = 0; i < mArgc; ++i) // FOR EACH ARG:
	{
		ArgStruct &this_arg = mArg[i]; // For performance and convenience.

		// Accumulate the total of how much space we will need.
		if (this_arg.type == ARG_TYPE_OUTPUT_VAR)  // These should never be included in the space calculation.
		{
			if (mActionType != ACT_ASSIGN) // PerformAssign() already resolved its output-var, so don't do it again here.
			{
				if (   !(aArgVar[i] = ResolveVarOfArg(i))   ) // v1.0.45: Resolve output variables too, which eliminates a ton of calls to ResolveVarOfArg() in various other functions.  This helps code size more than performance.
					return VARSIZE_ERROR;  // The above will have already displayed the error.
			}
			continue;
		}
		// Otherwise, set default aArgVar[] (above took care of setting aArgVar[] for itself).
		aArgVar[i] = NULL;

		if (this_arg.is_expression)
		{
			space_needed += EXPR_BUF_SIZE(this_arg.length); // See comments at macro definition.
			continue;
		}

		// Always do this check before attempting to traverse the list of dereferences, since
		// such an attempt would be invalid in this case:
		the_only_var_of_this_arg = NULL;
		if (this_arg.type == ARG_TYPE_INPUT_VAR) // Previous stage has ensured that arg can't be an expression if it's an input var.
			if (   !(the_only_var_of_this_arg = ResolveVarOfArg(i, false))   )
				return VARSIZE_ERROR;  // The above will have already displayed the error.

		if (!the_only_var_of_this_arg) // It's not an input var.
		{
			if (NO_DEREF)
			{
				if (!aCalcDerefBufSize) // i.e. we want the total size of what the args resolve to.
					space_needed += this_arg.length + 1;  // +1 for the zero terminator.
				// else don't increase space_needed, even by 1 for the zero terminator, because
				// the terminator isn't needed if the arg won't exist in the buffer at all.
				continue;
			}
			// Now we know it has at least one deref.  If the second deref's marker is NULL,
			// the first is the only deref in this arg.  UPDATE: The following will return
			// false for function calls since they are always followed by a set of parentheses
			// (empty or otherwise), thus they will never be seen as isolated by it:
			#define SINGLE_ISOLATED_DEREF (!this_arg.deref[1].marker\
				&& this_arg.deref[0].length == this_arg.length) // and the arg contains no literal text
			if (SINGLE_ISOLATED_DEREF) // This also ensures the deref isn't a function-call.  10/25/2006: It might be possible to avoid the need for detecting SINGLE_ISOLATED_DEREF by transforming them into INPUT_VARs at loadtime.  I almost finished such a mod but the testing and complications with things like ListLines didn't seem worth the tiny benefit.
				the_only_var_of_this_arg = this_arg.deref[0].var;
		}
		if (the_only_var_of_this_arg) // i.e. check it again in case the above block changed the value.
		{
			// This is set for our caller so that it doesn't have to call ResolveVarOfArg() again, which
			// would a performance hit if this variable is dynamically built and thus searched for at runtime:
			aArgVar[i] = the_only_var_of_this_arg; // For now, this is done regardless of whether it must be dereferenced.
			if (aCalcDerefBufSize) // In this case, caller doesn't want its size unconditionally included, but instead only if certain conditions are met.
			{
				if (   !(result = ArgMustBeDereferenced(the_only_var_of_this_arg, i, aArgVar))   )
					return VARSIZE_ERROR;
				if (result == CONDITION_FALSE)
					continue;
				//else the size of this arg is always included, so fall through to below.
			}
			//else caller wanted it's size unconditionally included, so continue on to below.
			space_needed += the_only_var_of_this_arg->Get() + 1;  // +1 for the zero terminator.
			// NOTE: Get() (with no params) can retrieve a size larger that what winds up actually
			// being needed, so our callers should be aware that that can happen.
			continue;
		}

		// Otherwise: This arg has more than one deref, or a single deref with some literal text around it.
		space_needed += this_arg.length + 1; // +1 for this arg's zero terminator in the buffer.
		if (this_arg.deref) // There's at least one deref.
		{
			for (DerefType *deref = this_arg.deref; deref->marker; ++deref)
			{
				// Replace the length of the deref's literal text with the length of its variable's contents:
				space_needed -= deref->length;
				if (!deref->is_function)
					space_needed += deref->var->Get(); // If an environment var, Get() will yield its length.
				//else it's a function-call's function name, in which case it's length is effectively zero since
				// the function name never gets copied into the deref buffer during ExpandExpression().
			}
		}
	} // For each arg.

	return space_needed;
}



ResultType Line::ArgMustBeDereferenced(Var *aVar, int aArgIndex, Var *aArgVar[]) // 10/22/2006: __forceinline didn't help enough to be worth the added code size of having two instances.
// Shouldn't be called only for args of type ARG_TYPE_OUTPUT_VAR because they never need to be dereferenced.
// aArgVar[] is used for performance; it's assumed to contain valid items only up to aArgIndex, not beyond
// (since normally output vars lie to the left of all input vars, so it doesn't seem worth doing anything
// more complicated).
// Returns CONDITION_TRUE, CONDITION_FALSE, or FAIL.
{
	if (mActionType == ACT_SORT) // See PerformSort() for why it's always dereferenced.
		return CONDITION_TRUE;
	aVar = aVar->ResolveAlias(); // Helps performance, but also necessary to accurately detect a match further below.
	if (aVar->Type() == VAR_CLIPBOARD)
		// Even if the clipboard is both an input and an output var, it still
		// doesn't need to be dereferenced into the temp buffer because the
		// clipboard has two buffers of its own.  The only exception is when
		// the clipboard has only files on it, in which case those files need
		// to be converted into plain text:
		return CLIPBOARD_CONTAINS_ONLY_FILES ? CONDITION_TRUE : CONDITION_FALSE;
	if (aVar->Type() != VAR_NORMAL || (!g_NoEnv && !aVar->Length()) || aVar == g_ErrorLevel) // v1.0.43.08: Added g_NoEnv.
		// Reserved vars must always be dereferenced due to their volatile nature.
		// When g_NoEnv==false, normal vars of length zero are dereferenced because they might exist
		// as system environment variables, whose contents are also potentially volatile (i.e. they
		// are sometimes changed by outside forces).
		// As of v1.0.25.12, g_ErrorLevel is always dereferenced also so that a command that sets ErrorLevel
		// can itself use ErrorLevel as in this example: StringReplace, EndKey, ErrorLevel, EndKey:
		return CONDITION_TRUE;

	// Before doing the below, the checks above must be done to ensure it's VAR_NORMAL.  Otherwise, things like
	// the following won't work: StringReplace, o, A_ScriptFullPath, xxx
	// v1.0.45: The following check improves performance slightly by avoiding the loop further below in cases
	// where it's known that a command either doesn't have an output_var or can tolerate the output_var's
	// contents being at the same address as that of one or more of the input-vars.  For example, the commands
	// StringRight/Left and similar can tolerate the same address because they always produce a string whose
	// length is less-than-or-equal to the input-string, thus Assign() will never need to free/realloc the
	// output-var prior to assigning the input-var's contents to it (whose contents are the same as output-var).
	if (!(g_act[mActionType].MaxParamsAu2WithHighBit & 0x80)) // Commands that have this bit don't need final check
		return CONDITION_FALSE;                               // further below (though they do need the ones above).

	// Since the above didn't return, we know that this is a NORMAL input var that isn't an
	// environment variable.  Such input vars only need to be dereferenced if they are also
	// used as an output var by the current script line:
	Var *output_var;
	for (int i = 0; i < mArgc; ++i)
		if (i != aArgIndex && mArg[i].type == ARG_TYPE_OUTPUT_VAR)
		{
			if (   !(output_var = (i < aArgIndex) ? aArgVar[i] : ResolveVarOfArg(i, false))   ) // aArgVar: See top of this function for comments.
				return FAIL;  // It will have already displayed the error.
			if (output_var->ResolveAlias() == aVar)
				return CONDITION_TRUE;
		}
	// Otherwise:
	return CONDITION_FALSE;
}



char *Line::ExpandArg(char *aBuf, int aArgIndex, Var *aArgVar) // 10/2/2006: Doesn't seem worth making it inline due to more complexity than expected.  It would also increase code size without being likely to help performance much.
// Caller must ensure that aArgVar is the variable of the aArgIndex arg when it's of type ARG_TYPE_INPUT_VAR.
// Caller must be sure not to call this for an arg that's marked as an expression, since
// expressions are handled by a different function.  Similarly, it must ensure that none
// of this arg's deref's are function-calls, i.e. that deref->is_function is always false.
// Caller must ensure that aBuf is large enough to accommodate the translation
// of the Arg.  No validation of above params is done, caller must do that.
// Returns a pointer to the char in aBuf that occurs after the zero terminator
// (because that's the position where the caller would normally resume writing
// if there are more args, since the zero terminator must normally be retained
// between args).
{
	ArgStruct &this_arg = mArg[aArgIndex]; // For performance and convenience.
#ifdef _DEBUG
	// This should never be called if the given arg is an output var, so flag that in DEBUG mode:
	if (this_arg.type == ARG_TYPE_OUTPUT_VAR)
	{
		LineError("DEBUG: ExpandArg() was called to expand an arg that contains only an output variable.");
		return NULL;
	}
#endif

	if (aArgVar)
		// +1 so that we return the position after the terminator, as required.
		return aBuf += aArgVar->Get(aBuf) + 1;

	char *this_marker, *pText = this_arg.text;  // Start at the begining of this arg's text.
	if (this_arg.deref) // There's at least one deref.
	{
		for (DerefType *deref = this_arg.deref  // Start off by looking for the first deref.
			; deref->marker; ++deref)  // A deref with a NULL marker terminates the list.
		{
			// FOR EACH DEREF IN AN ARG (if we're here, there's at least one):
			// Copy the chars that occur prior to deref->marker into the buffer:
			for (this_marker = deref->marker; pText < this_marker; *aBuf++ = *pText++); // memcpy() is typically slower for small copies like this, at least on some hardware.
			// Now copy the contents of the dereferenced var.  For all cases, aBuf has already
			// been verified to be large enough, assuming the value hasn't changed between the
			// time we were called and the time the caller calculated the space needed.
			aBuf += deref->var->Get(aBuf); // Caller has ensured that deref->is_function==false
			// Finally, jump over the dereference text. Note that in the case of an expression, there might not
			// be any percent signs within the text of the dereference, e.g. x + y, not %x% + %y%.
			pText += deref->length;
		}
	}
	// Copy any chars that occur after the final deref into the buffer:
	for (; *pText; *aBuf++ = *pText++); // memcpy() is typically slower for small copies like this, at least on some hardware.
	// Terminate the buffer, even if nothing was written into it:
	*aBuf++ = '\0';
	return aBuf; // Returns the position after the terminator.
}
