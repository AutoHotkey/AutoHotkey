﻿/*
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

//////////////////////////////////////////////////////////////////////////////////////
// v1.0.40.02: This is now a separate file to allow its compiler optimization settings
// to be set independently of those of the other modules.  In one benchmark, this
// improved performance of expressions and function calls by 9% (that is, when the
// other modules are set to "minimize size" such as for the AutoHotkeySC.bin file).
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
#include "script_object.h"
#include "globaldata.h" // for a lot of things
#include "qmath.h" // For ExpandExpression()

// __forceinline: Decided against it for this function because although it's only called by one caller,
// testing shows that it wastes stack space (room for its automatic variables would be unconditionally 
// reserved in the stack of its caller).  Also, the performance benefit of inlining this is too slight.
// Here's a simple way to verify wasted stack space in a caller that calls an inlined function:
//    DWORD stack
//    _asm mov stack, esp
//    MsgBox(stack);
LPTSTR Line::ExpandExpression(int aArgIndex, ResultType &aResult, ResultToken *aResultToken
		, LPTSTR &aTarget, LPTSTR &aDerefBuf, size_t &aDerefBufSize, LPTSTR aArgDeref[], size_t aExtraSize
		, Var **aArgVar)
// Caller should ignore aResult unless this function returns NULL.
// Returns a pointer to this expression's result, which can be one of the following:
// 1) NULL, in which case aResult will be either FAIL or EARLY_EXIT to indicate the means by which the current
//    quasi-thread was terminated as a result of a function call.
// 2) The constant empty string (""), in which case we do not alter aTarget for our caller.
// 3) Some persistent location not in aDerefBuf, namely the mContents of a variable or a literal string/number,
//    such as a function-call that returns "abc", 123, or a variable.
// 4) At position aTarget inside aDerefBuf (note that aDerefBuf might have been reallocated by us).
// aTarget is left unchanged except in case #4, in which case aTarget has been adjusted to the position after our
// result-string's terminator.  In addition, in case #4, aDerefBuf, aDerefBufSize, and aArgDeref[] have been adjusted
// for our caller if aDerefBuf was too small and needed to be enlarged.
//
// Thanks to Joost Mulders for providing the expression evaluation code upon which this function is based.
{
	LPTSTR target = aTarget; // "target" is used to track our usage (current position) within the aTarget buffer.

	// The following must be defined early so that to_free_count is initialized and the array is guaranteed to be
	// "in scope" in case of early "goto" (goto substantially boosts performance and reduces code size here).
	#define MAX_EXPR_MEM_ITEMS 200 // v1.0.47.01: Raised from 100 because a line consisting entirely of concat operators can exceed it.  However, there's probably not much point to going much above MAX_TOKENS/2 because then it would reach the MAX_TOKENS limit first.
	ExprTokenType *to_free[MAX_EXPR_MEM_ITEMS]; // No init necessary.  In many cases, it will never be used.
	int to_free_count = 0; // The actual number of items in use in the above array.
	LPTSTR result_to_return = _T(""); // By contrast, NULL is used to tell the caller to abort the current thread.  That isn't done for normal syntax errors, just critical conditions such as out-of-memory.
	Var *output_var = (mActionType == ACT_ASSIGNEXPR && aArgIndex == 1) ? *aArgVar : NULL; // Resolve early because it's similar in usage/scope to the above.  Plus MUST be resolved prior to calling any script-functions since they could change the values in sArgVar[].

	ExprTokenType *stack[MAX_TOKENS];
	int stack_count = 0;
	ExprTokenType *&postfix = mArg[aArgIndex].postfix;

	///////////////////////////////
	// EVALUATE POSTFIX EXPRESSION
	///////////////////////////////
	int i, actual_param_count, delta;
	SymbolType right_is_number, left_is_number, right_is_pure_number, left_is_pure_number, result_symbol;
	double right_double, left_double;
	__int64 right_int64, left_int64;
	LPTSTR right_string, left_string;
	size_t right_length, left_length;
	TCHAR left_buf[MAX_NUMBER_SIZE];  // BIF_OnMessage and SYM_DYNAMIC rely on this one being large enough to hold MAX_VAR_NAME_LENGTH.
	TCHAR right_buf[MAX_NUMBER_SIZE]; // Only needed for holding numbers
	LPTSTR result; // "result" is used for return values and also the final result.
	VarSizeType result_length;
	size_t result_size, alloca_usage = 0; // v1.0.45: Track amount of alloca mem to avoid stress on stack from extreme expressions (mostly theoretical).
	BOOL done, done_and_have_an_output_var, make_result_persistent, left_branch_is_true
		, left_was_negative, is_pre_op; // BOOL vs. bool benchmarks slightly faster, and is slightly smaller in code size (or maybe it's cp1's int vs. char that shrunk it).
	ExprTokenType *this_postfix, *p_postfix;
	Var *sym_assign_var, *temp_var;

	// v1.0.44.06: EXPR_SMALL_MEM_LIMIT is the means by which _alloca() is used to boost performance a
	// little by avoiding the overhead of malloc+free for small strings.  The limit should be something
	// small enough that the chance that even 10 of them would cause stack overflow is vanishingly small
	// (the program is currently compiled to allow stack to expand anyway).  Even in a worst-case
	// scenario where an expression is composed entirely of functions and they all need to use this
	// limit of stack space, there's a practical limit on how many functions you can call in an
	// expression due to MAX_TOKENS (probably around MAX_TOKENS / 3).
	#define EXPR_SMALL_MEM_LIMIT 4097 // The maximum size allowed for an item to qualify for alloca.
	#define EXPR_ALLOCA_LIMIT 40000  // The maximum amount of alloca memory for all items.  v1.0.45: An extra precaution against stack stress in extreme/theoretical cases.
	#define EXPR_IS_DONE (!stack_count && this_postfix[1].symbol == SYM_INVALID) // True if we've used up the last of the operators & operands.  Non-zero stack_count combined with SYM_INVALID would indicate an error (an exception will be thrown later, so don't take any shortcuts).

	// For each item in the postfix array: if it's an operand, push it onto stack; if it's an operator or
	// function call, evaluate it and push its result onto the stack.  SYM_INVALID is the special symbol
	// that marks the end of the postfix array.
	for (this_postfix = postfix; this_postfix->symbol != SYM_INVALID; ++this_postfix) // Using pointer vs. index (e.g. postfix[i]) reduces OBJ code size by ~122 and seems to perform at least as well.
	{
		// Set default early to simplify the code.  All struct members are needed: symbol_type (e.g. in
		// cases such as this token being an ordinary operand), circuit_token to preserve loadtime info,
		// buf for SYM_DYNAMIC, and marker (in cases such as literal strings).  Also, this_token is used
		// almost everywhere further below in preference to this_postfix because:
		// 1) The various SYM_ASSIGN_* operators (e.g. SYM_ASSIGN_CONCAT) are changed to different operators
		//    to simplify the code.  So must use the changed/new value in this_token, not the original value in
		//    this_postfix.
		// 2) Using a particular variable very frequently might help compiler to optimize that variable to
		//    generate faster code.
		ExprTokenType &this_token = *(ExprTokenType *)_alloca(sizeof(ExprTokenType)); // Saves a lot of stack space, and seems to perform just as well as something like the following (at the cost of ~82 byte increase in OBJ code size): ExprTokenType &this_token = new_token[new_token_count++]  // array size MAX_TOKENS
		this_token.CopyExprFrom(*this_postfix); // See comment section above.

		// At this stage, operands in the postfix array should be SYM_STRING, SYM_INTEGER, SYM_FLOAT or SYM_DYNAMIC.
		// But all are checked since that operation is just as fast:
		if (IS_OPERAND(this_token.symbol)) // If it's an operand, just push it onto stack for use by an operator in a future iteration.
		{
			if (this_token.symbol == SYM_DYNAMIC) // CONVERTED HERE/EARLY TO SOMETHING *OTHER* THAN SYM_DYNAMIC so that no later stages need any handling for them as operands. SYM_DYNAMIC is quite similar to SYM_FUNC/BIF in this respect.
			{
				if (SYM_DYNAMIC_IS_DOUBLE_DEREF(this_token)) // Double-deref such as Array%i%.
				{
					if (!stack_count) // Prevent stack underflow.
						goto abort_with_exception;
					ExprTokenType &right = *STACK_POP;
					right_string = TokenToString(right, right_buf, &right_length);
					// Do some basic validation to ensure a helpful error message is displayed on failure.
					if (right_length == 0)
					{
						LineError(ERR_DYNAMIC_BLANK, FAIL, mArg[aArgIndex].text);
						goto abort;
					}
					if (right_length > MAX_VAR_NAME_LENGTH)
					{
						LineError(ERR_DYNAMIC_TOO_LONG, FAIL, right_string);
						goto abort;
					}
					// In v1.0.31, FindOrAddVar() vs. FindVar() is called below to support the passing of non-existent
					// array elements ByRef, e.g. Var:=MyFunc(Array%i%) where the MyFunc function's parameter is
					// defined as ByRef, would effectively create the new element Array%i% if it doesn't already exist.
					// Since at this stage we don't know whether this particular double deref is to be sent as a param
					// to a function, or whether it will be byref, this is done unconditionally for all double derefs
					// since it seems relatively harmless to create a blank variable in something like var := Array%i%
					// (though it will produce a runtime error if the double resolves to an illegal variable name such
					// as one containing spaces).
					if (   !(temp_var = g_script.FindOrAddVar(right_string, right_length))   )
					{
						// Above already displayed the error.  As of v1.0.31, this type of error is displayed and
						// causes the current thread to terminate, which seems more useful than the old behavior
						// that tolerated anything in expressions.
						goto abort;
					}
					if (aArgVar && EXPR_IS_DONE && mArg[aArgIndex].type == ARG_TYPE_OUTPUT_VAR)
					{
						if (VAR_IS_READONLY(*temp_var))
						{
							// Having this check here allows us to display the variable name rather than its contents
							// in the error message.
							LineError(ERR_VAR_IS_READONLY, FAIL, temp_var->mName);
							goto abort;
						}
						// Take a shortcut to allow dynamic output vars to resolve to builtin vars such as Clipboard
						// or A_WorkingDir.  For additional comments, search for "SYM_VAR is somewhat unusual".
						// This also ensures that the var's content is not transferred to aResultToken, which means
						// that PerformLoopFor() is not required to check for/release an object in args 0 and 1.
						aArgVar[aArgIndex] = temp_var;
						goto normal_end_skip_output_var; // result_to_return is left at its default of "", though its value doesn't matter as long as it isn't NULL.
					}
					this_token.var = temp_var;
				}
				//else: It's a built-in variable.

				// Check if it's a normal variable rather than a built-in variable.
				switch (this_token.var->Type())
  				{
				case VAR_CLIPBOARD:
					if (!this_token.is_lvalue)
						break;
					// Otherwise, this is the target of an assignment, so must be SYM_VAR:
				case VAR_NORMAL:
					this_token.symbol = SYM_VAR; // The fact that a SYM_VAR operand is always VAR_NORMAL (with one limited exception) is relied upon in several places such as built-in functions.
					goto push_this_token;
				case VAR_VIRTUAL:
					if (this_token.is_lvalue)
					{
						this_token.symbol = SYM_VAR;
						goto push_this_token;
					}
					if (this_token.var->mVV->Get == BIV_LoopIndex) // v1.0.48.01: Improve performance of A_Index by treating it as an integer rather than a string in expressions (avoids conversions to/from strings).
					{
						this_token.SetValue(g->mLoopIteration);
						goto push_this_token;
					}
					if (this_token.var->mVV->Get == BIV_EventInfo) // v1.0.48.02: A_EventInfo is used often enough in performance-sensitive numeric contexts to seem worth special treatment like A_Index; e.g. LV_GetText(RowText, A_EventInfo) or RegisterCallback()'s A_EventInfo.
					{
						this_token.SetValue(g->EventInfo);
						goto push_this_token;
					}
					// ABOVE: Goto's and simple assignments (like the SYM_INTEGER ones above) are only a few
					// bytes in code size, so it would probably cost more than it's worth in performance
					// and code size to merge them into a code section shared by all of the above.  Although
					// each comparison "this_token.var->mBIV == BIV_xxx" is surprisingly large in OBJ size,
					// the resulting EXE does not reflect this: even 27 such comparisons and sections (all
					// to a different BIV) don't increase the uncompressed EXE size.
					//
					// OTHER CANDIDATES FOR THE ABOVE:
					// A_TickCount: Usually not performance-critical.
					// A_GuiWidth/Height: Maybe not used in expressions often enough.
					// A_GuiX/Y: Not performance-critical and/or too rare: Popup menu, DropFiles, PostMessage's coords.
					// A_Gui: Hard to say.
					// A_LastError: Seems too rare to justify the code size and loss of performance here.
					// A_Msec: Would help but it's probably rarely used; probably has poor granularity, not likely to be better than A_TickCount.
					// A_TimeIdle/Physical: These are seldom performance-critical.
					//
					// True, False, A_IsUnicode and A_PtrSize are handled in ExpressionToPostfix().
					// Since their values never change at run-time, they are replaced at load-time
					// with the appropriate SYM_INTEGER value.
					//
					break; // case VAR_VIRTUAL
				default:
					if (this_token.is_lvalue)
					{
						// Having this check here allows us to display the variable name rather than its contents
						// in the error message.
						LineError(ERR_VAR_IS_READONLY, FAIL, this_token.var->mName);
						goto abort;
					}
  				}
				// Otherwise, it's a built-in variable.
				result_size = this_token.var->Get() + 1;
				if (result_size == 1)
				{
					this_token.SetValue(_T(""), 0);
					goto push_this_token;
				}
				// Otherwise, it's a built-in variable which is not empty. Need some memory to store it.
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
					result = (LPTSTR)talloca(result_size);
					alloca_usage += result_size; // This might put alloca_usage over the limit by as much as EXPR_SMALL_MEM_LIMIT, but that is fine because it's more of a guideline than a limit.
				}
				else // Need to create some new persistent memory for our temporary use.
				{
					if (to_free_count == MAX_EXPR_MEM_ITEMS // No more slots left (should be nearly impossible).
						|| !(result = tmalloc(result_size)))
					{
						LineError(ERR_OUTOFMEM, FAIL, this_token.var->mName);
						goto abort;
					}
					to_free[to_free_count++] = &this_token;
				}
				result_length = this_token.var->Get(result);
				this_token.marker = result;  // Must be done after above because marker and var overlap in union.
				this_token.marker_length = result_length;
				this_token.symbol = SYM_STRING;
			} // if (this_token.symbol == SYM_DYNAMIC)
			goto push_this_token;
		} // if (IS_OPERAND(this_token.symbol))

		if (this_token.symbol == SYM_FUNC) // A call to a function (either built-in or defined by the script).
		{
			Func *func = this_token.deref->func;
			actual_param_count = this_token.deref->param_count; // For performance.
			if (actual_param_count > stack_count) // Prevent stack underflow (probably impossible if actual_param_count is accurate).
				goto abort_with_exception;
			// Adjust the stack early to simplify.  Above already confirmed that the following won't underflow.
			// Pop the actual number of params involved in this function-call off the stack.
			stack_count -= actual_param_count; // Now stack[stack_count] is the leftmost item in an array of function-parameters, which simplifies processing later on.
			ExprTokenType **params = stack + stack_count;

			if (!func)
			{
				// This is a dynamic function call.
				if (!stack_count) // SYM_DYNAMIC should have pushed a function name or reference onto the stack, but a syntax error may still cause this condition.
					goto abort_with_exception;
				stack_count--;
				func = TokenToFunc(*stack[stack_count]); // Supports function names and function references.
				if (!func)
				{
					// This isn't a function name or reference, but it could be an object emulating
					// a function reference.  Additionally, we want something like %emptyvar%() to
					// invoke g_MetaObject, so this part is done even if stack[stack_count] is not
					// an object.  To "call" the object/value, we need to insert an empty method
					// name between the object/value and the parameter list.  There should always
					// be room for this since the maximum number of operands at any one time <=
					// postfix token count < infix token count < MAX_TOKENS == _countof(stack).
					// That is, each extra (SYM_OPAREN, SYM_COMMA or SYM_CPAREN) token in infix
					// effectively reserves one stack slot.
					if (actual_param_count)
						memmove(params + 1, params, actual_param_count * sizeof(ExprTokenType *));
					// Insert an empty string:
					params[0] = (ExprTokenType *)_alloca(sizeof(ExprTokenType));
					params[0]->SetValue(_T("Call"), 4);
					params--; // Include the object, which is already in the right place.
					actual_param_count += 2;
					extern ExprOpFunc g_ObjCall;
					func = &g_ObjCall;
				}
				// Above has set func to a non-NULL value, but still need to verify there are enough params.
				// Although passing too many parameters is useful (due to the limitations of variadic calls),
				// passing too few parameters (and treating the missing ones as optional) seems a little
				// inappropriate because it would allow the function's caller to second-guess the function's
				// designer (the designer could provide a default value if a parameter is capable of being
				// omitted). Another issue might be misbehavior by built-in functions that assume that the
				// minimum number of parameters are present due to prior validation.  So either all the
				// built-in functions would have to be reviewed, or the minimum would have to be enforced
				// for them but not user-defined functions, which is inconsistent.  Finally, allowing too-
				// few parameters seems like it would reduce the ability to detect script bugs at runtime.
				// Param count is now checked in Func::Call(), so doesn't need to be checked here.
			}
			
			// The following two steps are done for built-in functions inside Func::Call:
			//result_token.symbol = SYM_INTEGER; // Set default return type so that functions don't have to do it if they return INTs.
			//result_token.func = func;          // Inform function of which built-in function called it (allows code sharing/reduction).
			
			// This is done by ResultToken below:
			//result_token.buf = left_buf;       // mBIF() can use this to store a string result, and for other purposes.
			//result_token.mem_to_free = NULL;   // Init to detect whether the called function allocates it.
			
			ResultToken result_token;
			result_token.InitResult(left_buf); // But we'll take charge of its contents INSTEAD of calling Free().

			// Call the user-defined or built-in function.
			if (!func->Call(result_token, params, actual_param_count, this_token.deref->type == DT_VARIADIC))
			{
				// Func::Call returning false indicates an EARLY_EXIT or FAIL result, meaning that the
				// thread should exit or transfer control to a Catch statement.  Abort the remainder
				// of this expression and pass the result back to our caller:
				aResult = result_token.Result();
				result_to_return = NULL; // Use NULL to inform our caller that this thread is finished (whether through normal means such as Exit or a critical error).
				// Above: The callers of this function know that the value of aResult (which already contains the
				// reason for early exit) should be considered valid/meaningful only if result_to_return is NULL.
				goto normal_end_skip_output_var; // output_var is left unchanged in these cases.
			}

#ifdef CONFIG_DEBUGGER
			// See PostExecFunctionCall() itself for comments.
			if (g_Debugger.IsConnected())
				g_Debugger.PostExecFunctionCall(this);
#endif
			g_script.mCurrLine = this; // For error-reporting.

			if (result_token.symbol != SYM_STRING)
			{
				// No need for make_result_persistent or early Assign().  Any numeric or object result can
				// be considered final because it's already stored in permanent memory (the token itself).
				// Additionally, this_token.mem_to_free is assumed to be NULL since the result is not
				// a string; i.e. the function would've had no need to return memory to us.
				this_token.value_int64 = result_token.value_int64;
				this_token.symbol = result_token.symbol;
				if (this_token.symbol == SYM_OBJECT)
				{
					if (to_free_count == MAX_EXPR_MEM_ITEMS) // No more slots left (should be nearly impossible).
					{
						this_token.object->Release();
						LineError(ERR_OUTOFMEM);
						goto abort;
					}
					to_free[to_free_count++] = &this_token;
				}
				goto push_this_token;
			}
			
			done = EXPR_IS_DONE;

			// v1.0.45: If possible, take a shortcut for performance.  Doing it this way saves at least
			// two memcpy's (one into deref buffer and then another back into the output_var by
			// ACT_ASSIGNEXPR itself).  In some cases is also saves from having to expand the deref
			// buffer as well as the output_var (since it's current memory might be too small to hold
			// the new memory block). Thus we give it a new block directly to avoid all of that.
			// This should be a big boost to performance when long strings are involved.
			Var *internal_output_var;
			if (done) // i.e. we've now produced the final result.
			{
				if (mActionType == ACT_EXPRESSION) // Isolated expression: Outermost function call's result will be ignored, so no need to store it.
				{
					result_token.Free(); // Since we're not taking charge of it in this case.
					goto normal_end_skip_output_var; // No output_var is possible for ACT_EXPRESSION.
				}
				internal_output_var = output_var; // NULL unless this is ACT_ASSIGNEXPR.
			}
			// It's fairly rare that the following optimization is even applicable because it requires
			// an assignment *internal* to an expression, such as "if not var:=func()", or "a:=b, c:=func()".
			// But it seems best to optimize these cases so that commas aren't penalized.
			else if (this_postfix[1].symbol == SYM_ASSIGN  // Next operation is ":=".
					&& stack_count && stack[stack_count-1]->symbol == SYM_VAR // i.e. let the next iteration handle errors instead of doing it here.  Further below relies on this having been checked.
					&& stack[stack_count-1]->var->Type() == VAR_NORMAL) // Don't do clipboard here because: 1) AcceptNewMem() doesn't support it; 2) Could probably use Assign() and then make its result be a newly added mem_count item, but the code complexity doesn't seem worth it given the rarity.
				internal_output_var = stack[stack_count-1]->var;
			else
				internal_output_var = NULL;
			
			// RELIES ON THE SYM_STRING CHECK above having been done first.
			result        = result_token.marker;
			result_length = result_token.marker_length;
			if (result_length == -1)
				result_length = (VarSizeType)_tcslen(result);

			if (internal_output_var)
			{
				// Check if the called function allocated some memory for its result and turned it over to us.
				// In most cases, the string stored in mem_to_free (if it has been set) is the same address as
				// this_token.marker (i.e. what is named "result" further below), because that's what the
				// built-in functions are normally using the memory for.
				if (result_token.mem_to_free)
				{
					ASSERT(result_token.mem_to_free == result); // See similar line below for comments.
					// So now, turn over responsibility for this memory to the variable. The called function
					// is responsible for having stored the length of what's in the memory as an overload of
					// this_token.buf, but only when that memory is the result (currently might always be true).
					// AcceptNewMem() will shrink the memory for us, via _expand(), if there's a lot of
					// extra/unused space in it.
					internal_output_var->AcceptNewMem(result_token.mem_to_free, result_token.marker_length);
				}
				else
				{
					// 1.0.46.06: If the UDF has stored its result in its deref buffer, take possession
					// of that buffer, which saves a memcpy of a potentially huge string.  The cost
					// of this is that if there are any other UDF-calls pending after this one, the
					// code in their bodies will have to create another deref buffer if they need one.
					if (result == sDerefBuf && result_length >= MAX_ALLOC_SIMPLE) // Result is in their buffer and it's longer than what can fit in a SimpleHeap variable (avoids wasting SimpleHeap memory).
					{
						internal_output_var->AcceptNewMem(result, result_length);
						NULLIFY_S_DEREF_BUF // Force any UDFs called subsequently by us to create a new deref buffer because this one was just taken over by a variable.
					}
					else
					{
						// v1.0.45: This mode improves performance by avoiding the need to copy the result into
						// more persistent memory, then avoiding the need to copy it into the defer buffer (which
						// also avoids the possibility of needing to expand that buffer).
						if (!internal_output_var->Assign(result, result_length)) // Assign() contains an optimization that avoids actually doing the mem-copying if output_var is being assigned to itself (which can happen in cases like RegExMatch()).
							goto abort;
					}
				}
				if (done)
					goto normal_end_skip_output_var; // No need to restore circuit_token because the expression is finished.
				// Next operation is ":=" and above has verified the target is SYM_VAR and VAR_NORMAL.
				--stack_count; // STACK_POP;
				this_token.var = internal_output_var; // Make the result a variable rather than a normal operand so that its
				this_token.symbol = SYM_VAR; // address can be taken, and it can be passed ByRef. e.g. &(x:=1)
				++this_postfix; // We've fully handled the assignment.
				goto push_this_token;
			}
			// Otherwise, there's no output_var or the expression isn't finished yet, so do normal processing.
				
			make_result_persistent = true; // Set default.
			this_token.symbol = SYM_STRING;
			this_token.marker_length = result_length;

			if (result_token.mem_to_free) // The called function allocated some memory and turned it over to us.
			{
				// mem_to_free == result is checked only in debug mode because it should always be true.
				// Other sections rely on mem_to_free not needing to be freed if symbol != SYM_STRING,
				// so users of mem_to_free must never use it other than to return the result.
				ASSERT(result_token.mem_to_free == result);
				if (done && aResultToken)
				{
					// Return this memory block to our caller.  This is handled here rather than
					// at a later stage in order to avoid an unnecessary _tcslen() call.
					aResultToken->AcceptMem(result_to_return = result, result_length);
					goto normal_end_skip_output_var;
				}
				if (to_free_count == MAX_EXPR_MEM_ITEMS) // No more slots left (should be nearly impossible).
				{
					LineError(ERR_OUTOFMEM, FAIL, func->mName);
					goto abort;
				}
				// Mark it to be freed at the time we return.
				to_free[to_free_count++] = &this_token;
				// Invariant: any string token put in to_free must have marker set to the memory block
				// to be freed.  marker = result is set further below, but only when result_length != 0.
				this_token.marker = result;
				goto push_this_token;
			}
			//else this_token.mem_to_free==NULL, so the BIF just called didn't allocate memory to give to us.
			
			// Empty strings are returned pretty often by UDFs, such as when they don't use "return"
			// at all.  Therefore, handle them fully now, which should improve performance (since it
			// avoids all the other checking later on).  It also doesn't hurt code size because this
			// check avoids having to check for empty string in other sections later on.
			if (result_length == 0) // Various make-persistent sections further below may rely on this check.
			{
				this_token.marker = _T(""); // Ensure it's a constant memory area, not a buf that might get overwritten soon.
				goto push_this_token;
			}

			if (make_result_persistent) // At this stage, this means that the above wasn't able to determine its correct value yet.
			if (func->mIsBuiltIn)
			{
				// Since above didn't goto, "result" is not SYM_INTEGER/FLOAT/VAR, and not "".  Therefore, it's
				// either a pointer to static memory (such as a constant string), or more likely the small buf
				// we gave to the BIF for storing small strings.  For simplicity assume it's the buf, which is
				// volatile and must be made persistent if called for below.
				make_result_persistent = !done;
			}
			else // It's not a built-in function.
			{
				// Since above didn't goto, the result may need to be copied to a more persistent location.

				// For UDFs, "result" can be any of the following (some of which were handled above):
				//	- mem_to_free:  Passed back from some other function call.
				//	- mem_to_free:  The result of a concat (if other allocation methods were unavailable).
				//	- mem_to_free:  "Stolen" from a local var by ToReturnValue().
				//	- left_buf:  A local var's contents, copied into result_token.buf by ToReturnValue().
				//	- sDerefBuf:  Any other string result of ExpandExpression().
				//	- sDerefBuf:  A value copied from a variable by ExpandArgs():  return just_a_var
				//	- A literal string which was optimized into a non-expression:  return "just a string"
				//	- The Contents() of a static or global variable, which ExpandArgs() determined did not need
				//	  to be dereferenced. Only applies when !is_expression; i.e.  return static_var

				// Old method, not necessary to be so thorough because "return" always puts its result as the
				// very first item in its deref buf.  So this is commented out in favor of the line below it:
				//if (result < sDerefBuf || result >= sDerefBuf + sDerefBufSize)
				if (result != sDerefBuf) // Not in their deref buffer (yields correct result even if sDerefBuf is NULL; also, see above.)
					// In this case, the result can probably only be left_buf or the Contents() of a var,
					// either of which may need to be made persistent if the expression isn't finished:
					make_result_persistent = !done;
				else // The result must be in their deref buffer, perhaps due to something like "return x+3" or "return bif()" on their part.
				{
					make_result_persistent = false; // Set default to be possibly overridden below.
					if (!done) // There are more operators/operands to be evaluated, but if there are no more function calls, we don't have to make it persistent since their deref buf won't be overwritten by anything during the time we need it.
					{
						// Since there's more in the stack or postfix array to be evaluated, and since the return
						// value is in the new deref buffer, must copy result to somewhere non-volatile whenever
						// there's another function-call pending by us.  Note that an empty-string result was
						// already checked and fully handled higher above.
						// If we don't have any more user-defined function calls pending, we can skip the
						// make-persistent section since this deref buffer will not be overwritten during the
						// period we need it.
						for (p_postfix = this_postfix + 1; p_postfix->symbol != SYM_INVALID; ++p_postfix)
							if (p_postfix->symbol == SYM_FUNC)
							{
								make_result_persistent = true;
								break;
							}
					}
					//else done==true, so don't have to make it persistent here because the final stage will
					// copy it from their deref buf into ours (since theirs is only deleted later, by our caller).
					// In this case, leave make_result_persistent set to false.
				}
				// This is the end of the section that determines the value of "make_result_persistent" for UDFs.
			}

			if (make_result_persistent) // Both UDFs and built-in functions have ensured make_result_persistent is set.
			{
				// BELOW RELIES ON THE ABOVE ALWAYS HAVING VERIFIED AND FULLY HANDLED RESULT BEING AN EMPTY STRING.
				// So now we know result isn't an empty string, which in turn ensures that size > 1 and length > 0,
				// which might be relied upon by things further below.
				result_size = result_length + 1;
				// Must cast to int to avoid loss of negative values:
				if (result_size <= aDerefBufSize - (target - aDerefBuf)) // There is room at the end of our deref buf, so use it.
				{
					// Make the token's result the new, more persistent location:
					this_token.marker = tmemcpy(target, result, result_size); // Benches slightly faster than strcpy().
					target += result_size; // Point it to the location where the next string would be written.
				}
				else if (result_size < EXPR_SMALL_MEM_LIMIT && alloca_usage < EXPR_ALLOCA_LIMIT) // See comments at EXPR_SMALL_MEM_LIMIT.
				{
					this_token.marker = tmemcpy(talloca(result_size), result, result_size); // Benches slightly faster than strcpy().
					alloca_usage += result_size; // This might put alloca_usage over the limit by as much as EXPR_SMALL_MEM_LIMIT, but that is fine because it's more of a guideline than a limit.
				}
				else // Need to create some new persistent memory for our temporary use.
				{
					// In real-world scripts the need for additional memory allocation should be quite
					// rare because it requires a combination of worst-case situations:
					// - Called-function's return value is in their new deref buf (rare because return
					//   values are more often literal numbers, true/false, or variables).
					// - We still have more functions to call here (which is somewhat atypical).
					// - There's insufficient room at the end of the deref buf to store the return value
					//   (unusual because the deref buf expands in block-increments, and also because
					//   return values are usually small, such as numbers).
					if (to_free_count == MAX_EXPR_MEM_ITEMS // No more slots left (should be nearly impossible).
						|| !(this_token.marker = tmalloc(result_size)))
					{
						LineError(ERR_OUTOFMEM, FAIL, func->mName);
						goto abort;
					}
					tmemcpy(this_token.marker, result, result_size); // Benches slightly faster than strcpy().
					to_free[to_free_count++] = &this_token;
				}
			}
			else // make_result_persistent==false
				this_token.marker = result;

			goto push_this_token;
		} // if (this_token.symbol == SYM_FUNC)

		if (this_token.symbol == SYM_IFF_ELSE)
		{
			// SYM_IFF_ELSE is encountered only when a previous iteration has determined that the ternary's condition
			// is true.  At this stage, the ternary's "THEN" branch has already been evaluated and stored at the top
			// of the stack.  So skip over its "else" branch (short-circuit) because that doesn't need to be evaluated.
			this_postfix = this_token.circuit_token; // The address in any circuit_token always points into the arg's postfix array (never any temporary array or token created here) due to the nature/definition of circuit_token.
			// And very soon, the outer loop will skip over the SYM_IFF_ELSE just found above.
			continue;
		}

		// Since the above didn't goto or continue, this token must be a unary or binary operator.
		// Get the first operand for this operator (for non-unary operators, this is the right-side operand):
		if (!stack_count) // Prevent stack underflow.  An expression such as -*3 causes this.
			goto abort_with_exception;
		ExprTokenType &right = *STACK_POP;
		if (!IS_OPERAND(right.symbol)) // Haven't found a way to produce this situation yet, but safe to assume it's possible.
			goto abort_with_exception;

		switch (this_token.symbol)
		{
		case SYM_ASSIGN:        // These don't need "right_is_number" to be resolved. v1.0.48.01: Also avoid
		case SYM_CONCAT:        // resolving right_is_number for CONCAT because TokenIsPureNumeric() will take
		case SYM_ASSIGN_CONCAT: // a long time if the string is very long and consists entirely of digits/whitespace.
		case SYM_IS:
		case SYM_IN:
		case SYM_CONTAINS:
			right_is_pure_number = right_is_number = PURE_NOT_NUMERIC; // Init for convenience/maintainability.
		case SYM_ADDRESS:
		case SYM_AND:			// v2: These don't need it either since even numeric strings are considered "true".
		case SYM_OR:			//
		case SYM_LOWNOT:		//
		case SYM_HIGHNOT:		//
			break;
			
		case SYM_COMMA: // This can only be a statement-separator comma, not a function comma, since function commas weren't put into the postfix array.
			// Do nothing other than discarding the operand that was just popped off the stack, which is the
			// result of the comma's left-hand sub-statement.  At this point the right-hand sub-statement
			// has not yet been evaluated.  Like C++ and other languages, but unlike AutoHotkey v1, the
			// rightmost operand is preserved, not the leftmost.
			continue;

		default:
			// If the operand is still generic/undetermined, find out whether it is a string, integer, or float:
			right_is_pure_number = TokenIsPureNumeric(right, right_is_number); // If it's SYM_VAR, it can be the clipboard in this case, but it works even then.
		}

		// IF THIS IS A UNARY OPERATOR, we now have the single operand needed to perform the operation.
		// The cases in the switch() below are all unary operators.  The other operators are handled
		// in the switch()'s default section:
		sym_assign_var = NULL; // Set default for use at the bottom of the following switch().
		switch (this_token.symbol)
		{
		case SYM_AND:
		case SYM_OR:
		case SYM_IFF_THEN:
			// this_token is the left branch of an AND/OR or the condition of a ternary op.  Check for short-circuit.
			left_branch_is_true = TokenToBOOL(right);

			if (left_branch_is_true == (this_token.symbol == SYM_OR))
			{
				// The ternary's condition is false or this AND/OR causes a short-circuit.
				// Discard the entire right branch of this AND/OR or "then" branch of this IFF:
				this_postfix = this_token.circuit_token; // The address in any circuit_token always points into the arg's postfix array (never any temporary array or token created here) due to the nature/definition of circuit_token.

				if (this_token.symbol != SYM_IFF_THEN)
				{
					// This will be the final result of this AND/OR because it's right branch was
					// discarded above without having been evaluated nor any of its functions called:
					this_token.CopyValueFrom(right);
					// Any SYM_OBJECT on our stack was already put into to_free[], so if this is SYM_OBJECT,
					// there's no need to do anything; we actually MUST NOT AddRef() unless we also put it
					// into to_free[].
					break;
				}
			}
			else
			{
				// AND/OR: This left branch is simply discarded (by means of the outer loop) because its
				//	right branch will be the sole determination of whether this AND/OR is true or false.
				// IFF: The ternary's condition is true.  Do nothing; just let subsequent iterations evaluate
				//	the THEN portion; the SYM_IFF_ELSE which follows it will jump over the ELSE branch.
			}
			continue;

		case SYM_LOWNOT:  // The operator-word "not".
		case SYM_HIGHNOT: // The symbol '!'. Both NOTs are equivalent at this stage because precedence was already acted upon by infix-to-postfix.
			this_token.SetValue(!TokenToBOOL(right)); // Result is always one or zero.
			break;

		case SYM_NEGATIVE:  // Unary-minus.
			if (right_is_number == PURE_INTEGER)
				this_token.value_int64 = -TokenToInt64(right);
			else if (right_is_number == PURE_FLOAT)
				this_token.value_double = -TokenToDouble(right, FALSE); // Pass FALSE for aCheckForHex since PURE_FLOAT is never hex.
			else // String.
			{
				// Seems best to consider the application of unary minus to a string to be a failure.
				this_token.SetValue(EXPR_NAN);
				break;
			}
			// Since above didn't "break":
			this_token.symbol = right_is_number;
			break;

		case SYM_POSITIVE: // Added in v2 for symmetry with SYM_NEGATIVE; i.e. if -x produces NaN, so should +x.
			if (right_is_number)
				TokenToDoubleOrInt64(right, this_token);
			else
				this_token.SetValue(EXPR_NAN); // For consistency with unary minus (see above).
			break;

		case SYM_POST_INCREMENT: // These were added in v1.0.46.  It doesn't seem worth translating them into
		case SYM_POST_DECREMENT: // += and -= at load-time or during the tokenizing phase higher above because 
		case SYM_PRE_INCREMENT:  // it might introduce precedence problems, plus the post-inc/dec's nature is
		case SYM_PRE_DECREMENT:  // unique among all the operators in that it pushes an operand before the evaluation.
			if (right.symbol != SYM_VAR) // Syntax error.
				goto abort_with_exception;
			is_pre_op = SYM_INCREMENT_OR_DECREMENT_IS_PRE(this_token.symbol); // Store this early because its symbol will soon be overwritten.
			if (!*right.var->Contents()) // It's empty (this also serves to display a warning if applicable).
			{
				// For convenience, treat an empty variable as zero for ++ and --.
				// Consistent with v1 ++/-- when not combined with another expression.
				right.var->Assign(0);
				right_is_number = PURE_INTEGER;
			}
			else if (right_is_number == PURE_NOT_NUMERIC) // Not empty and not numeric: invalid operation.
			{
				right.var->Assign(EXPR_NAN); // Clipboard is also supported here.
				if (is_pre_op)
				{
					// v1.0.46.01: For consistency, it seems best to make the result of a pre-op be a
					// variable whenever a variable came in.  This allows its address to be taken, and it
					// to be passed by reference, and other SYM_VAR behaviors, even if the operation itself
					// produces a blank value.
					if (right.var->Type() == VAR_NORMAL)
					{
						this_token.var = right.var;  // Make the result a variable rather than a normal operand so that its
						this_token.symbol = SYM_VAR; // address can be taken, and it can be passed ByRef. e.g. &(++x)
						break;
					}
					//else VAR_CLIPBOARD, which is allowed in only when it's the lvalue of an assignment or
					// inc/dec.  So fall through to make the result blank because clipboard isn't allowed as
					// SYM_VAR beyond this point (to simplify the code and improve maintainability).
				}
				this_token.SetValue(EXPR_NAN); // Indicate invalid operation (increment/decrement a non-number).
				break;
			} // end of "invalid operation" block.

			// DUE TO CODE SIZE AND PERFORMANCE decided not to support things like the following:
			// -> ++++i ; This one actually works because pre-ops produce a variable (usable by future pre-ops).
			// -> i++++ ; Fails because the first ++ produces an operand that isn't a variable.  It could be
			//    supported via a cascade loop here to pull all remaining consecutive post/pre ops out of
			//    the postfix array and apply them to "delta", but it just doesn't seem worth it.
			// -> --Var++ ; Fails because ++ has higher precedence than --, but it produces an operand that isn't
			//    a variable, so the "--" fails.  Things like --Var++ seem pointless anyway because they seem
			//    nearly identical to the sub-expression (Var+1)? Anyway, --Var++ could probably be supported
			//    using the loop described in the previous example.
			delta = (this_token.symbol == SYM_POST_INCREMENT || this_token.symbol == SYM_PRE_INCREMENT) ? 1 : -1;
			if (right_is_number == PURE_INTEGER)
			{
				this_token.value_int64 = TokenToInt64(right);
				right.var->Assign(this_token.value_int64 + delta);
			}
			else // right_is_number must be PURE_FLOAT because it's the only remaining alternative.
			{
				this_token.value_double = TokenToDouble(right, FALSE); // Pass FALSE for aCheckForHex since PURE_FLOAT is never hex.
				right.var->Assign(this_token.value_double + delta);
			}
			if (is_pre_op)
			{
				// Push the variable itself so that the operation will have already taken effect for whoever
				// uses this operand/result in the future (i.e. pre-op vs. post-op).
				// KNOWN LIMITATION: Although this behavior is convenient to have, I realize now
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
				}
				else // VAR_CLIPBOARD, which is allowed in only when it's the lvalue of an assignment or inc/dec.
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
			if (IObject *obj = TokenToObject(right))
			{
				this_token.SetValue((__int64)obj);
			}
			else if (right.symbol == SYM_VAR) // At this stage, SYM_VAR is always a normal variable, never a built-in one, so taking its address should be safe.
			{
				Var *right_var = right.var->ResolveAlias();
				if (right_var->IsPureNumeric())
					this_token.value_int64 = (__int64)&right_var->mContentsInt64; // Since the value is a pure number, this seems more useful and less confusing than returning the address of a numeric string.
				else
					this_token.value_int64 = (__int64)right_var->Contents(); // Contents() vs. mContents to support VAR_CLIPBOARD, and in case mContents needs to be updated by Contents().
				this_token.symbol = SYM_INTEGER;
			}
			else // Syntax error: operand is not an object or a variable reference.
				goto abort_with_exception;
			break;

		case SYM_BITNOT:  // The tilde (~) operator.
			if (right_is_number == PURE_NOT_NUMERIC) // String.  Seems best to consider the application of '*' or '~' to a non-numeric string to be a failure.
			{
				this_token.SetValue(EXPR_NAN);
				break;
			}
			// Since above didn't "break": right_is_number is PURE_INTEGER or PURE_FLOAT.
			right_int64 = TokenToInt64(right); // Although PURE_FLOAT can't be hex, for simplicity and due to the rarity of encountering a PURE_FLOAT in this case, the slight performance reduction of calling TokenToInt64() is done for both PURE_FLOAT and PURE_INTEGER.
			// Note that it is not legal to perform ~, &, |, or ^ on doubles.  Because of this,
			// any floating point operand is truncated to an integer above.
			if (right_int64 < 0 || right_int64 > UINT_MAX)
				// Treat it as a 64-bit signed value, since no other aspects of the program
				// (e.g. IfEqual) will recognize an unsigned 64 bit number.
				this_token.value_int64 = ~right_int64;
			else
				// Treat it as a 32-bit unsigned value when inverting and assigning.  This is
				// because assigning it as a signed value would "convert" it into a 64-bit
				// value, which in turn is caused by the fact that the script sees all negative
				// numbers as 64-bit values (e.g. -1 is 0xFFFFFFFFFFFFFFFF).
				this_token.value_int64 = (size_t)(DWORD)~(DWORD)right_int64; // Casting this way avoids compiler warning.
			this_token.symbol = SYM_INTEGER; // Must be done only after its old value was used above. v1.0.36.07: Fixed to be SYM_INTEGER vs. right_is_number for SYM_BITNOT.
			break;

		default: // NON-UNARY OPERATOR.
			// GET THE SECOND (LEFT-SIDE) OPERAND FOR THIS OPERATOR:
			if (!stack_count) // Prevent stack underflow.
				goto abort_with_exception;
			ExprTokenType &left = *STACK_POP; // i.e. the right operand always comes off the stack before the left.
			if (!IS_OPERAND(left.symbol)) // Haven't found a way to produce this situation yet, but safe to assume it's possible.
				goto abort_with_exception;
			
			if (IS_ASSIGNMENT_EXCEPT_POST_AND_PRE(this_token.symbol)) // v1.0.46: Added support for various assignment operators.
			{
				if (left.symbol != SYM_VAR) // Syntax error.
					goto abort_with_exception;

				switch(this_token.symbol)
				{
				case SYM_ASSIGN: // Listed first for performance (it's probably the most common because things like ++ and += aren't expressions when they're by themselves on a line).
					if (!left.var->Assign(right)) // left.var can be VAR_CLIPBOARD in this case.
						goto abort;
					if (left.var->Type() != VAR_NORMAL) // Could be VAR_CLIPBOARD or VAR_VIRTUAL, which should not yield SYM_VAR (as some sections of the code wouldn't handle it correctly).
					{
						this_token.CopyValueFrom(right); // Doing it this way is more maintainable than other methods, and is unlikely to perform much worse.
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
				// Since above didn't goto or break out of the outer loop, this is an assignment other than
				// SYM_ASSIGN, so it needs further evaluation later below before the assignment will actually be made.
				sym_assign_var = left.var; // This tells the bottom of this switch() to do extra steps for this assignment.
				if ((this_token.symbol == SYM_ADD || this_token.symbol == SYM_SUBTRACT)
					&& !*sym_assign_var->Contents()) // It's empty (this also serves to display a warning if applicable).
				{
					// For convenience, treat an empty variable as zero for += and -=.
					// Consistent with v1 EnvAdd/EnvSub or +=/-= when not combined with another expression.
					left.symbol = SYM_INTEGER;
					left.value_int64 = 0;
				}
			}

			// The following section needs done even for assignments such as += because the type of value
			// inside the target variable (integer vs. float vs. string) must be known, to determine how
			// the operation should proceed.
			// Since above didn't goto/break, this is a non-unary operator that needs further processing.
			// If the operand is still generic/undetermined, find out whether it is a string, integer, or float.
			// Fix for v1.0.48.01: For performance, call TokenIsPureNumeric(left) only as a last resort
			// because "left" could be a very long string consisting entirely of digits or whitespace, which
			// would make the call take a long time.  
			if (right_is_number) // right_is_number is always PURE_NOT_NUMERIC for SYM_CONCAT.
				left_is_pure_number = TokenIsPureNumeric(left, left_is_number);
			// Otherwise, leave left_is' uninitialized as below will short-circuit.
			if (  !(right_is_number && left_is_number)  // i.e. they're not both numeric (or this is SYM_CONCAT).
				|| IS_RELATIONAL_OPERATOR(this_token.symbol) && !right_is_pure_number && !left_is_pure_number  ) // i.e. if both are strings, compare them alphabetically.
			{
				// L31: Handle binary ops supported by objects (= == !=).
				switch (this_token.symbol)
				{
				case SYM_EQUAL:
				case SYM_EQUALCASE:
				case SYM_NOTEQUAL:
					IObject *right_obj = TokenToObject(right);
					IObject *left_obj = TokenToObject(left);
					// To support a future "implicit default value" feature, both operands must be objects.
					// Otherwise, an object operand will be treated as its default value, currently always "".
					// This is also consistent with unsupported operands such as < and > - i.e. because obj<""
					// and obj>"" are always false and obj<="" and obj>="" are always true, obj must be "".
					// When the default value feature is implemented all operators (excluding =, == and !=
					// if both operands are objects) may use the default value of any object operand.
					// UPDATE: Above is not done because it seems more intuitive to document the other
					// comparison operators as unsupported than for (obj == "") to evaluate to true.
					if (right_obj || left_obj)
					{
						this_token.SetValue((this_token.symbol != SYM_NOTEQUAL) == (right_obj == left_obj));
						goto push_this_token;
					}
				}

				// Above check has ensured that at least one of them is a string.  But the other
				// one might be a number such as in 5+10="15", in which 5+10 would be a numerical
				// result being compared to the raw string literal "15".
				right_string = TokenToString(right, right_buf, &right_length);
				left_string = TokenToString(left, left_buf, &left_length);
				result_symbol = SYM_INTEGER; // Set default.  Boolean results are treated as integers.
				switch(this_token.symbol)
				{
				case SYM_EQUAL:     this_token.value_int64 = !((g->StringCaseSense == SCS_INSENSITIVE)
										? _tcsicmp(left_string, right_string)
										: lstrcmpi(left_string, right_string)); break; // i.e. use the "more correct mode" except when explicitly told to use the fast mode (v1.0.43.03).
				case SYM_EQUALCASE: // Case sensitive.  Also supports binary data.
					// Support basic equality checking of binary data by using tmemcmp rather than _tcscmp.
					// The results should be the same for strings, but faster.  Length must be checked first
					// since tmemcmp wouldn't stop at the null-terminator (and that's why we're using it).
					// As a result, the comparison is much faster when the length differs.
					this_token.value_int64 = (left_length == right_length) && !tmemcmp(left_string, right_string, left_length);
					break; 
				// The rest all obey g->StringCaseSense since they have no case sensitive counterparts:
				case SYM_NOTEQUAL:  this_token.value_int64 = g_tcscmp(left_string, right_string) ? 1 : 0; break;
				case SYM_GT:        this_token.value_int64 = g_tcscmp(left_string, right_string) > 0; break;
				case SYM_LT:        this_token.value_int64 = g_tcscmp(left_string, right_string) < 0; break;
				case SYM_GTOE:      this_token.value_int64 = g_tcscmp(left_string, right_string) > -1; break;
				case SYM_LTOE:      this_token.value_int64 = g_tcscmp(left_string, right_string) < 1; break;

				case SYM_CONCAT:
					// Even if the left or right is "", must copy the result to temporary memory, at least
					// when integers and floats had to be converted to temporary strings above.
					// Binary clipboard is ignored because it's documented that except for certain features,
					// binary clipboard variables are seen only up to the first binary zero (mostly to
					// simplify the code).
					if (sym_assign_var && sym_assign_var->Type() == VAR_NORMAL) // Since "right" is being appended onto a variable ("left"), an optimization is possible.
					{
						// Append() is particularly efficient when the var already has room to append the value,
						// but improves performance even in other cases by avoiding an extra memcpy and allocating
						// extra space for future expansion.  It is necessary to completely handle this case here
						// because otherwise this_token might need to be put into to_mem[], in which case it must
						// not be converted to SYM_VAR.
						if (!sym_assign_var->Append(right_string, (VarSizeType)right_length))
							goto abort;
						this_token.var = sym_assign_var; // Make the result a variable rather than a normal operand so that its
						this_token.symbol = SYM_VAR;     // address can be taken, and it can be passed ByRef. e.g. &(x+=1)
						goto push_this_token; // Skip over all other sections such as subsequent checks of sym_assign_var because it was all taken care of here.
					}
					// Otherwise, fall back to the other concat methods:
					result_size = right_length + left_length + 1;

					if (sym_assign_var)  // Fix for v1.0.48: These 2 lines were added, and they must take
						temp_var = NULL; // precendence over the other checks below to allow an expression like the following to work: var := var2 .= "abc"
					else if (output_var && EXPR_IS_DONE) // i.e. this is ACT_ASSIGNEXPR and we're at the final operator, a concat.
					{
						temp_var = output_var;
						done_and_have_an_output_var = TRUE;
					}
					else if (this_postfix[1].symbol == SYM_ASSIGN // Next operation is ":=".
						&& stack_count && stack[stack_count-1]->symbol == SYM_VAR // i.e. let the next iteration handle it instead of doing it here.  Further below relies on this having been checked.
						&& stack[stack_count-1]->var->Type() == VAR_NORMAL) // Don't do clipboard here because: 1) AcceptNewMem() doesn't support it; 2) Could probably use Assign() and then make its result be a newly added mem_count item, but the code complexity doesn't seem worth it given the rarity.
					{
						temp_var = stack[stack_count-1]->var;
						done_and_have_an_output_var = FALSE;
					}
					else
						temp_var = NULL;

					if (temp_var)
					{
						result = temp_var->Contents(FALSE); // No need to update the contents because we just want to know if the current address of mContents matches some other addresses.
						if (result == Var::sEmptyString) // Added in v1.1.09.03.
						{
							// One of the following is true:
							//   1) temp_var has zero capacity and is empty.
							//   2) temp_var has zero capacity and contains an unflushed binary number.
							// In the first case, AppendIfRoom() will always fail, so we want to skip it and use
							// the "no overlap" optimization below. In the second case, calling AppendIfRoom()
							// would produce the wrong result; e.g. (x := 0+1, x := y 0) would produce "10".
							result = NULL;
						}
						if (result == left_string) // This is something like x := x . y, so simplify it to x .= y
						{
							// MUST DO THE ABOVE CHECK because the next section further below might free the
							// destination memory before doing the operation. Thus, if the destination is the
							// same as one of the sources, freeing it beforehand would obviously be a problem.
							if (temp_var->AppendIfRoom(right_string, (VarSizeType)right_length))
							{
								if (done_and_have_an_output_var) // Fix for v1.0.48: Checking "temp_var == output_var" would not be enough for cases like v := (v := v . "a") . "b"
									goto normal_end_skip_output_var; // Nothing more to do because it has even taken care of output_var already.
								else // temp_var is from look-ahead to a future assignment.
								{
									++this_postfix;
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
							if (!temp_var->AssignString(NULL, (VarSizeType)result_size - 1)) // Resize the destination, if necessary.
								goto abort; // Above should have already reported the error.
							result = temp_var->Contents(); // Call Contents() AGAIN because Assign() may have changed it.  No need to pass FALSE because the call to Assign() above already reset the contents.
							if (left_length)
								tmemcpy(result, left_string, left_length);  // Not +1 because don't need the zero terminator.
							tmemcpy(result + left_length, right_string, right_length + 1); // +1 to include its zero terminator.
							temp_var->Close(); // Must be called after Assign(NULL, ...) or when Contents() has been altered because it updates the variable's attributes and properly handles VAR_CLIPBOARD.
							if (done_and_have_an_output_var) // Fix for v1.0.48: Checking "temp_var == output_var" would not be enough for cases like v := (v := "a" . "b") . "c".
								goto normal_end_skip_output_var; // Nothing more to do because it has even taken care of output_var already.
							else // temp_var is from look-ahead to a future assignment.
							{
								++this_postfix;
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
						target += result_size;  // Adjust target for potential future use by another concat or function call.
					}
					else if (result_size < EXPR_SMALL_MEM_LIMIT && alloca_usage < EXPR_ALLOCA_LIMIT) // See comments at EXPR_SMALL_MEM_LIMIT.
					{
						this_token.marker = talloca(result_size);
						alloca_usage += result_size; // This might put alloca_usage over the limit by as much as EXPR_SMALL_MEM_LIMIT, but that is fine because it's more of a guideline than a limit.
					}
					else // Need to create some new persistent memory for our temporary use.
					{
						// See the nearly identical section higher above for comments:
						if (to_free_count == MAX_EXPR_MEM_ITEMS // No more slots left (should be nearly impossible).
							|| !(this_token.marker = tmalloc(result_size)))
						{
							LineError(ERR_OUTOFMEM);
							goto abort;
						}
						to_free[to_free_count++] = &this_token;
					}
					if (left_length)
						tmemcpy(this_token.marker, left_string, left_length);  // Not +1 because don't need the zero terminator.
					tmemcpy(this_token.marker + left_length, right_string, right_length + 1); // +1 to include its zero terminator.
					this_token.marker_length = result_size - 1;
					result_symbol = SYM_STRING;
					break;

				case SYM_IS:
					if (!ValueIsType(this_token, left, left_string, right, right_string))
						goto abort;
					break;

				default:
					// Other operators do not support non-numeric operands, so the result is NaN (not a number).
					this_token.marker = EXPR_NAN_STR;
					this_token.marker_length = EXPR_NAN_LEN;
					result_symbol = SYM_STRING;
				}
				this_token.symbol = result_symbol; // Must be done only after the switch() above.
			}

			else if (right_is_number == PURE_INTEGER && left_is_number == PURE_INTEGER && this_token.symbol != SYM_DIVIDE
				|| IS_BIT_OPERATOR(this_token.symbol))
			{
				// Because both are integers and the operation isn't division, the result is integer.
				// The result is also an integer for the bitwise operations listed in the if-statement
				// above.  This is because it is not legal to perform ~, &, |, or ^ on doubles.  Any
				// floating point operands are truncated to integers prior to doing the bitwise operation.
				right_int64 = TokenToInt64(right); // It can't be SYM_STRING because in here, both right and
				left_int64 = TokenToInt64(left);    // left are known to be numbers (otherwise an earlier "else if" would have executed instead of this one).
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
					if (right_int64 == 0) // Divide by zero produces "not a number".
					{
						this_token.marker = EXPR_NAN_STR;
						this_token.marker_length = EXPR_NAN_LEN;
						result_symbol = SYM_STRING;
					}
					else
						this_token.value_int64 = left_int64 / right_int64;
					break;
				case SYM_POWER:
					// Note: The function pow() in math.h adds about 28 KB of code size (uncompressed)!
					// Even assuming pow() supports negative bases such as (-2)**2, its size is why it's not used.
					// v1.0.44.11: With Laszlo's help, negative integer bases are now supported.
					if (!left_int64 && right_int64 < 0) // In essence, this is divide-by-zero.
					{
						// Return a consistent result rather than something that varies:
						this_token.marker = EXPR_NAN_STR;
						this_token.marker_length = EXPR_NAN_LEN;
						result_symbol = SYM_STRING;
					}
					else // We have a valid base and exponent and both are integers, so the calculation will always have a defined result.
					{
						if (left_was_negative = (left_int64 < 0))
							left_int64 = -left_int64; // Force a positive due to the limitations of qmathPow().
						this_token.value_double = qmathPow((double)left_int64, (double)right_int64);
						if (left_was_negative && right_int64 % 2) // Negative base and odd exponent (not zero or even).
							this_token.value_double = -this_token.value_double;
						if (right_int64 < 0)
							result_symbol = SYM_FLOAT; // Due to negative exponent, override to float.
						else
							this_token.value_int64 = (__int64)this_token.value_double;
					}
					break;
				}
				this_token.symbol = result_symbol; // Must be done only after the switch() above.
			}

			else // Since one or both operands are floating point (or this is the division of two integers), the result will be floating point.
			{
				right_double = TokenToDouble(right, TRUE); // Pass TRUE for aCheckForHex in case one of them is an integer to
				left_double = TokenToDouble(left, TRUE);   // be converted to a float for the purpose of this calculation.
				result_symbol = IS_RELATIONAL_OPERATOR(this_token.symbol) ? SYM_INTEGER : SYM_FLOAT; // Set default. v1.0.47.01: Changed relational operators to yield integers vs. floats because it's more intuitive and traditional (might also make relational operators perform better).
				switch(this_token.symbol)
				{
				case SYM_ADD:      this_token.value_double = left_double + right_double; break;
				case SYM_SUBTRACT: this_token.value_double = left_double - right_double; break;
				case SYM_MULTIPLY: this_token.value_double = left_double * right_double; break;
				case SYM_DIVIDE:
				case SYM_FLOORDIVIDE:
					if (right_double == 0.0) // Divide by zero produces "not a number".
					{
						this_token.marker = EXPR_NAN_STR;
						this_token.marker_length = EXPR_NAN_LEN;
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
					if (left_double == 0.0 && right_double < 0  // In essence, this is divide-by-zero.
						|| left_was_negative && qmathFmod(right_double, 1.0) != 0.0) // Negative base, but exponent isn't close enough to being an integer: unsupported (to simplify code).
					{
						this_token.marker = EXPR_NAN_STR;
						this_token.marker_length = EXPR_NAN_LEN;
						result_symbol = SYM_STRING;
					}
					else
					{
						if (left_was_negative)
							left_double = -left_double; // Force a positive due to the limitations of qmathPow().
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
			if (!sym_assign_var->Assign(this_token)) // Assign the result (based on its type) to the target variable.
				goto abort;
			if (sym_assign_var->Type() == VAR_NORMAL)
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

push_this_token:
		STACK_PUSH(&this_token);   // Push the result onto the stack for use as an operand by a future operator.
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
		goto abort_with_exception; // the overall result. Examples of errors include: () ... x y ... (x + y) (x + z) ... etc. (some of these might no longer produce this issue due to auto-concat).

	ExprTokenType &result_token = *stack[0];  // For performance and convenience.  Even for multi-statement, the bottommost item on the stack is the final result so that things like var1:=1,var2:=2 work.

	if (output_var)
	{
		// v1.0.45: Take a shortcut, which in the case of SYM_STRING/OPERAND/VAR avoids one memcpy
		// (into the deref buffer).  In some cases, this also saves from having to expand the deref buffer.
		if (!output_var->Assign(result_token))
			goto abort;
		goto normal_end_skip_output_var; // result_to_return is left at its default of "", though its value doesn't matter as long as it isn't NULL.
	}

	if (mActionType == ACT_IF || mActionType == ACT_WHILE || mActionType == ACT_UNTIL)
	{
		// This is an optimization that improves the speed of ACT_IF by up to 50% (ACT_WHILE is
		// probably improved by only up-to-15%). Simple expressions like "if (x < y)" see the biggest
		// speedup.
		result_to_return = TokenToBOOL(result_token) ? _T("1") : _T(""); // Return "" vs. "0" for FALSE for consistency with "goto abnormal_end" (which bypasses this section).
		goto normal_end_skip_output_var; // ACT_IF never has an output_var.
	}
	
	if (aResultToken)
	{
		switch (result_token.symbol)
		{
		case SYM_INTEGER:
		case SYM_FLOAT:
		case SYM_OBJECT:
			// Return numeric or object result as-is.
			aResultToken->symbol = result_token.symbol;
			aResultToken->value_int64 = result_token.value_int64; // Union copy.
			if (result_token.symbol == SYM_OBJECT)
				result_token.object->AddRef();
			goto normal_end_skip_output_var; // result_to_return is left at its default of "".
		case SYM_VAR:
			// This check must be done first to allow a variable containing a number or object to be passed ByRef:
			if (mActionType == ACT_FUNC || mActionType == ACT_METHOD)
			{
				aArgVar[aArgIndex] = result_token.var; // Let the command refer to this variable directly.
				goto normal_end_skip_output_var; // result_to_return is left at its default of "", though its value doesn't matter as long as it isn't NULL.
			}
			if (result_token.var->IsPureNumericOrObject())
			{
				result_token.var->ToToken(*aResultToken);
				goto normal_end_skip_output_var; // result_to_return is left at its default of "".
			}
			if (mActionType == ACT_RETURN && result_token.var->ToReturnValue(*aResultToken))
			{
				// This is a non-static local var and the call above has either transferred its memory block
				// into aResultToken or copied its value into aResultToken->buf.  This should be faster than
				// copying it into the deref buffer, and also allows binary data to be retained.
				// "case ACT_RETURN:" does something similar for non-expressions like "return local_var".
				goto normal_end_skip_output_var; // result_to_return is left at its default of "", though its value doesn't matter as long as it isn't NULL.
			}
			break;
		case SYM_STRING:
			if (to_free_count && to_free[to_free_count - 1] == &result_token)
			{
				// Pass this mem item back to caller instead of freeing it when we return.
				aResultToken->AcceptMem(result_to_return = result_token.marker, result_token.marker_length);
				--to_free_count;
				goto normal_end_skip_output_var;
			}
		}
		// Since above didn't return, the result is a string.  Continue on below to copy it into persistent memory.
	}
	
	//
	// Store the result of the expression in the deref buffer for the caller.
	//
	result_to_return = aTarget; // Set default.
	switch (result_token.symbol)
	{
	case SYM_INTEGER:
		// SYM_INTEGER and SYM_FLOAT will fit into our deref buffer because an earlier stage has already ensured
		// that the buffer is large enough to hold at least one number.  But a string/generic might not fit if it's
		// a concatenation and/or a large string returned from a called function.
		aTarget += _tcslen(ITOA64(result_token.value_int64, aTarget)) + 1; // Store in hex or decimal format, as appropriate.
		// Above: +1 because that's what callers want; i.e. the position after the terminator.
		goto normal_end_skip_output_var; // output_var was already checked higher above, so no need to consider it again.
	case SYM_FLOAT:
		aTarget += FTOA(result_token.value_double, aTarget, MAX_NUMBER_SIZE) + 1; // +1 because that's what callers want; i.e. the position after the terminator.
		goto normal_end_skip_output_var; // output_var was already checked higher above, so no need to consider it again.
	case SYM_VAR:
		// SYM_VAR is somewhat unusual at this late a stage.  Dynamic output vars were already handled by the
		// SYM_DYNAMIC code, while ACT_FUNC and ACT_METHOD were handled above.
		// It is tempting to simply return now and let ExpandArgs() decide whether the var needs to be dereferenced.
		// However, the var's length might not fit within the amount calculated by GetExpandedArgSize(), and in that
		// case would overflow the deref buffer.  The var can be safely returned if it won't be dereferenced, but it
		// doesn't seem worth duplicating the ArgMustBeDereferenced() logic here given the rarity of SYM_VAR results.
		// aArgVar[] isn't set because it would be redundant, and might cause issues if the var is modified as a
		// side-effect of a later arg (e.g. ArgLength() might return a wrong value).
	// FALL THROUGH TO BELOW:
	case SYM_STRING:
		// At this stage, we know the result has to go into our deref buffer because if a way existed to
		// avoid that, we would already have goto/returned higher above (e.g. for ACT_ASSIGNEXPR OR ACT_EXPRESSION.
		// Also, at this stage, the pending result can exist in one of several places:
		// 1) Our deref buf (due to being a single-deref, a function's return value that was copied to the
		//    end of our buf because there was enough room, etc.)
		// 2) In a called function's deref buffer, namely sDerefBuf, which will be deleted by our caller
		//    shortly after we return to it.
		// 3) In an area of memory we alloc'd for lack of any better place to put it.
		if (result_token.symbol == SYM_VAR)
		{
			result = result_token.var->Contents();
			result_size = result_token.var->Length() + 1;
		}
		else
		{
			result = result_token.marker;
			result_size = result_token.marker_length + 1; // At this stage, marker_length should always be valid, not -1.
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
		// 2) Any numeric result (i.e. MAX_NUMBER_LENGTH).  If the expression needs to store a string
		//    result, it will take care of expanding the deref buffer.
		#define EXPR_BUF_SIZE(raw_expr_len) (raw_expr_len < MAX_NUMBER_LENGTH \
			? MAX_NUMBER_LENGTH : raw_expr_len) + 1 // +1 for the overall terminator.

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
			LPTSTR new_buf;
			if (   !(new_buf = tmalloc(new_buf_size))   )
			{
				LineError(ERR_OUTOFMEM);
				goto abort;
			}
			if (new_buf_size > LARGE_DEREF_BUF_SIZE)
				++sLargeDerefBufs; // And if the old deref buf was larger too, this value is decremented later below. SET_DEREF_TIMER() is handled by our caller because aDerefBufSize is updated further below, which the caller will see.

			// Copy only that portion of the old buffer that is in front of our portion of the buffer
			// because we no longer need our portion (except for result.marker if it happens to be
			// in the old buffer, but that is handled after this):
			size_t aTarget_offset = aTarget - aDerefBuf;
			if (aTarget_offset) // aDerefBuf has contents that must be preserved.
				tmemcpy(new_buf, aDerefBuf, aTarget_offset); // This will also copy the empty string if the buffer first and only character is that.
			aTarget = new_buf + aTarget_offset;
			result_to_return = aTarget; // Update to reflect new value above.
			// NOTE: result_token.marker might extend too far to the right in our deref buffer and thus be
			// larger than capacity_of_our_buf_portion because other arg(s) exist in this line after ours
			// that will be using a larger total portion of the buffer than ours.  Thus, the following must be
			// done prior to free(), but memcpy() vs. memmove() is safe in any case:
			tmemcpy(aTarget, result, result_size); // Copy from old location to the newly allocated one.

			free(aDerefBuf); // Free our original buffer since it's contents are no longer needed.
			if (aDerefBufSize > LARGE_DEREF_BUF_SIZE)
				--sLargeDerefBufs;

			// Now that the buffer has been enlarged, need to adjust any other pointers that pointed into
			// the old buffer:
			LPTSTR aDerefBuf_end = aDerefBuf + aDerefBufSize; // Point it to the character after the end of the old buf.
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
				tmemmove(aTarget, result, result_size); // memmove() vs. memcpy() in this case, since source and dest might overlap (i.e. "target" may have been used to put temporary things into aTarget, but those things are no longer needed and now safe to overwrite).
		if (aResultToken)
		{
			aResultToken->marker = aTarget;
			aResultToken->marker_length = result_size - 1;
		}
		aTarget += result_size;
		goto normal_end_skip_output_var; // output_var was already checked higher above, so no need to consider it again.

	case SYM_OBJECT:
		// At this point we aren't capable of returning an object, otherwise above would have
		// already returned.  The documented fallback behaviour is for the object to be treated
		// as an empty string.
		result_to_return = _T("");
		goto normal_end_skip_output_var;
	} // switch (result_token.symbol)

// ALL PATHS ABOVE SHOULD "GOTO".  TO CATCH BUGS, ANY THAT DON'T FALL INTO "ABORT" BELOW.
abort_with_exception:
	LineError(ERR_EXPR_EVAL);
	// FALL THROUGH:
abort:
	// The callers of this function know that the value of aResult (which contains the reason
	// for early exit) should be considered valid/meaningful only if result_to_return is NULL.
	result_to_return = NULL; // Use NULL to inform our caller that this entire thread is to be terminated.
	aResult = FAIL; // Indicate reason to caller.
	goto normal_end_skip_output_var; // output_var is skipped as part of standard abort behavior.

//abnormal_end: // Currently the same as normal_end; it's separate to improve readability.  When this happens, result_to_return is typically "" (unless the caller overrode that default).
//normal_end: // This isn't currently used, but is available for future-use and readability.
	// v1.0.45: ACT_ASSIGNEXPR relies on us to set the output_var (i.e. whenever it's ARG1's is_expression==true).
	// Our taking charge of output_var allows certain performance optimizations in other parts of this function,
	// such as avoiding excess memcpy's and malloc's during intermediate stages.
	// v2: Leave output_var unchanged in this case so that ACT_ASSIGNEXPR behaves the same as SYM_ASSIGN.
	//if (output_var && result_to_return) // i.e. don't assign if NULL to preserve backward compatibility with scripts that rely on the old value being changed in cases where an expression fails (unlikely).
	//	if (!output_var->Assign(result_to_return))
	//		aResult = FAIL;

normal_end_skip_output_var:
	for (i = to_free_count; i--;) // Free any temporary memory blocks that were used.  Using reverse order might reduce memory fragmentation a little (depending on implementation of malloc).
	{
		if (to_free[i]->symbol == SYM_STRING)
			free(to_free[i]->marker);
		else // SYM_OBJECT
			to_free[i]->object->Release();
	}

	return result_to_return;
}



bool Func::Call(ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount, bool aIsVariadic)
{
	Object *param_obj = NULL;
	if (aIsVariadic) // i.e. this is a variadic function call.
	{
		ExprTokenType *rvalue = NULL;
		if (mBIF == &Op_ObjInvoke && mID == IT_SET && aParamCount > 1) // x[y*]:=z
			rvalue = aParam[--aParamCount];
		
		--aParamCount; // i.e. make aParamCount the count of normal params.
		if (param_obj = dynamic_cast<Object *>(TokenToObject(*aParam[aParamCount])))
		{
			int extra_params = param_obj->MaxIndex();
			if (extra_params > 0 || param_obj->HasNonnumericKeys())
			{
				// Calculate space required for ...
				size_t space_needed = extra_params * sizeof(ExprTokenType) // ... new param tokens
					+ max(mParamCount, aParamCount + extra_params) * sizeof(ExprTokenType *); // ... existing and new param pointers
				if (rvalue)
					space_needed += sizeof(rvalue); // ... extra slot for aRValue
				// Allocate new param list and tokens; tokens first for convenience.
				ExprTokenType *token = (ExprTokenType *)_alloca(space_needed);
				ExprTokenType **param_list = (ExprTokenType **)(token + extra_params);
				// Since built-in functions don't have variables we can directly assign to,
				// we need to expand the param object's contents into an array of tokens:
				param_obj->ArrayToParams(token, param_list, extra_params, aParam, aParamCount);
				aParam = param_list;
				aParamCount += extra_params;
			}
		}
		if (rvalue)
			aParam[aParamCount++] = rvalue; // In place of the variadic param.
	}

	if (mIsBuiltIn)
	{
		// mMinParams is validated at load-time where possible; so not for variadic or dynamic calls,
		// nor for calls via objects.  This check could be avoided for normal calls by instead checking
		// in each of the above cases, but any performance gain would probably be marginal and not worth
		// the slightly larger code size and loss of maintainability.  This check is not done for UDFS
		// since param_obj might contain the remaining parameters as name-value pairs.  Missing required
		// parameters are instead detected by the absence of a default value.
		if (aParamCount < mMinParams)
		{
			aResultToken.Error(ERR_TOO_FEW_PARAMS, mName);
			return false; // Abort expression.
		}
		// Otherwise, even if some params are SYM_MISSING, it is relatively safe to call the function.
		// The TokenTo' set of functions will produce 0 or "" for missing params.  Although that isn't
		// technically correct, it is simple and fairly logical.

		aResultToken.symbol = SYM_INTEGER; // Set default return type so that functions don't have to do it if they return INTs.
		aResultToken.func = this;          // Inform function of which built-in function called it (allows code sharing/reduction).

		// Push an entry onto the debugger's stack.  This has two purposes:
		//  1) Allow CreateRuntimeException() to know which function is throwing an exception.
		//  2) If a UDF is called before the BIF returns, it will show on the call stack.
		//     e.g. DllCall(RegisterCallback("F")) will show DllCall while F is running.
		DEBUGGER_STACK_PUSH(this)

		// CALL THE BUILT-IN FUNCTION:
		mBIF(aResultToken, aParam, aParamCount);

		DEBUGGER_STACK_POP()
		
		// There shouldn't be any need to check g->ThrownToken since built-in functions
		// currently throw exceptions via aResultToken.Error():
		//if (g->ThrownToken)
		//	aResultToken.SetExitResult(FAIL); // Abort thread.
	}
	else // It's not a built-in function, or it's a built-in that was overridden with a custom function.
	{
		ResultType result;
		VarBkp *backup = NULL;
		int backup_count;

		int j, count_of_actuals_that_have_formals;
		count_of_actuals_that_have_formals = (aParamCount > mParamCount)
			? mParamCount  // Omit any actuals that lack formals (this can happen when a dynamic call passes too many parameters).
			: aParamCount;
		
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
		if (mInstances > 0) // i.e. treat negatives as zero to help catch any bugs in the way mInstances is maintained.
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
			for (j = 0; j < aParamCount; ++j) // For each actual parameter.
			{
				ExprTokenType &this_param_token = *aParam[j]; // stack[stack_count] is the first actual parameter. A check higher above has already ensured that this line won't cause stack overflow.
				if (this_param_token.symbol == SYM_VAR && !(j < mParamCount && mParam[j].is_byref))
				{
					// Since this formal parameter is passed by value, if it's SYM_VAR, convert it to
					// a non-var to allow the variables to be backed up and reset further below without
					// corrupting any SYM_VARs that happen to be locals or params of this very same
					// function.
					// DllCall() relies on the fact that this transformation is only done for user
					// functions, not built-in ones such as DllCall().  This is because DllCall()
					// sometimes needs the variable of a parameter for use as an output parameter.
					// Skip AddRef() if this is an object because Release() won't be called, and
					// AddRef() will be called when the object is assigned to a parameter.
					this_param_token.var->ToTokenSkipAddRef(this_param_token);
				}
			}
			// BackupFunctionVars() will also clear each local variable and formal parameter so that
			// if that parameter or local var is assigned a value by any other means during our call
			// to it, new memory will be allocated to hold that value rather than overwriting the
			// underlying recursed/interrupted instance's memory, which it will need intact when it's resumed.
			if (!Var::BackupFunctionVars(*this, backup, backup_count)) // Out of memory.
			{
				aResultToken.Error(ERR_OUTOFMEM, mName);
				return false;
			}
		} // if (func.mInstances > 0)
		//else backup is not needed because there are no other instances of this function on the call-stack.
		// So by definition, this function is not calling itself directly or indirectly, therefore there's no
		// need to do the conversion of SYM_VAR because those SYM_VARs can't be ones that were blanked out
		// due to a function exiting.  In other words, it seems impossible for a there to be no other
		// instances of this function on the call-stack and yet SYM_VAR to be one of this function's own
		// locals or formal params because it would have no legitimate origin.

		// From this point on, mInstances must be decremented before returning, even on error:
		++mInstances;

		for (j = 0; j < mParamCount; ++j) // For each formal parameter.
		{
			FuncParam &this_formal_param = mParam[j]; // For performance and convenience.

			if (j >= aParamCount || aParam[j]->symbol == SYM_MISSING)
			{
				if (this_formal_param.is_byref) // v1.0.46.13: Allow ByRef parameters to be optional by converting an omitted-actual into a non-alias formal/local.
					this_formal_param.var->ConvertToNonAliasIfNecessary(); // Convert from alias-to-normal, if necessary.

				if (param_obj)
				{
					ExprTokenType named_value;
					if (param_obj->GetItem(named_value, this_formal_param.var->mName))
					{
						this_formal_param.var->Assign(named_value);
						continue;
					}
				}
			
				switch(this_formal_param.default_type)
				{
				case PARAM_DEFAULT_STR:   this_formal_param.var->Assign(this_formal_param.default_str);    break;
				case PARAM_DEFAULT_INT:   this_formal_param.var->Assign(this_formal_param.default_int64);  break;
				case PARAM_DEFAULT_FLOAT: this_formal_param.var->Assign(this_formal_param.default_double); break;
				default: //case PARAM_DEFAULT_NONE:
					// No value has been supplied for this REQUIRED parameter.
					aResultToken.Error(ERR_PARAM_REQUIRED, this_formal_param.var->mName); // Abort thread.
					goto free_and_return;
				}
				continue;
			}

			ExprTokenType &token = *aParam[j];
			
			if (this_formal_param.is_byref)
			{
				// Note that the previous loop might not have checked things like the following because that
				// loop never ran unless a backup was needed:
				if (token.symbol != SYM_VAR)
				{
					// L60: Seems more useful and in the spirit of AutoHotkey to allow ByRef parameters
					// to act like regular parameters when no var was specified.  If we force script
					// authors to pass a variable, they may pass a temporary variable which is then
					// discarded, adding a little overhead and impacting the readability of the script.
					this_formal_param.var->ConvertToNonAliasIfNecessary();
				}
				else
				{
					this_formal_param.var->UpdateAlias(token.var); // Make the formal parameter point directly to the actual parameter's contents.
					continue;
				}
			}
			//else // This parameter is passed "by value".
			// Assign actual parameter's value to the formal parameter (which is itself a
			// local variable in the function).  
			// token.var's Type() is always VAR_NORMAL (e.g. never the clipboard).
			// A SYM_VAR token can still happen because the previous loop's conversion of all
			// by-value SYM_VAR operands into the appropriate operand symbol would not have
			// happened if no backup was needed for this function (which is usually the case).
			if (!this_formal_param.var->Assign(token))
			{
				aResultToken.SetExitResult(FAIL); // Abort thread.
				goto free_and_return;
			}
		} // for each formal parameter.
		
		if (mIsVariadic) // i.e. this function is capable of accepting excess params via an object/array.
		{
			// If the caller supplied an array of parameters, copy any key-value pairs with non-numbered keys;
			// otherwise, just create a new object.  Either way, numbered params will be inserted below.
			Object *vararg_obj = param_obj ? param_obj->Clone(true) : Object::Create();
			if (!vararg_obj)
			{
				aResultToken.Error(ERR_OUTOFMEM, mName); // Abort thread.
				goto free_and_return;
			}
			if (j < aParamCount)
				// Insert the excess parameters from the actual parameter list.
				vararg_obj->InsertAt(0, 1, aParam + j, aParamCount - j);
			// Assign to the "param*" var:
			mParam[mParamCount].var->AssignSkipAddRef(vararg_obj);
		}

		result = Call(&aResultToken); // Call the UDF.
		
		// Setting this unconditionally isn't likely to perform any worse than checking for EXIT/FAIL,
		// and likely produces smaller code.  Currently EARLY_RETURN results are possible and must be
		// passed back in case this is a meta-function, but this should be revised at some point since
		// most of our callers only expect OK, FAIL or EARLY_EXIT.
		aResultToken.SetResult(result);

free_and_return:
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
		Var::FreeAndRestoreFunctionVars(*this, backup, backup_count);

		// mInstances must remain non-zero until this point to ensure that any recursive calls by an
		// object's __Delete meta-function receive fresh variables, and none partially-destructed.
		--mInstances;
	}
	return !aResultToken.Exited(); // i.e. aResultToken.SetExitResult() or aResultToken.Error() was not called.
}


bool Func::Call(ResultToken &aResultToken, int aParamCount, ...)
{
	ASSERT(aParamCount >= 0);

	// Allocate and build the parameter array
	ExprTokenType **aParam = (ExprTokenType**) _alloca(aParamCount*sizeof(ExprTokenType*));
	ExprTokenType *args    = (ExprTokenType*)  _alloca(aParamCount*sizeof(ExprTokenType));
	va_list va;
	va_start(va, aParamCount);
	for (int i = 0; i < aParamCount; i ++)
	{
		ExprTokenType &cur_arg = args[i];
		aParam[i] = &cur_arg;

		// Initialize the argument structure
		cur_arg.symbol = va_arg(va, SymbolType);

		// Fill in the argument value
		switch (cur_arg.symbol)
		{
			case SYM_INTEGER: cur_arg.value_int64  = va_arg(va, __int64);  break;
			case SYM_STRING:
				cur_arg.marker = va_arg(va, LPTSTR);
				cur_arg.marker_length = -1;
				break;
			case SYM_FLOAT:   cur_arg.value_double = va_arg(va, double);   break;
			case SYM_OBJECT:  cur_arg.object       = va_arg(va, IObject*); break;
		}
	}
	va_end(va);

	// Perform function call
	return Call(aResultToken, aParam, aParamCount);
}



ResultType Line::ExpandArgs(ResultToken *aResultTokens)
// Caller should either provide both or omit both of the parameters.  If provided, it means
// caller already called GetExpandedArgSize for us.
// Returns OK, FAIL, or EARLY_EXIT.  EARLY_EXIT occurs when a function-call inside an expression
// used the EXIT command to terminate the thread.
{
	// The counterparts of sArgDeref and sArgVar kept on our stack to protect them from recursion caused by
	// the calling of functions in the script:
	LPTSTR arg_deref[MAX_ARGS];
	Var *arg_var[MAX_ARGS];
	int i;

	// Make two passes through this line's arg list.  This is done because the performance of
	// realloc() is worse than doing a free() and malloc() because the former often does a memcpy()
	// in addition to the latter's steps.  In addition, realloc() as much as doubles the memory
	// load on the system during the brief time that both the old and the new blocks of memory exist.
	// First pass: determine how much space will be needed to do all the args and allocate
	// more memory if needed.  Second pass: dereference the args into the buffer.

	// First pass. It takes into account the same things as 2nd pass.
	size_t space_needed = GetExpandedArgSize(arg_var);
	if (space_needed == VARSIZE_ERROR)
		return FAIL;  // It will have already displayed the error.

	// Only allocate the buf at the last possible moment, when it's sure the buffer will be used
	// (improves performance when only a short script with no derefs is being run):
	if (space_needed > sDerefBufSize)
	{
		// KNOWN LIMITATION: The memory utilization of *recursive* user-defined functions is rather high because
		// of the size of DEREF_BUF_EXPAND_INCREMENT, which is used to create a new deref buffer for each layer
		// of recursion.  Due to limited stack space, the limit of recursion is about 300 to 800 layers depending
		// on the build.  For 800 layers on Unicode, about 25MB (32KB*800) of memory would be temporarily allocated,
		// which in a worst-case scenario would cause swapping and kill performance.  However, on most systems it
		// wouldn't be an issue, and the bigger problem is that recursion may be limited to ~300 layers.
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
		if (   !(sDerefBuf = tmalloc(new_buf_size))   )
		{
			// Error msg was formerly: "Ran out of memory while attempting to dereference this line's parameters."
			sDerefBufSize = 0;  // Reset so that it can make another attempt, possibly smaller, next time.
			return LineError(ERR_OUTOFMEM); // Short msg since so rare.
		}
		sDerefBufSize = new_buf_size;
		if (sDerefBufSize > LARGE_DEREF_BUF_SIZE)
			++sLargeDerefBufs;
	}

	// Always init our_buf_marker even if zero iterations, because we want to enforce
	// the fact that its prior contents become invalid once we're called.
	// It's also necessary due to the fact that all the old memory is discarded by
	// the above if more space was needed to accommodate this line.
	LPTSTR our_buf_marker = sDerefBuf;  // Prior contents of buffer will be overwritten in any case.

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
	PRIVATIZE_S_DEREF_BUF;

	ResultType result, result_to_return = OK;  // Set default return value.
	Var *the_only_var_of_this_arg;

	if (!mArgc)            // v1.0.45: Required by some commands that can have zero parameters (such as Random and
		sArgVar[0] = NULL; // PixelSearch), even if it's just to allow their output-var(s) to be omitted.  This allows OUTPUT_VAR to be used without any need to check mArgC.
	else
	{
		size_t extra_size = our_deref_buf_size - space_needed;
		for (i = 0; i < mArgc; ++i) // For each arg:
		{
			ArgStruct &this_arg = mArg[i]; // For performance and convenience.

			// Load-time routines have already ensured that an arg can be an expression only if
			// it's not an input or output var.
			if (this_arg.is_expression)
			{
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
				arg_deref[i] = ExpandExpression(i, result, aResultTokens ? &aResultTokens[i] : NULL
					, our_buf_marker, our_deref_buf, our_deref_buf_size, arg_deref, extra_size, arg_var);
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
				if (this_arg.type == ARG_TYPE_OUTPUT_VAR)
				{
					if (!arg_var[i])
					{
						// This arg contains an expression which failed to produce a writable variable.
						// The error message is vague enough to cover actual read-only vars and other
						// expressions which don't produce any kind of variable.
						LineError(ERR_VAR_IS_READONLY, FAIL, arg_deref[i]);
						result_to_return = FAIL;
						goto end;
					}
					//else arg_var[i] has been set to a writable variable by the double-deref code.
				}
				continue;
			}

			if (this_arg.type == ARG_TYPE_OUTPUT_VAR)  // Don't bother wasting the mem to deref output var.
			{
				// In case its "dereferenced" contents are ever directly examined, set it to be
				// the empty string.  This also allows the ARG to be passed a dummy param, which
				// makes things more convenient and maintainable in other places:
				arg_deref[i] = _T("");
				continue;
			}

			// arg_var[i] was previously set by GetExpandedArgSize() or ExpandExpression() above.
			if (   !(the_only_var_of_this_arg = arg_var[i])   )
			{
				if (aResultTokens && this_arg.postfix)
				{
					// Since above did not "continue", this arg must have been an expression which was
					// converted back to a plain value.  *postfix is a single numeric or string literal.
					aResultTokens[i].CopyValueFrom(*this_arg.postfix);
				}
				// Since above did not "continue" and arg_var[i] is NULL, this arg can't be an expression
				// or input/output var and must therefore be plain text.
				arg_deref[i] = this_arg.text;  // Point the dereferenced arg to the arg text itself.
				continue;  // Don't need to use the deref buffer in this case.
			}
			// Since above didn't continue, this arg is a plain variable reference.
			// Even if aResultTokens != NULL, it isn't set because our callers handle vars
			// in different ways (and checking sArgVar is easier than checking for SYM_VAR).

			switch(ArgMustBeDereferenced(the_only_var_of_this_arg, i, arg_var)) // Yes, it was called by GetExpandedArgSize() too, but a review shows it's difficult to avoid this without being worse than the disease (10/22/2006).
			{
			case CONDITION_FALSE:
				// This arg contains only a single dereference variable, and no
				// other text at all.  So rather than copy the contents into the
				// temp buffer, it's much better for performance (especially for
				// potentially huge variables like %clipboard%) to simply set
				// the pointer to be the variable itself.  However, this can only
				// be done if the var is the clipboard or a normal var.
				// Update: Changed it so that it will deref the clipboard if it contains only
				// files and no text, so that the files will be transcribed into the deref buffer.
				// This is because the clipboard object needs a memory area into which to write
				// the filespecs it translated:
				// Update #2: When possible, avoid calling Contents() because that flushes the
				// cached binary number, which some commands don't need to happen. Only the args that
				// are specifically written to be optimized should skip it.  Otherwise there would be
				// problems in things like: date += 31, %Var% (where Var contains "Days")
				// Update #3: If an expression in an arg after this one causes the var's contents
				// to be reallocated, it would invalidate any pointer we could get from Contents()
				// in this iteration.  So instead of calling Contents() here, store a NULL value
				// as a special indicator for the loop below to call Contents().
				arg_deref[i] = // The following is ordered for short-circuit performance:
					(   mActionType == ACT_ASSIGNEXPR && i == 1  // By contrast, for the below i==anything (all args):
					||  mActionType == ACT_IF
					//|| mActionType == ACT_WHILE // Not necessary to check this one because loadtime leaves ACT_WHILE as an expression in all common cases.
					) && the_only_var_of_this_arg->Type() == VAR_NORMAL // Otherwise, users of this optimization would have to reproduce more of the logic in ArgMustBeDereferenced().
					? _T("") : NULL; // See "Update #2" and later comments above.
				break;
			case CONDITION_TRUE:
				// the_only_var_of_this_arg is either a reserved var or a normal var which is either
				// numeric or is used again in this line as an output variable.  In all these cases,
				// it must be expanded into the buffer rather than accessed directly:
				arg_deref[i] = our_buf_marker; // Point it to its location in the buffer.
				our_buf_marker += the_only_var_of_this_arg->Get(our_buf_marker) + 1; // +1 for terminator.
				break;
			default: // FAIL should be the only other possibility.
				result_to_return = FAIL; // ArgMustBeDereferenced() will already have displayed the error.
				goto end;
			}
		} // for each arg.

		// See "Update #3" comment above.  This must be done separately to the loop below since Contents()
		// may cause a warning dialog, which in turn may cause a new thread to launch, thus potentially
		// corrupting sArgDeref/sArgVar.
		for (i = 0; i < mArgc; ++i)
			if (arg_deref[i] == NULL)
				arg_deref[i] = arg_var[i]->Contents();

		// IT'S NOT SAFE to do the following until the above loops FULLY complete because any calls made above to
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
	for (i = mArgc; i < max_params; ++i) // START AT mArgc.  For performance, this only does the actual max args for THIS command, not MAX_ARGS.
		sArgDeref[i] = _T("");
		// But sArgVar isn't done (since it's more rarely used) except sArgVar[0] = NULL higher above.
		// Therefore, users of sArgVar must check mArgC if they have any doubt how many args are present in
		// the script line (this is now enforced via macros).

	// When the main/large loop above ends normally, it falls into the label below and uses the original/default
	// value of "result_to_return".

end:
	// As of v1.0.31, there can be multiple deref buffers simultaneously if one or more called functions
	// requires a deref buffer of its own (separate from ours).  In addition, if a called function is
	// interrupted by a new thread before it finishes, the interrupting thread will also use the
	// new/separate deref buffer.  To minimize the amount of memory used in such cases,
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
	// Must always restore the original buffer (if there was one), not keep the new one, because our
	// caller needs the arg_deref addresses, which point into the original buffer.
	DEPRIVATIZE_S_DEREF_BUF;

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
	// It could be argued that the timer should only be activated when a hypothetical static
	// var sLayers that we maintain here indicates that we're the only layer.  However, if that
	// were done and the launch of a script function creates (directly or through thread
	// interruption, indirectly) a large deref buffer, and that thread is waiting for something
	// such as WinWait, that large deref buffer would never get freed.
	if (sDerefBufSize > LARGE_DEREF_BUF_SIZE)
		SET_DEREF_TIMER(10000) // Reset the timer right before the deref buf is possibly about to become idle.

	if (aResultTokens && result_to_return != OK)
	{
		// For maintainability, release any objects here.  Caller was responsible for ensuring
		// aResultTokens is initialized and has at least mArgc elements.  Caller must assume
		// contents of aResultTokens are invalid if we return != OK.
		for (int i = 0; i < mArgc; ++i)
			aResultTokens[i].Free();
	}

	return result_to_return;
}

	

VarSizeType Line::GetExpandedArgSize(Var *aArgVar[])
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
		
		aArgVar[i] = NULL; // Set default.

		if (this_arg.is_expression)
		{
			// The length used below is more room than is strictly necessary, but given how little
			// space is typically wasted (and that only while the expression is being evaluated),
			// it doesn't seem worth worrying about it.  See other comments at macro definition.
			space_needed += EXPR_BUF_SIZE(this_arg.length);
			continue;
		}
		// Since is_expression is false, it must be plain text or a non-dynamic input/output var.
		
		if (this_arg.type == ARG_TYPE_OUTPUT_VAR)
		{
			// Pre-resolved output vars should never be included in the space calculation,
			// but we do need to store the var reference in aArgVar for our caller.
			ASSERT(!*this_arg.text);
			aArgVar[i] = VAR(this_arg);
			continue;
		}

		if (this_arg.type == ARG_TYPE_INPUT_VAR)
		{
			ASSERT(!*this_arg.text);
			the_only_var_of_this_arg = VAR(this_arg);
			aArgVar[i] = the_only_var_of_this_arg; // For now, this is done regardless of whether it must be dereferenced.
			if (   !(result = ArgMustBeDereferenced(the_only_var_of_this_arg, i, aArgVar))   )
				return VARSIZE_ERROR;
			if (result == CONDITION_FALSE)
				continue;
			//else the size of this arg is always included, so fall through to below.
			//else caller wanted it's size unconditionally included, so continue on to below.
			space_needed += the_only_var_of_this_arg->Get() + 1;  // +1 for the zero terminator.
			// NOTE: Get() (with no params) can retrieve a size larger that what winds up actually
			// being needed, so our callers should be aware that that can happen.
			continue;
		}
	}

	return space_needed;
}



ResultType Line::ArgMustBeDereferenced(Var *aVar, int aArgIndex, Var *aArgVar[]) // 10/22/2006: __forceinline didn't help enough to be worth the added code size of having two instances.
// Shouldn't be called only for args of type ARG_TYPE_OUTPUT_VAR because they never need to be dereferenced.
// aArgVar[] is used for performance; it's assumed to contain valid items only up to aArgIndex, not beyond
// (since normally output vars lie to the left of all input vars, so it doesn't seem worth doing anything
// more complicated).
// Returns CONDITION_TRUE, CONDITION_FALSE, or FAIL.
// There are some other functions like ArgLength() that have procedures similar to this one, so
// maintain them together.
{
	aVar = aVar->ResolveAlias(); // Helps performance, but also necessary to accurately detect a match further below.
	VarTypeType aVar_type = aVar->Type();
	if (aVar_type == VAR_CLIPBOARD)
		// Even if the clipboard is both an input and an output var, it still
		// doesn't need to be dereferenced into the temp buffer because the
		// clipboard has two buffers of its own.  The only exception is when
		// the clipboard has only files on it, in which case those files need
		// to be converted into plain text:
		return CLIPBOARD_CONTAINS_ONLY_FILES ? CONDITION_TRUE : CONDITION_FALSE;
	if (aVar_type != VAR_NORMAL || aVar == g_ErrorLevel)
		// Reserved vars must always be dereferenced due to their volatile nature.
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
	if (!g_act[mActionType].CheckOverlap) // Commands that have this flag don't need final check
		return CONDITION_FALSE;           // further below (though they do need the ones above).

	// Since the above didn't return, we know that this is a NORMAL input var that isn't an
	// environment variable.  Such input vars only need to be dereferenced if they are also
	// used as an output var by the current script line:
	Var *output_var;
	for (int i = 0; i < mArgc; ++i)
		if (mArg[i].type == ARG_TYPE_OUTPUT_VAR) // Implies i != aArgIndex, since this function is not called for output vars.
		{
			output_var = (i < aArgIndex) ? aArgVar[i] : mArg[i].is_expression ? NULL : VAR(mArg[i]); // aArgVar: See top of this function for comments.
			if (!output_var) // Var hasn't been resolved yet.  To be safe, we must assume deref is required.
				return CONDITION_TRUE;
			if (output_var->ResolveAlias() == aVar)
				return CONDITION_TRUE;
		}
	// Otherwise:
	return CONDITION_FALSE;
}


ResultType Line::ValueIsType(ExprTokenType &aResultToken, ExprTokenType &aValue, LPTSTR aValueStr, ExprTokenType &aType, LPTSTR aTypeStr)
{
	VariableTypeType variable_type = ConvertVariableTypeName(aTypeStr);
	bool if_condition;
	TCHAR *cp;

	if (variable_type == VAR_TYPE_BYREF)
	{
		if (aValue.symbol == SYM_VAR)
		{
			aResultToken.value_int64 = aValue.var->ResolveAlias() != aValue.var;
			return OK;
		}
		// Otherwise, the comparison is invalid.
	}
	else if (IObject *type_obj = TokenToObject(aType))
	{
		// Is the value an object which can derive, and is it derived from type_obj?
		Object *value_obj = dynamic_cast<Object *>(TokenToObject(aValue));
		aResultToken.value_int64 = value_obj && value_obj->IsDerivedFrom(type_obj);
		return OK;
	}
	else if (TokenToObject(aValue))
	{
		// Since it's an object, the only type it should match is "object" (even though aValueStr
		// is an empty string, which matches several other types).
		aResultToken.value_int64 = variable_type == VAR_TYPE_OBJECT;
		return OK;
	}

	// The remainder of this function is based on the original code for ACT_IFIS, which was removed
	// in commit 3382e6e2.
	switch (variable_type)
	{
	case VAR_TYPE_NUMBER:
		if_condition = IsNumeric(aValueStr, true, false, true);
		break;
	case VAR_TYPE_INTEGER:
		if_condition = IsNumeric(aValueStr, true, false, false);  // Passes false for aAllowFloat.
		break;
	case VAR_TYPE_FLOAT:
		if_condition = (IsNumeric(aValueStr, true, false, true) == PURE_FLOAT);
		break;
	case VAR_TYPE_OBJECT:
		// if aValue was an object, it was already handled above.
		if_condition = false;
		break;
	case VAR_TYPE_TIME:
	{
		SYSTEMTIME st;
		// Also insist on numeric, because even though YYYYMMDDToFileTime() will properly convert a
		// non-conformant string such as "2004.4", for future compatibility, we don't want to
		// report that such strings are valid times:
		if_condition = IsNumeric(aValueStr, false, false, false) && YYYYMMDDToSystemTime(aValueStr, st, true); // Can't call Var::IsNumeric() here because it doesn't support aAllowNegative.
		break;
	}
	case VAR_TYPE_DIGIT:
		if_condition = true;
		for (cp = aValueStr; *cp; ++cp)
			if (!_istdigit((UCHAR)*cp))
			{
				if_condition = false;
				break;
			}
		break;
	case VAR_TYPE_XDIGIT:
		cp = aValueStr;
		if (!_tcsnicmp(cp, _T("0x"), 2)) // Allow 0x prefix.
			cp += 2;
		if_condition = true;
		for (; *cp; ++cp)
			if (!_istxdigit((UCHAR)*cp))
			{
				if_condition = false;
				break;
			}
		break;
	case VAR_TYPE_ALNUM:
		if_condition = true;
		for (cp = aValueStr; *cp; ++cp)
			//if (!IsCharAlphaNumeric(*cp)) // Use this to better support chars from non-English languages.
			if (!aisalnum(*cp)) // But some users don't like it, Chinese users for example.
			{
				if_condition = false;
				break;
			}
		break;
	case VAR_TYPE_ALPHA:
		// Like AutoIt3, the empty string is considered to be alphabetic, which is only slightly debatable.
		if_condition = true;
		for (cp = aValueStr; *cp; ++cp)
			//if (!IsCharAlpha(*cp)) // Use this to better support chars from non-English languages.
			if (!aisalpha(*cp)) // But some users don't like it, Chinese users for example.
			{
				if_condition = false;
				break;
			}
		break;
	case VAR_TYPE_UPPER:
		if_condition = true;
		for (cp = aValueStr; *cp; ++cp)
			//if (!IsCharUpper(*cp)) // Use this to better support chars from non-English languages.
			if (!aisupper(*cp)) // But some users don't like it, Chinese users for example.
			{
				if_condition = false;
				break;
			}
		break;
	case VAR_TYPE_LOWER:
		if_condition = true;
		for (cp = aValueStr; *cp; ++cp)
			//if (!IsCharLower(*cp)) // Use this to better support chars from non-English languages.
			if (!aislower(*cp)) // But some users don't like it, Chinese users for example.
			{
				if_condition = false;
				break;
			}
		break;
	case VAR_TYPE_SPACE:
		if_condition = true;
		for (cp = aValueStr; *cp; ++cp)
			if (!_istspace(*cp))
			{
				if_condition = false;
				break;
			}
		break;
	default:
		return LineError(_T("Unsupported comparison type."), FAIL, aTypeStr);
	}
	aResultToken.value_int64 = if_condition;
	return OK;
}

