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
LPTSTR Line::ExpandExpression(int aArgIndex, ResultType &aResult, ExprTokenType *aResultToken
		, LPTSTR &aTarget, LPTSTR &aDerefBuf, size_t &aDerefBufSize, LPTSTR aArgDeref[], size_t aExtraSize)
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

	// The following must be defined early so that mem_count is initialized and the array is guaranteed to be
	// "in scope" in case of early "goto" (goto substantially boosts performance and reduces code size here).
	#define MAX_EXPR_MEM_ITEMS 200 // v1.0.47.01: Raised from 100 because a line consisting entirely of concat operators can exceed it.  However, there's probably not much point to going much above MAX_TOKENS/2 because then it would reach the MAX_TOKENS limit first.
	LPTSTR mem[MAX_EXPR_MEM_ITEMS]; // No init necessary.  In most cases, it will never be used.
	int mem_count = 0; // The actual number of items in use in the above array.
	LPTSTR result_to_return = _T(""); // By contrast, NULL is used to tell the caller to abort the current thread.  That isn't done for normal syntax errors, just critical conditions such as out-of-memory.
	Var *output_var = (mActionType == ACT_ASSIGNEXPR) ? OUTPUT_VAR : NULL; // Resolve early because it's similar in usage/scope to the above.  Plus MUST be resolved prior to calling any script-functions since they could change the values in sArgVar[].

	ExprTokenType *stack[MAX_TOKENS];
	int stack_count = 0, high_water_mark = 0; // L31: high_water_mark is used to simplify object management.
	ExprTokenType *&postfix = mArg[aArgIndex].postfix;

	DerefType *deref;
	LPTSTR cp;

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
		this_token = *this_postfix; // Struct copy. See comment section above.

		// At this stage, operands in the postfix array should be SYM_STRING, SYM_INTEGER, SYM_FLOAT or SYM_DYNAMIC.
		// But all are checked since that operation is just as fast:
		if (IS_OPERAND(this_token.symbol)) // If it's an operand, just push it onto stack for use by an operator in a future iteration.
		{
			if (this_token.symbol == SYM_DYNAMIC) // CONVERTED HERE/EARLY TO SOMETHING *OTHER* THAN SYM_DYNAMIC so that no later stages need any handling for them as operands. SYM_DYNAMIC is quite similar to SYM_FUNC/BIF in this respect.
			{
				if (SYM_DYNAMIC_IS_DOUBLE_DEREF(this_token)) // Double-deref such as Array%i%.
				{
					// Start off by looking for the first deref.
					deref = (DerefType *)this_token.var; // MUST BE DONE PRIOR TO OVERWRITING MARKER/UNION BELOW.
					cp = this_token.buf; // Start at the beginning of this arg's text.
					size_t var_name_length = 0;
					
					this_token.marker = _T("");         // Set default in case of early goto.  Must be done after above.
					this_token.symbol = SYM_STRING; //

					if (deref->marker == cp && !cp[deref->length] && (deref+1)->is_function // %soleDeref%()
						&& (deref->var->HasObject() // It's an object.
						|| !deref->var->HasContents())) // It's an empty string (which may be passed to __Call).
					{
						// This dynamic function call deref consists of a single var containing an object
						// or an empty string; both need special handling.
						if (deref->var->HasObject())
						{
							// Push the object, not the var, in case this variable is modified in the
							// parameter list of this function call (which is evaluated prior to SYM_FUNC).
							this_token.symbol = SYM_OBJECT;
							this_token.object = deref->var->Object();
							this_token.object->AddRef();
						}
						// Otherwise, leave it set to the empty string.
						goto push_this_token;
					}
					// Otherwise, it could still be %var%() where var contains a string which is too
					// long for the section below to handle.  Supporting that case doesn't seem
					// worthwhile for the following reasons:
					//  1) It would require some logic in the above section to allocate memory for the
					//     string, keeping in mind that deref->var could be VAR_CLIPBOARD/BUILTIN which
					//     can't be pushed directly onto the stack.
					//  2) Pushing a var onto the stack would cause inconsistent behaviour if that var
					//     is modified within the function call's parameter list (although that already
					//     causes inconsistent behaviour if the var appears multiple times in the list).
					//  3) Even if the %var%() case was supported, we don't know at this point whether
					//     xxx%var%() or similar will exceed MAX_VAR_NAME_LENGTH, which means one more
					//     inconsistency.
					//  4) The only benefit of supporting long strings is consistency with regard to
					//     __Call, which will probably only be used for debugging.  Future changes to
					//     error-handling in expressions could supersede that.

					// Loadtime validation has ensured that none of these derefs are function-calls
					// (i.e. deref->is_function is alway false) with the possible exception of the final
					// deref (the one with a NULL marker), which can be a function-call if this
					// SYM_DYNAMIC is a dynamic call. Other than this, loadtime logic seems incapable
					// of producing function-derefs inside something that would later be interpreted
					// as a double-deref.
					for (; deref->marker; ++deref)  // A deref with a NULL marker terminates the list. "deref" was initialized higher above.
					{
						// FOR EACH DEREF IN AN ARG (if we're here, there's at least one):
						// Copy the chars that occur prior to deref->marker into the buffer:
						for (; cp < deref->marker && var_name_length < MAX_VAR_NAME_LENGTH; left_buf[var_name_length++] = *cp++);
						if (var_name_length >= MAX_VAR_NAME_LENGTH && cp < deref->marker) // The variable name would be too long!
							goto double_deref_fail; // For simplicity and in keeping with the tradition that expressions generally don't display runtime errors, just treat it as a blank.
						// Now copy the contents of the dereferenced var.  For all cases, aBuf has already
						// been verified to be large enough, assuming the value hasn't changed between the
						// time we were called and the time the caller calculated the space needed.
						if (deref->var->Get() > (VarSizeType)(MAX_VAR_NAME_LENGTH - var_name_length)) // The variable name would be too long!
							goto double_deref_fail; // For simplicity and in keeping with the tradition that expressions generally don't display runtime errors, just treat it as a blank.
						var_name_length += deref->var->Get(left_buf + var_name_length);
						// Finally, jump over the dereference text:
						cp += deref->length;
					}

					// Copy any chars that occur after the final deref into the buffer:
					for (; *cp && var_name_length < MAX_VAR_NAME_LENGTH; left_buf[var_name_length++] = *cp++);
					if (var_name_length >= MAX_VAR_NAME_LENGTH && *cp // The variable name would be too long!
						|| !var_name_length) // It resolves to an empty string (e.g. a simple dynamic var like %Var% where Var is blank).
						goto double_deref_fail; // For simplicity and in keeping with the tradition that expressions generally don't display runtime errors, just treat it as a blank.

					// Terminate the buffer, even if nothing was written into it:
					left_buf[var_name_length] = '\0';

					// v1.0.47.06 dynamic function calls: As a result of a prior loop, deref = the null-marker
					// deref which terminates the deref list. is_function is set by the infix processing code.
					if (deref->is_function)
					{	
						// Since the parameter list is evaluated between SYM_DYNAMIC and its corresponding
						// SYM_FUNC, it is possible for a recursive script to re-evaluate this token a second
						// time prior to actually calling the function the first time.  With the old approach
						// of setting the SYM_FUNC's deref->func in this section, such recursion could cause
						// the wrong function to be called.  Instead, push the function name onto the stack
						// to await SYM_FUNC:
						this_token.symbol = SYM_STRING;
						this_token.marker = _tcscpy(talloca(var_name_length + 1), left_buf);
						goto push_this_token;
					}

					// In v1.0.31, FindOrAddVar() vs. FindVar() is called below to support the passing of non-existent
					// array elements ByRef, e.g. Var:=MyFunc(Array%i%) where the MyFunc function's parameter is
					// defined as ByRef, would effectively create the new element Array%i% if it doesn't already exist.
					// Since at this stage we don't know whether this particular double deref is to be sent as a param
					// to a function, or whether it will be byref, this is done unconditionally for all double derefs
					// since it seems relatively harmless to create a blank variable in something like var := Array%i%
					// (though it will produce a runtime error if the double resolves to an illegal variable name such
					// as one containing spaces).
					if (   !(temp_var = g_script.FindOrAddVar(left_buf, var_name_length))   )
					{
						// Above already displayed the error.  As of v1.0.31, this type of error is displayed and
						// causes the current thread to terminate, which seems more useful than the old behavior
						// that tolerated anything in expressions.
						goto abort;
					}
					// Otherwise, var was found or created.
					this_token.var = temp_var;
				} // Double-deref.
				//else: It's a built-in variable or potential environment variable.

				// Check if it's a normal variable rather than a built-in or environment variable.
				// This happens when g_NoEnv==FALSE.
				switch (this_token.var->Type())
  				{
				case VAR_NORMAL:
					this_token.symbol = SYM_VAR; // The fact that a SYM_VAR operand is always VAR_NORMAL (with one limited exception) is relied upon in several places such as built-in functions.
					goto push_this_token;
				case VAR_BUILTIN: // v1.0.48.02: Ensure it's VAR_BUILTIN prior to below because mBIV is a union with mCapacity.
					if (this_token.var->mBIV == BIV_LoopIndex) // v1.0.48.01: Improve performance of A_Index by treating it as an integer rather than a string in expressions (avoids conversions to/from strings).
					{
						this_token.value_int64 = g->mLoopIteration;
						this_token.symbol = SYM_INTEGER;
						goto push_this_token;
					}
					if (this_token.var->mBIV == BIV_EventInfo) // v1.0.48.02: A_EventInfo is used often enough in performance-sensitive numeric contexts to seem worth special treatment like A_Index; e.g. LV_GetText(RowText, A_EventInfo) or RegisterCallback()'s A_EventInfo.
					{
						this_token.value_int64 = g->EventInfo;
						this_token.symbol = SYM_INTEGER;
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
					break; // case VAR_BUILTIN
  				}
				// Otherwise, it's a built-in variable.
				result_size = this_token.var->Get() + 1;
				if (result_size == 1)
				{
					this_token.marker = _T("");
					this_token.symbol = SYM_STRING;
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
					if (mem_count == MAX_EXPR_MEM_ITEMS // No more slots left (should be nearly impossible).
						|| !(mem[mem_count] = tmalloc(result_size)))
					{
						LineError(ERR_OUTOFMEM, FAIL, this_token.var->mName);
						goto abort;
					}
					// Point result to its new, more persistent location:
					result = mem[mem_count];
					++mem_count; // Must be done last.
				}
				this_token.var->Get(result); // MUST USE "result" TO AVOID OVERWRITING MARKER/VAR UNION.
				this_token.marker = result;  // Must be done after above because marker and var overlap in union.
				this_token.symbol = SYM_STRING;
			} // if (this_token.symbol == SYM_DYNAMIC)
			goto push_this_token;
		} // if (IS_OPERAND(this_token.symbol))

		if (this_token.symbol == SYM_FUNC) // A call to a function (either built-in or defined by the script).
		{
			Func *func = this_token.deref->func;
			actual_param_count = this_token.deref->param_count; // For performance.
			if (actual_param_count > stack_count) // Prevent stack underflow (probably impossible if actual_param_count is accurate).
				goto abnormal_end;
			// Adjust the stack early to simplify.  Above already confirmed that the following won't underflow.
			// Pop the actual number of params involved in this function-call off the stack.
			stack_count -= actual_param_count; // Now stack[stack_count] is the leftmost item in an array of function-parameters, which simplifies processing later on.
			ExprTokenType **params = stack + stack_count;

			if (!func)
			{
				// This is a dynamic function call.
				if (!stack_count) // SYM_DYNAMIC should have pushed a function name or reference onto the stack, but a syntax error may still cause this condition.
					goto abnormal_end;
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
					params[0]->symbol = SYM_STRING;
					params[0]->marker = _T("");
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
				// Traditionally, expressions don't display any runtime errors.  So if the function is being
				// called incorrectly by the script, the expression is aborted like it would be for other
				// syntax errors:
				if (actual_param_count < func->mMinParams && this_token.deref->is_function != DEREF_VARIADIC)
					goto abnormal_end;
			}
			
			// The following two steps are now done inside Func::Call:
			//this_token.symbol = SYM_INTEGER; // Set default return type so that functions don't have to do it if they return INTs.
			//this_token.marker = func.mName;  // Inform function of which built-in function called it (allows code sharing/reduction).
			this_token.buf = left_buf;       // mBIF() can use this to store a string result, and for other purposes.
			
			this_token.mem_to_free = NULL; // Init to detect whether the called function allocates it.
			
			FuncCallData func_call;
			// Call the user-defined or built-in function.  Func::Call takes care of variadic parameter
			// lists and stores local var backups for UDFs in func_call.  Once func_call goes out of scope,
			// its destructor calls Var::FreeAndRestoreFunctionVars() if appropriate.
			if (!func->Call(func_call, aResult, this_token, params, actual_param_count
									, this_token.deref->is_function == DEREF_VARIADIC))
			{
				// Take a shortcut because for backward compatibility, ACT_ASSIGNEXPR (and anything else
				// for that matter) is being aborted by this type of early return (i.e. if there's an
				// output_var, its contents are left as-is).  In other words, this expression will have
				// no result storable by the outside world.
				if (aResult != OK) // i.e. EARLY_EXIT or FAIL
					result_to_return = NULL; // Use NULL to inform our caller that this thread is finished (whether through normal means such as Exit or a critical error).
					// Above: The callers of this function know that the value of aResult (which already contains the
					// reason for early exit) should be considered valid/meaningful only if result_to_return is NULL.
				goto normal_end_skip_output_var; // output_var is left unchanged in these cases.
			}

#ifdef CONFIG_DEBUGGER
			// See PostExecFunctionCall() itself for comments.
			g_Debugger.PostExecFunctionCall(this);
#endif
			g_script.mCurrLine = this; // For error-reporting.

			if (this_token.symbol != SYM_STRING)
			{
				// No need for make_result_persistent or early Assign().  Any numeric or object result can
				// be considered final because it's already stored in permanent memory (the token itself).
				// Additionally, this_token.mem_to_free is assumed to be NULL since the result is not
				// a string; i.e. the function would've had no need to return memory to us.
				goto push_this_token;
			}
			
			#define EXPR_IS_DONE (!stack_count && this_postfix[1].symbol == SYM_INVALID) // True if we've used up the last of the operators & operands.
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
					if (this_token.mem_to_free)
						free(this_token.mem_to_free); // Don't bother putting it into mem[].
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
			
			// RELIES ON THE IS_NUMERIC() CHECK above having been done first.
			result = this_token.marker; // Marker can be used because symbol will never be SYM_VAR in this case.

			if (internal_output_var
				&& !(g->CurrentFunc == func && internal_output_var->IsNonStaticLocal())) // Ordered for short-circuit performance.
				// Above line is a fix for v1.0.45.03: It detects whether output_var is among the variables
				// that are about to be restored from backup.  If it is, we can't assign to it now
				// because it's currently a local that belongs to the instance we're in the middle of
				// calling; i.e. it doesn't belong to our instance (which is beneath it on the call stack
				// until after the restore-from-backup is done later below).  And we can't assign "result"
				// to it *after* the restore because by then result may have been freed (if it happens to be
				// a local variable too).  Therefore, continue on to the normal method, which will check
				// whether "result" needs to be stored in more persistent memory.
			{
				// Check if the called function allocated some memory for its result and turned it over to us.
				// In most cases, the string stored in mem_to_free (if it has been set) is the same address as
				// this_token.marker (i.e. what is named "result" further below), because that's what the
				// built-in functions are normally using the memory for.
				if (this_token.mem_to_free == this_token.marker) // marker is checked in case caller alloc'd mem but didn't use it as its actual result.  Relies on the fact that marker is never NULL.
				{
					// So now, turn over responsibility for this memory to the variable. The called function
					// is responsible for having stored the length of what's in the memory as an overload of
					// this_token.buf, but only when that memory is the result (currently might always be true).
					// AcceptNewMem() will shrink the memory for us, via _expand(), if there's a lot of
					// extra/unused space in it.
					internal_output_var->AcceptNewMem(this_token.mem_to_free, this_token.marker_length);
				}
				else
				{
					result_length = (VarSizeType)_tcslen(result);
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
						if (!internal_output_var->Assign(this_token.marker)) // Marker can be used because symbol will never be SYM_VAR in this case.
							goto abort;										 // ALSO: Assign() contains an optimization that avoids actually doing the mem-copying if output_var is being assigned to itself (which can happen in cases like RegExMatch()).
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

			// RESTORE THE CIRCUIT TOKEN (after handling what came back inside it):
			if (this_token.mem_to_free) // The called function allocated some memory and turned it over to us.
			{
				if (mem_count == MAX_EXPR_MEM_ITEMS) // No more slots left (should be nearly impossible).
				{
					LineError(ERR_OUTOFMEM, FAIL, func->mName);
					goto abort;
				}
				if (this_token.mem_to_free == this_token.marker) // mem_to_free is checked in case caller alloc'd mem but didn't use it as its actual result.
				{
					make_result_persistent = false; // Override the default set higher above.
				}
				// Mark it to be freed at the time we return.
				mem[mem_count++] = this_token.mem_to_free;
			}
			//else this_token.mem_to_free==NULL, so the BIF just called didn't allocate memory to give to us.
			
			// Empty strings are returned pretty often by UDFs, such as when they don't use "return"
			// at all.  Therefore, handle them fully now, which should improve performance (since it
			// avoids all the other checking later on).  It also doesn't hurt code size because this
			// check avoids having to check for empty string in other sections later on.
			if (!*this_token.marker) // Various make-persistent sections further below may rely on this check.
			{
				this_token.marker = _T(""); // Ensure it's a constant memory area, not a buf that might get overwritten soon.
				goto push_this_token; // For code simplicity, the optimization for numeric results is done at a later stage.
			}

			if (func->mIsBuiltIn)
			{
				// Since above didn't goto, "result" is not SYM_INTEGER/FLOAT/VAR, and not "".  Therefore, it's
				// either a pointer to static memory (such as a constant string), or more likely the small buf
				// we gave to the BIF for storing small strings.  For simplicity assume it's the buf, which is
				// volatile and must be made persistent if called for below.
				if (make_result_persistent) // At this stage, this means that the above wasn't able to determine its correct value yet.
					make_result_persistent = !done;
			}
			else // It's not a built-in function.
			{
				// Since above didn't goto:
				// The result just returned may need to be copied to a more persistent location.  This is done right
				// away if the result is the contents of a local variable (since all locals are about to be freed
				// and overwritten), which is assumed to be the case if it's not in the new deref buf because it's
				// difficult to distinguish between when the function returned one of its own local variables
				// rather than a global or a string/numeric literal).  The only exceptions are covered below.
				// Old method, not necessary to be so thorough because "return" always puts its result as the
				// very first item in its deref buf.  So this is commented out in favor of the line below it:
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
				} // This is the end of the section that determines the value of "make_result_persistent" for UDFs.
			}

			if (make_result_persistent) // Both UDFs and built-in functions have ensured make_result_persistent is set.
			{
				// BELOW RELIES ON THE ABOVE ALWAYS HAVING VERIFIED AND FULLY HANDLED RESULT BEING AN EMPTY STRING.
				// So now we know result isn't an empty string, which in turn ensures that size > 1 and length > 0,
				// which might be relied upon by things further below.
				result_size = _tcslen(result) + 1; // No easy way to avoid strlen currently. Maybe some future revisions to architecture will provide a length.
				// Must cast to int to avoid loss of negative values:
				if (result_size <= aDerefBufSize - (target - aDerefBuf)) // There is room at the end of our deref buf, so use it.
				{
					// Make the token's result the new, more persistent location:
					this_token.marker = (LPTSTR)tmemcpy(target, result, result_size); // Benches slightly faster than strcpy().
					target += result_size; // Point it to the location where the next string would be written.
				}
				else if (result_size < EXPR_SMALL_MEM_LIMIT && alloca_usage < EXPR_ALLOCA_LIMIT) // See comments at EXPR_SMALL_MEM_LIMIT.
				{
					this_token.marker = (LPTSTR)tmemcpy((LPTSTR)talloca(result_size), result, result_size); // Benches slightly faster than strcpy().
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
					if (mem_count == MAX_EXPR_MEM_ITEMS // No more slots left (should be nearly impossible).
						|| !(mem[mem_count] = tmalloc(result_size)))
					{
						LineError(ERR_OUTOFMEM, FAIL, func->mName);
						goto abort;
					}
					// Make the token's result the new, more persistent location:
					this_token.marker = (LPTSTR)tmemcpy(mem[mem_count], result, result_size); // Benches slightly faster than strcpy().
					++mem_count; // Must be done last.
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
			goto abnormal_end;
		ExprTokenType &right = *STACK_POP;
		if (!IS_OPERAND(right.symbol)) // Haven't found a way to produce this situation yet, but safe to assume it's possible.
			goto abnormal_end;

		switch (this_token.symbol)
		{
		case SYM_ASSIGN:        // These don't need "right_is_number" to be resolved. v1.0.48.01: Also avoid
		case SYM_CONCAT:        // resolving right_is_number for CONCAT because TokenIsPureNumeric() will take
		case SYM_ASSIGN_CONCAT: // a long time if the string is very long and consists entirely of digits/whitespace.
			right_is_pure_number = right_is_number = PURE_NOT_NUMERIC; // Init for convenience/maintainability.
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
					this_token = right;
					right.symbol = SYM_INTEGER; // i.e if it was SYM_OBJECT, don't call Release() for this copy of the pointer.
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
			this_token.value_int64 = !TokenToBOOL(right);
			this_token.symbol = SYM_INTEGER; // Result is always one or zero.
			break;

		case SYM_NEGATIVE:  // Unary-minus.
			if (right_is_number == PURE_INTEGER)
				this_token.value_int64 = -TokenToInt64(right);
			else if (right_is_number == PURE_FLOAT)
				this_token.value_double = -TokenToDouble(right, FALSE); // Pass FALSE for aCheckForHex since PURE_FLOAT is never hex.
			else // String.
			{
				// Seems best to consider the application of unary minus to a string, even a quoted string
				// literal such as "15", to be a failure.  UPDATE: For v1.0.25.06, invalid operations like
				// this instead treat the operand as an empty string.  This avoids aborting a long, complex
				// expression entirely just because on of its operands is invalid.  However, the net effect
				// in most cases might be the same, since the empty string is a non-numeric result and thus
				// will cause any operator it is involved with to treat its other operand as a string too.
				// And the result of a math operation on two strings is typically an empty string.
				this_token.marker = _T("");
				this_token.symbol = SYM_STRING;
				break;
			}
			// Since above didn't "break":
			this_token.symbol = right_is_number;
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
					right.var->MaybeWarnUninitialized(); // This line should always be reached if the var is uninitialized.
					right.var->Assign(); // If target var contains "" or "non-numeric text", make it blank. Clipboard is also supported here.
					if (is_pre_op)
					{
						// v1.0.46.01: For consistency, it seems best to make the result of a pre-op be a
						// variable whenever a variable came in.  This allows its address to be taken, and it
						// to be passed by reference, and other SYM_VAR behaviors, even if the operation itself
						// produces a blank value.
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
							break;
						}
						//else VAR_CLIPBOARD, which is allowed in only when it's the lvalue of an assignment or
						// inc/dec.  So fall through to make the result blank because clipboard isn't allowed as
						// SYM_VAR beyond this point (to simplify the code and improve maintainability).
					}
					//else post_op against non-numeric target-var.  Fall through to below to yield blank result.
				}
				//else target isn't a var.  Fall through to below to yield blank result.
				this_token.marker = _T("");          // Make the result blank to indicate invalid operation
				this_token.symbol = SYM_STRING;  // (assign to non-lvalue or increment/decrement a non-number).
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
				this_token.symbol = SYM_INTEGER;
				this_token.value_int64 = (__int64)obj;
			}
			else if (right.symbol == SYM_VAR // At this stage, SYM_VAR is always a normal variable, never a built-in one, so taking its address should be safe.
				&& !right.var->IsPureNumeric()) // Seems best not to return Contents() in this case since any changes to it might cause inconsistency.
			{
				this_token.symbol = SYM_INTEGER;
				this_token.value_int64 = (__int64)right.var->Contents(); // Contents() vs. mContents to support VAR_CLIPBOARD, and in case mContents needs to be updated by Contents().
			}
			else // Invalid, so make it a localized blank value.
			{
				this_token.marker = _T("");
				this_token.symbol = SYM_STRING;
			}
			break;

		case SYM_DEREF:   // Dereference an address to retrieve a single byte.
		case SYM_BITNOT:  // The tilde (~) operator.
			if (right_is_number == PURE_NOT_NUMERIC) // String.  Seems best to consider the application of '*' or '~' to a string, even a quoted string literal such as "15", to be a failure.
			{
				this_token.marker = _T("");
				this_token.symbol = SYM_STRING;
				break;
			}
			// Since above didn't "break": right_is_number is PURE_INTEGER or PURE_FLOAT.
			right_int64 = TokenToInt64(right); // Although PURE_FLOAT can't be hex, for simplicity and due to the rarity of encountering a PURE_FLOAT in this case, the slight performance reduction of calling TokenToInt64() is done for both PURE_FLOAT and PURE_INTEGER.
			if (this_token.symbol == SYM_BITNOT)
			{
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
				this_token.value_int64 = ((size_t)right_int64 < 4096)
					? 0 : *(UCHAR *)right_int64; // Dereference to extract one unsigned character, just like Asc().
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
					this_token.marker = _T("");          // Make the result blank to indicate invalid operation
					this_token.symbol = SYM_STRING;  // (assign to non-lvalue).
					break; // Equivalent to "goto push_this_token" in this case.
				}
				switch(this_token.symbol)
				{
				case SYM_ASSIGN: // Listed first for performance (it's probably the most common because things like ++ and += aren't expressions when they're by themselves on a line).
					if (!left.var->Assign(right)) // left.var can be VAR_CLIPBOARD in this case.
						goto abort;
					if (left.var->Type() == VAR_CLIPBOARD) // v1.0.46.01: Clipboard is present as SYM_VAR, but only for assign-to-clipboard so that built-in functions and other code sections don't need handling for VAR_CLIPBOARD.
					{
						this_token = right; // Struct copy.  Doing it this way is more maintainable than other methods, and is unlikely to perform much worse.
						right.symbol = SYM_INTEGER; // i.e if it was SYM_OBJECT (which would be pointless in this case, but could happen), don't call Release() for this copy of the pointer.
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
						this_token.value_int64 = (this_token.symbol != SYM_NOTEQUAL) == (right_obj == left_obj);
						this_token.symbol = SYM_INTEGER; // Must be set *after* above checks symbol.
						goto push_this_token;
					}
				}

				// Above check has ensured that at least one of them is a string.  But the other
				// one might be a number such as in 5+10="15", in which 5+10 would be a numerical
				// result being compared to the raw string literal "15".
				right_string = TokenToString(right, right_buf);
				left_string = TokenToString(left, left_buf);
				result_symbol = SYM_INTEGER; // Set default.  Boolean results are treated as integers.
				switch(this_token.symbol)
				{
				case SYM_EQUAL:     this_token.value_int64 = !((g->StringCaseSense == SCS_INSENSITIVE)
										? _tcsicmp(left_string, right_string)
										: lstrcmpi(left_string, right_string)); break; // i.e. use the "more correct mode" except when explicitly told to use the fast mode (v1.0.43.03).
				case SYM_EQUALCASE: this_token.value_int64 = !_tcscmp(left_string, right_string); break; // Case sensitive.
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
					right_length = (right.symbol == SYM_VAR) ? right.var->LengthIgnoreBinaryClip() : _tcslen(right_string);
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
					left_length = (left.symbol == SYM_VAR) ? left.var->LengthIgnoreBinaryClip() : _tcslen(left_string);
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
						this_token.marker = (LPTSTR)talloca(result_size);
						alloca_usage += result_size; // This might put alloca_usage over the limit by as much as EXPR_SMALL_MEM_LIMIT, but that is fine because it's more of a guideline than a limit.
					}
					else // Need to create some new persistent memory for our temporary use.
					{
						// See the nearly identical section higher above for comments:
						if (mem_count == MAX_EXPR_MEM_ITEMS // No more slots left (should be nearly impossible).
							|| !(this_token.marker = mem[mem_count] = tmalloc(result_size)))
						{
							LineError(ERR_OUTOFMEM);
							goto abort;
						}
						++mem_count;
					}
					if (left_length)
						tmemcpy(this_token.marker, left_string, left_length);  // Not +1 because don't need the zero terminator.
					tmemcpy(this_token.marker + left_length, right_string, right_length + 1); // +1 to include its zero terminator.
					result_symbol = SYM_STRING;
					break;

				default:
					// Other operators do not support string operands, so the result is an empty string.
					this_token.marker = _T("");
					result_symbol = SYM_STRING;
				}
				this_token.symbol = result_symbol; // Must be done only after the switch() above.
			}

			else if (right_is_number == PURE_INTEGER && left_is_number == PURE_INTEGER && this_token.symbol != SYM_DIVIDE
				|| this_token.symbol <= SYM_BITSHIFTRIGHT && this_token.symbol >= SYM_BITOR) // Check upper bound first for short-circuit performance (because operators like +-*/ are much more frequently used).
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
					if (right_int64 == 0) // Divide by zero produces blank result (perhaps will produce exception if scripts ever support exception handlers).
					{
						this_token.marker = _T("");
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
						this_token.marker = _T("");
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
					if (right_double == 0.0) // Divide by zero produces blank result (perhaps will produce exception if script's ever support exception handlers).
					{
						this_token.marker = _T("");
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
						this_token.marker = _T("");
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

push_this_token:
		while (high_water_mark > stack_count)
		{
			// Release any objects which have been previously popped off the stack. This seems to be the
			// simplest way to do it as tokens are popped off the stack at multiple points, but only this
			// one point where parameters are pushed.  high_water_mark allows us to determine if there were
			// tokens on the stack before returning, regardless of how expression evaluation ended (abort,
			// abnormal_end, normal_end_skip_output_var...).  This method also ensures any objects passed as
			// parameters to a function (such as ObjGet()) are released *AFTER* the return value is made
			// persistent, which is important if the return value refers to the object's memory (but that
			// might currently never happen?).
			--high_water_mark;
			if (stack[high_water_mark]->symbol == SYM_OBJECT)
				stack[high_water_mark]->object->Release();
		}
		STACK_PUSH(&this_token);   // Push the result onto the stack for use as an operand by a future operator.
		high_water_mark = stack_count;
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
			result_to_return = _T(""); // Must not return NULL; any other value is OK (and will be ignored).
			goto normal_end_skip_output_var;
		case SYM_VAR:
			if (result_token.var->IsPureNumericOrObject())
			{
				result_token.var->ToToken(*aResultToken);
				result_to_return = _T("");
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
		// In case of float formats that are too long to be supported, use snprint() to restrict the length.
		 // %f probably defaults to %0.6f.  %f can handle doubles in MSVC++.
		aTarget += sntprintf(aTarget, MAX_NUMBER_SIZE, FORMAT_FLOAT, result_token.value_double) + 1; // +1 because that's what callers want; i.e. the position after the terminator.
		goto normal_end_skip_output_var; // output_var was already checked higher above, so no need to consider it again.
	case SYM_STRING:
	case SYM_VAR: // SYM_VAR is somewhat unusual at this late a stage.
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
			result_size = result_token.var->LengthIgnoreBinaryClip() + 1; // Ignore binary clipboard for anything other than ACT_ASSIGNEXPR (i.e. output_var!=NULL) because it's documented that except for certain features, binary clipboard variables are seen only up to the first binary zero (mostly to simplify the code).
		}
		else
		{
			result = result_token.marker;
			result_size = _tcslen(result) + 1;
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
		aTarget += result_size;
		goto normal_end_skip_output_var; // output_var was already checked higher above, so no need to consider it again.

	case SYM_OBJECT: // L31: Objects are always treated as empty strings; except with ACT_RETURN, which was handled above, and any usage which expects a boolean result.
		result_to_return = _T("");
		// result_token is still on the stack, so the object will be released below.
		goto normal_end_skip_output_var;

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
		if (!output_var->Assign(result_to_return))
			aResult = FAIL;

normal_end_skip_output_var:
	for (i = mem_count; i--;) // Free any temporary memory blocks that were used.  Using reverse order might reduce memory fragmentation a little (depending on implementation of malloc).
		free(mem[i]);

	// L31: Release any objects which have been previous pushed onto the stack and not yet released.
	while (high_water_mark)
	{	// See similar section under push_this_token for comments.
		--high_water_mark;
		if (stack[high_water_mark]->symbol == SYM_OBJECT)
			stack[high_water_mark]->object->Release();
	}

	return result_to_return;

	// Listing the following label last should slightly improve performance because it avoids an extra "goto"
	// in the postfix loop above the push_this_token label.  Also, keeping seldom-reached code at the end
	// may improve how well the code fits into the CPU cache.
double_deref_fail: // For the rare cases when the name of a dynamic function call is too long or improper.
	// A deref with a NULL marker terminates the list, and also indicates whether this is a dynamic function
	// call. "deref" has been set by the caller, and may or may not be the NULL marker deref.
	for (; deref->marker; ++deref);
	if (deref->is_function)
		goto abnormal_end;
	else
		goto push_this_token;
}



bool Func::Call(FuncCallData &aFuncCall, ResultType &aResult, ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount, bool aIsVariadic)
// aFuncCall: Caller passes a variable which should go out of scope after the function call's result
//   has been used; this automatically frees and restores a UDFs local vars (where applicable).
// aSpaceAvailable: -1 indicates this is a regular function call.  Otherwise this must be the amount of
//   space available after aParam for expanding the array of parameters for a variadic function call.
{
	aResult = OK; // Set default.

	if (mIsBuiltIn)
	{
		aResultToken.symbol = SYM_INTEGER; // Set default return type so that functions don't have to do it if they return INTs.
		aResultToken.marker = mName;       // Inform function of which built-in function called it (allows code sharing/reduction). Can't use circuit_token because it's value is still needed later below.

		if (aIsVariadic) // i.e. this is a variadic function call.
		{
			Object *param_obj;
			--aParamCount; // i.e. make aParamCount the count of normal params.
			if (param_obj = dynamic_cast<Object *>(TokenToObject(*aParam[aParamCount])))
			{
				void *mem_to_free;
				// Since built-in functions don't have variables we can directly assign to,
				// we need to expand the param object's contents into an array of tokens:
				if (!param_obj->ArrayToParams(mem_to_free, aParam, aParamCount, mMinParams))
					return false; // Abort expression.

				// CALL THE BUILT-IN FUNCTION:
				mBIF(aResult, aResultToken, aParam, aParamCount);

				if (mem_to_free)
					free(mem_to_free);
				if (g->ThrownToken)
					aResult = FAIL; // Abort thread.
				return (aResult != EARLY_EXIT && aResult != FAIL);
			}
			// Caller-supplied "params*" is not an Object, so treat it like an empty list; however,
			// mMinParams isn't validated at load-time for variadic calls, so we must do it here:
			if (aParamCount < mMinParams)
				return false; // Abort expression.
			// Otherwise just call the function normally.
		}

		// CALL THE BUILT-IN FUNCTION:
		mBIF(aResult, aResultToken, aParam, aParamCount);
		
		if (g->ThrownToken)
			aResult = FAIL; // Abort thread.
	}
	else // It's not a built-in function, or it's a built-in that was overridden with a custom function.
	{
		ExprTokenType indexed_token, named_token; // Separate for code simplicity.
		INT_PTR param_offset, param_key = -1;
		Object *param_obj = NULL;
		if (aIsVariadic) // i.e. this is a variadic function call.
		{
			--aParamCount; // i.e. make aParamCount the count of normal params.
			// For performance, only the Object class is supported:
			if (param_obj = dynamic_cast<Object *>(TokenToObject(*aParam[aParamCount])))
			{
				param_offset = -1;
				// Below retrieves the first item with integer key >= 1, or sets
				// param_offset to the offset of the first item with a non-int key.
				// Performance note: this might be slow if there are many items
				// with negative integer keys; but that should be vanishingly rare.
				while (param_obj->GetNextItem(indexed_token, param_offset, param_key)
						&& param_key < 1);
			}
		}

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
				if (this_param_token.symbol == SYM_VAR && !mParam[j].is_byref)
				{
					// Since this formal parameter is passed by value, if it's SYM_VAR, convert it to
					// a non-var to allow the variables to be backed up and reset further below without
					// corrupting any SYM_VARs that happen to be locals or params of this very same
					// function.
					// DllCall() relies on the fact that this transformation is only done for user
					// functions, not built-in ones such as DllCall().  This is because DllCall()
					// sometimes needs the variable of a parameter for use as an output parameter.
					this_param_token.var->ToToken(this_param_token);
				}
			}
			// BackupFunctionVars() will also clear each local variable and formal parameter so that
			// if that parameter or local var is assigned a value by any other means during our call
			// to it, new memory will be allocated to hold that value rather than overwriting the
			// underlying recursed/interrupted instance's memory, which it will need intact when it's resumed.
			if (!Var::BackupFunctionVars(*this, aFuncCall.mBackup, aFuncCall.mBackupCount)) // Out of memory.
			{
				aResult = g_script.ScriptError(ERR_OUTOFMEM, mName);
				return false;
			}
		} // if (func.mInstances > 0)
		//else backup is not needed because there are no other instances of this function on the call-stack.
		// So by definition, this function is not calling itself directly or indirectly, therefore there's no
		// need to do the conversion of SYM_VAR because those SYM_VARs can't be ones that were blanked out
		// due to a function exiting.  In other words, it seems impossible for a there to be no other
		// instances of this function on the call-stack and yet SYM_VAR to be one of this function's own
		// locals or formal params because it would have no legitimate origin.

		// Set after above succeeds to ensure local vars are freed and the backup is restored (at some point):
		aFuncCall.mFunc = this;
		
		// The following loop will have zero iterations unless at least one formal parameter lacks an actual,
		// which should be possible only if the parameter is optional (i.e. has a default value).
		for (j = aParamCount; j < mParamCount; ++j) // For each formal parameter that lacks an actual, provide a default value.
		{
			FuncParam &this_formal_param = mParam[j]; // For performance and convenience.
			if (this_formal_param.is_byref) // v1.0.46.13: Allow ByRef parameters to be optional by converting an omitted-actual into a non-alias formal/local.
				this_formal_param.var->ConvertToNonAliasIfNecessary(); // Convert from alias-to-normal, if necessary.
			// Check if this parameter has been supplied a value via a param array/object:
			if (param_obj)
			{
				// Numbered parameter?
				if (param_key == j - aParamCount + 1)
				{
					if (!this_formal_param.var->Assign(indexed_token))
					{
						aResult = FAIL; // Abort thread.
						return false;
					}
					// Get the next item, which might be at [i+1] (in which case the next iteration will
					// get indexed_token) or a later index (in which case the next iteration will get a
					// regular default value and a later iteration may get indexed_token). If there aren't
					// any more items, param_key will be left as-is (so this IF won't be re-entered).
					// If there aren't any more parameters needing values, param_offset will remain the
					// offset of the first unused item; this is relied on below to copy the remaining
					// items into the function's "param*" array if it has one (i.e. if mIsVariadic).
					param_obj->GetNextItem(indexed_token, param_offset, param_key);
					// This parameter has been supplied a value, so don't assign it a default value.
					continue;
				}
				// Named parameter?
				if (param_obj->GetItem(named_token, this_formal_param.var->mName))
				{
					if (!this_formal_param.var->Assign(named_token))
					{
						aResult = FAIL; // Abort thread.
						return false;
					}
					continue;
				}
			}
			switch(this_formal_param.default_type)
			{
			case PARAM_DEFAULT_STR:   this_formal_param.var->Assign(this_formal_param.default_str);    break;
			case PARAM_DEFAULT_INT:   this_formal_param.var->Assign(this_formal_param.default_int64);  break;
			case PARAM_DEFAULT_FLOAT: this_formal_param.var->Assign(this_formal_param.default_double); break;
			default: //case PARAM_DEFAULT_NONE:
				// Since above didn't continue, no value has been supplied for this REQUIRED parameter.
				return false; // Abort expression.
			}
		}

		for (j = 0; j < count_of_actuals_that_have_formals; ++j) // For each actual parameter that has a formal, assign the actual to the formal.
		{
			ExprTokenType &token = *aParam[j];
			
			if (!IS_OPERAND(token.symbol)) // Haven't found a way to produce this situation yet, but safe to assume it's possible.
				return false; // Abort expression.
			
			if (mParam[j].is_byref)
			{
				// Note that the previous loop might not have checked things like the following because that
				// loop never ran unless a backup was needed:
				if (token.symbol != SYM_VAR)
				{
					// L60: Seems more useful and in the spirit of AutoHotkey to allow ByRef parameters
					// to act like regular parameters when no var was specified.  If we force script
					// authors to pass a variable, they may pass a temporary variable which is then
					// discarded, adding a little overhead and impacting the readability of the script.
					mParam[j].var->ConvertToNonAliasIfNecessary();
				}
				else
				{
					mParam[j].var->UpdateAlias(token.var); // Make the formal parameter point directly to the actual parameter's contents.
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
			if (!mParam[j].var->Assign(token))
			{
				aResult = FAIL; // Abort thread.
				return false;
			}
		} // for each formal parameter.
		
		if (mIsVariadic) // i.e. this function is capable of accepting excess params via an object/array.
		{
			Object *obj;
			if (param_obj) // i.e. caller supplied an array of params.
			{
				// Clone the caller's param object, excluding the numbered keys we've used:
				if (obj = param_obj->Clone(param_offset))
				{
					if (mParamCount > aParamCount)
						// Adjust numeric keys based on how many items we would've used if the "array" was contiguous:
						obj->ReduceKeys(mParamCount - aParamCount); // Should be harmless if there are no numeric keys left.
					//else param_offset should be 0; we didn't use any items.
				}
			}
			else
				obj = (Object *)Object::Create();
			
			if (obj)
			{
				if (j < aParamCount)
					// Insert the excess parameters from the actual parameter list.
					obj->InsertAt(0, 1, aParam + j, aParamCount - j);
				// Assign to the "param*" var:
				mParam[mParamCount].var->AssignSkipAddRef(obj);
			}
		}

		aResult = Call(&aResultToken); // Call the UDF.
	}
	return (aResult != EARLY_EXIT && aResult != FAIL);
}

// This is used for maintainability: to ensure it's never forgotten and to reduce code repetition.
FuncCallData::~FuncCallData()
{
	if (mFunc) // mFunc != NULL implies it is a UDF and Var::BackupFunctionVars() has succeeded.
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
		Var::FreeAndRestoreFunctionVars(*mFunc, mBackup, mBackupCount);
	}
}



ResultType Line::ExpandArgs(ExprTokenType *aResultToken, VarSizeType aSpaceNeeded, Var *aArgVar[])
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
	size_t space_needed;
	if (aSpaceNeeded == VARSIZE_ERROR)
	{
		space_needed = GetExpandedArgSize(arg_var);
		if (space_needed == VARSIZE_ERROR)
			return FAIL;  // It will have already displayed the error.
	}
	else // Caller already determined it.
	{
		space_needed = aSpaceNeeded;
		for (i = 0; i < mArgc; ++i) // Copying only the actual/used elements is probably faster than using memcpy to copy both entire arrays.
			arg_var[i] = aArgVar[i]; // Init to values determined by caller, which helps performance if any of the args are dynamic variables.
	}

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
				arg_deref[i] = ExpandExpression(i, result, mActionType == ACT_RETURN ? aResultToken : NULL  // L31: aResultToken is used to return a non-string value. Pass NULL if mMctionType != ACT_RETURN for maintainability; non-NULL aResultToken should mean we want a token returned - this can be used in future for numeric params or array support in commands.
					, our_buf_marker, our_deref_buf, our_deref_buf_size, arg_deref, extra_size);
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
				arg_deref[i] = _T("");
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
				// Otherwise there's more than one variable in the arg, so it must be expanded in the normal,
				// lower-performance way.
				arg_deref[i] = our_buf_marker; // Point it to its location in the buffer.
				if (   !(our_buf_marker = ExpandArg(our_buf_marker, i))   ) // Expand the arg into that location.
				{
					result_to_return = FAIL; // ExpandArg() will have already displayed the error.
					goto end;
				}
				continue;
			}

			// Since above didn't "continue", the_only_var_of_this_arg==true, so this arg resolves to
			// only a single, naked var.
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
					(   ACT_IS_ASSIGN(mActionType) && i == 1  // By contrast, for the below i==anything (all args):
					||  mActionType == ACT_IF
					//|| mActionType == ACT_WHILE // Not necessary to check this one because loadtime leaves ACT_WHILE as an expression in all common cases. Also, there's no easy way to get ACT_WHILE into the range above due to the overlap of other ranges in enum_act.
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

		// IT'S NOT SAFE to do the following until the above loop FULLY completes because any calls made above to
		// ExpandExpression() might call functions, which in turn might result in a recursive call to ExpandArgs(),
		// which in turn might change the values in the static arrays sArgDeref and sArgVar.
		// Also, only when the loop ends normally is the following needed, since otherwise it's a failure condition.
		// Now that any recursive calls to ExpandArgs() above us on the stack have collapsed back to us, it's
		// safe to set the args of this command for use by our caller, to whom we're about to return.
		for (i = 0; i < mArgc; ++i) // Copying actual/used elements is probably faster than using memcpy to copy both entire arrays.
		{
			sArgDeref[i] = arg_deref[i] ? arg_deref[i] : arg_var[i]->Contents(); // See "Update #3" comment above.
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

		// Accumulate the total of how much space we will need.
		if (this_arg.type == ARG_TYPE_OUTPUT_VAR)  // These should never be included in the space calculation.
		{
			if (   !(aArgVar[i] = ResolveVarOfArg(i))   ) // v1.0.45: Resolve output variables too, which eliminates a ton of calls to ResolveVarOfArg() in various other functions.  This helps code size more than performance.
				return VARSIZE_ERROR;  // The above will have already displayed the error.
			continue;
		}
		// Otherwise, set default aArgVar[] (above took care of setting aArgVar[] for itself).
		aArgVar[i] = NULL;

		if (this_arg.is_expression)
		{
			// Now that literal strings/numbers are handled by ExpressionToPostfix(), the length used below
			// is more room than is strictly necessary. But given how little space is typically wasted (and
			// that only while the expression is being evaluated), it doesn't seem worth worrying about it.
			// See other comments at macro definition.
			space_needed += EXPR_BUF_SIZE(this_arg.length);
			continue;
		}

		// Otherwise:
		// Always do this check before attempting to traverse the list of dereferences, since
		// such an attempt would be invalid in this case:
		the_only_var_of_this_arg = NULL;
		if (this_arg.type == ARG_TYPE_INPUT_VAR) // Previous stage has ensured that arg can't be an expression if it's an input var.
			if (   !(the_only_var_of_this_arg = ResolveVarOfArg(i, false))   )
				return VARSIZE_ERROR;  // The above will have already displayed the error.

		if (!the_only_var_of_this_arg) // It's not an input var.
		{
			if (NO_DEREF)
				// Don't increase space_needed, even by 1 for the zero terminator, because
				// the terminator isn't needed if the arg won't exist in the buffer at all.
				continue;
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

		// Otherwise: This arg has more than one deref, or a single deref with some literal text around it.
		space_needed += this_arg.length + 1; // +1 for this arg's zero terminator in the buffer.
		if (this_arg.deref) // There's at least one deref.
		{
			for (DerefType *deref = this_arg.deref; deref->marker; ++deref)
			{
				// Replace the length of the deref's literal text with the length of its variable's contents.
				// At this point, this_arg.is_expression is known to be false. Since non-expressions can't
				// contain function-calls, there's no need to check deref->is_function.
				space_needed -= deref->length;
				space_needed += deref->var->Get(); // If an environment var, Get() will yield its length.
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
// There are some other functions like ArgLength() that have procedures similar to this one, so
// maintain them together.
{
	if (mActionType == ACT_SORT) // See PerformSort() for why it's always dereferenced.
		return CONDITION_TRUE;
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



LPTSTR Line::ExpandArg(LPTSTR aBuf, int aArgIndex, Var *aArgVar) // 10/2/2006: Doesn't seem worth making it inline due to more complexity than expected.  It would also increase code size without being likely to help performance much.
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
		LineError(_T("DEBUG: ExpandArg() was called to expand an arg that contains only an output variable."));
		return NULL;
	}
#endif

	if (aArgVar)
		// +1 so that we return the position after the terminator, as required.
		return aBuf += aArgVar->Get(aBuf) + 1;

	LPTSTR this_marker, pText = this_arg.text;  // Start at the beginning of this arg's text.
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
