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
LPTSTR Line::ExpandExpression(int aArgIndex, ResultType &aResult, ResultToken *aResultToken
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

	// The following must be defined early so that to_free_count is initialized and the array is guaranteed to be
	// "in scope" in case of early "goto" (goto substantially boosts performance and reduces code size here).
	ExprTokenType **to_free = (ExprTokenType **)_alloca(mArg[aArgIndex].max_alloc * sizeof(ExprTokenType *));
	int to_free_count = 0; // The actual number of items in use in the above array.
	LPTSTR result_to_return = _T(""); // By contrast, NULL is used to tell the caller to abort the current thread.
	LPCTSTR error_msg = ERR_EXPR_EVAL, error_info = _T("");
	ExprTokenType *error_value;
	Var *output_var = (mActionType == ACT_ASSIGNEXPR) ? VAR(mArg[0]) : NULL; // Resolve early because it's similar in usage/scope to the above.

	ExprTokenType **stack = (ExprTokenType **)_alloca(mArg[aArgIndex].max_stack * sizeof(ExprTokenType *));
	int stack_count = 0;
	ExprTokenType *&postfix = mArg[aArgIndex].postfix;

	///////////////////////////////
	// EVALUATE POSTFIX EXPRESSION
	///////////////////////////////
	int i, delta;
	SymbolType right_is_number, left_is_number, right_is_pure_number, left_is_pure_number, result_symbol;
	double right_double, left_double;
	__int64 right_int64, left_int64;
	LPTSTR right_string, left_string;
	size_t right_length, left_length;
	TCHAR left_buf[max(MAX_NUMBER_SIZE, _f_retval_buf_size)];
	TCHAR right_buf[MAX_NUMBER_SIZE];
	LPTSTR result; // "result" is used for return values and also the final result.
	VarSizeType result_length;
	size_t result_size, alloca_usage = 0; // v1.0.45: Track amount of alloca mem to avoid stress on stack from extreme expressions (mostly theoretical).
	BOOL done, done_and_have_an_output_var, left_branch_is_true
		, left_was_negative, is_pre_op; // BOOL vs. bool benchmarks slightly faster, and is slightly smaller in code size (or maybe it's cp1's int vs. char that shrunk it).
	ExprTokenType *this_postfix, *p_postfix;
	Var *sym_assign_var, *temp_var;

	// v1.0.44.06: EXPR_SMALL_MEM_LIMIT is the means by which _alloca() is used to boost performance a
	// little by avoiding the overhead of malloc+free for small strings.  The limit should be something
	// small enough that the chance that even 10 of them would cause stack overflow is vanishingly small.
	#define EXPR_SMALL_MEM_LIMIT 4097 // The maximum size allowed for an item to qualify for alloca.
	#define EXPR_ALLOCA_LIMIT 40000  // The maximum amount of alloca memory for all items.  v1.0.45: An extra precaution against stack stress in extreme/theoretical cases.
	#define EXPR_IS_DONE (!stack_count && this_postfix[1].symbol == SYM_INVALID) // True if we've used up the last of the operators & operands.  Non-zero stack_count combined with SYM_INVALID would indicate an error (an exception will be thrown later, so don't take any shortcuts).

	// For each item in the postfix array: if it's an operand, push it onto stack; if it's an operator or
	// function call, evaluate it and push its result onto the stack.  SYM_INVALID is the special symbol
	// that marks the end of the postfix array.
	for (this_postfix = postfix; this_postfix->symbol != SYM_INVALID; ++this_postfix) // Using pointer vs. index (e.g. postfix[i]) reduces OBJ code size by ~122 and seems to perform at least as well.
	{
		// Set default early to simplify the code.  All struct members may be needed.  Also, this_token is used
		// almost everywhere further below in preference to this_postfix because:
		// 1) The various SYM_ASSIGN_* operators (e.g. SYM_ASSIGN_CONCAT) are changed to different operators
		//    to simplify the code.  So must use the changed/new value in this_token, not the original value in
		//    this_postfix.
		// 2) Using a particular variable very frequently might help compiler to optimize that variable to
		//    generate faster code.
		// 3) It might help performance due to locality of reference/CPU caching (this_token is on the stack).
		ExprTokenType &this_token = *(ExprTokenType *)_alloca(sizeof(ExprTokenType)); // Saves a lot of stack space, and seems to perform just as well as something like the following (at the cost of ~82 byte increase in OBJ code size): ExprTokenType &this_token = new_token[new_token_count++]
		this_token.CopyExprFrom(*this_postfix); // See comment section above.

		if (IS_OPERAND(this_token.symbol)) // If it's an operand, just push it onto stack for use by an operator in a future iteration.
		{
			if (this_token.symbol == SYM_DYNAMIC) // Dynamic variable reference.
			{
				if (!stack_count) // Prevent stack underflow.
					goto abort_with_exception;
				ExprTokenType &right = *STACK_POP;
				if (auto right_obj = TokenToObject(right))
				{
					if (right_obj->Base() == Object::sVarRefPrototype)
						temp_var = static_cast<VarRef *>(right_obj);
					else
					{
						error_info = _T("String or VarRef");
						error_value = &right;
						goto type_mismatch;
					}
				}
				else
				{
					right_string = TokenToString(right, right_buf, &right_length);
					// Do some basic validation to ensure a helpful error message is displayed on failure.
					if (right_length == 0)
					{
						error_msg = ERR_DYNAMIC_BLANK;
						error_info = mArg[aArgIndex].text;
						goto abort_with_exception;
					}
					// v2.0: Use of FindVar() vs. FindOrAddVar() makes this check unnecessary.
					//if (right_length > MAX_VAR_NAME_LENGTH)
					//{
					//	error_msg = ERR_DYNAMIC_TOO_LONG;
					//	error_info = right_string;
					//	goto abort_with_exception;
					//}
					// v2.0: Dynamic creation of variables is not permitted, so FindOrAddVar() is not used.
					if (!(temp_var = g_script.FindVar(right_string, right_length
						, VARREF_IS_WRITE(this_token.var_usage) ? FINDVAR_FOR_WRITE : FINDVAR_FOR_READ)))
					{
						if (this_token.var_usage == VARREF_ISSET) // this_token is to be passed to IsSet().
						{
							ASSERT(this_postfix[1].symbol == SYM_FUNC);
							++this_postfix; // Skip the actual IsSet call since we're pushing its result directly.
							this_token.SetValue(FALSE);
							goto push_this_token;
						}
						if (this_token.var_usage == VARREF_READ_MAYBE)
						{
							this_token.symbol = SYM_MISSING;
							goto push_this_token;
						}
						if (g->CurrentFunc && g_script.FindGlobalVar(right_string, right_length))
							error_msg = ERR_DYNAMIC_BAD_GLOBAL;
						else
							error_msg = ERR_DYNAMIC_NOT_FOUND;
						error_info = right_string;
						goto abort_with_exception;
					}
				}
				// var_usage was already validated for non-dynamic variables.  SYM_REF relies on this check.
				// It also flags invalid assignments, but those can be handled by Var::Assign() anyway.
				if (!ValidateVarUsage(temp_var, this_token.var_usage))
					goto abort;
				this_token.var = temp_var;
				this_token.symbol = SYM_VAR;
			}
			if (this_token.symbol == SYM_VAR && (!VARREF_IS_WRITE(this_token.var_usage) || this_token.var_usage == VARREF_LVALUE_MAYBE))
			{
				if (this_token.var->Type() == VAR_VIRTUAL && VARREF_IS_READ(this_token.var_usage))
				{
					// FUTURE: This should be merged with the SYM_FUNC handling at some point to improve
					// maintainability, reduce code size, and take advantage of SYM_FUNC's optimizations.
					ResultToken result_token;
					result_token.InitResult(left_buf);
					result_token.symbol = SYM_INTEGER; // For _f_return_i() and consistency with BIFs.

					// Call this virtual variable's getter.
					this_token.var->Get(result_token);

					if (result_token.Exited())
					{
						aResult = result_token.Result(); // See similar section under SYM_FUNC for comments.
						result_to_return = NULL;
						goto normal_end_skip_output_var;
					}

					if (result_token.symbol != SYM_STRING || result_token.marker_length == 0)
					{
						// Currently SYM_OBJECT is not added to to_free[] as there aren't any built-in
						// vars that create an object or call AddRef().  If that's changed, must update
						// BIV_TrayMenu and Debugger::GetPropertyValue.
						this_token.CopyValueFrom(result_token);
						goto push_this_token;
					}

					result_length = result_token.marker_length;
					if (result_length == -1)
						result_length = _tcslen(result_token.marker);

					if (result_token.marker != left_buf)
					{
						if (result_token.mem_to_free) // Persistent memory was already allocated for the result.
							to_free[to_free_count++] = &this_token; // A slot was reserved for this SYM_DYNAMIC.
							// Also push the value, below.
						//else: Currently marker is assumed to point to persistent memory, such as a literal
						// string, which should be safe to use at least until expression evaluation completes.
						this_token.SetValue(result_token.marker, result_length);
						goto push_this_token;
					}

					result_size = 1 + result_length;
					if (result_size <= (int)(aDerefBufSize - (target - aDerefBuf))) // There is room at the end of our deref buf, so use it.
					{
						result = target; // Point result to its new, more persistent location.
						target += result_size; // Point it to the location where the next string would be written.
					}
					else // if (result_size < EXPR_SMALL_MEM_LIMIT && alloca_usage < EXPR_ALLOCA_LIMIT) // See comments at EXPR_SMALL_MEM_LIMIT.
					{	// Above: Anything longer than MAX_NUMBER_SIZE can't be in left_buf and therefore
						// was already handled above.  alloca_usage is a safeguard but not an absolute limit.
						result = (LPTSTR)talloca(result_size);
						alloca_usage += result_size; // This might put alloca_usage over the limit by as much as EXPR_SMALL_MEM_LIMIT, but that is fine because it's more of a guideline than a limit.
					}
					tmemcpy(result, result_token.marker, result_length + 1);
					this_token.SetValue(result, result_length);
					goto push_this_token;
				} // end if (reading a var of type VAR_VIRTUAL)
				if (this_token.var->IsUninitialized())
				{
					if (this_token.var->Type() == VAR_CONSTANT)
					{
						auto result = this_token.var->InitializeConstant();
						if (result != OK)
						{
							aResult = result;
							result_to_return = NULL;
							goto normal_end_skip_output_var;
						}
					}
					else if (this_token.var_usage == VARREF_READ)
					{
						// The expression is always aborted in this case, even if the user chooses to continue the thread.
						// If this is changed, check all other callers of unset_var and VarUnsetError() for consistency.
						error_value = &this_token;
						goto unset_var;
					}
					else if (this_token.var_usage == VARREF_READ_MAYBE)
					{
						this_token.symbol = SYM_MISSING;
						goto push_this_token;
					}
					else if (this_token.var_usage == VARREF_LVALUE_MAYBE)
					{
						// Skip the short-circuit operator and push the variable onto the stack for assignment.
						++this_postfix;
						ASSERT(this_postfix->symbol == SYM_OR_MAYBE);
					}
				}
			}
			goto push_this_token;
		} // if (IS_OPERAND(this_token.symbol))

		if (this_token.symbol == SYM_FUNC)
		{
			IObject *func = this_token.callsite->func; // For convenience and because any modifications should not be persistent.
			auto member = this_token.callsite->member;
			auto flags = this_token.callsite->flags;
			auto param_count = this_token.callsite->param_count;
			if (param_count > stack_count) // Prevent stack underflow (probably impossible if actual_param_count is accurate).
				goto abort_with_exception;
			// Adjust the stack early to simplify.  Above already confirmed that the following won't underflow.
			// Pop the actual number of params involved in this function-call off the stack.
			int prev_stack_count = stack_count;
			stack_count -= param_count;
			ExprTokenType **params = stack + stack_count;
			ExprTokenType *func_token;

			if (flags & EIF_STACK_MEMBER)
			{
				if (!stack_count)
					goto abort_with_exception;
				stack_count--;
				flags &= ~EIF_STACK_MEMBER;
				if (stack[stack_count]->symbol != SYM_MISSING)
				{
					member = TokenToString(*stack[stack_count], right_buf);
					if (!*member && TokenToObject(*stack[stack_count]))
					{
						error_info = _T("String");
						error_value = stack[stack_count];
						goto type_mismatch;
					}
				}
			}

			if (func)
			{
				func_token = (ExprTokenType *)_alloca(sizeof(ExprTokenType));
				func_token->SetValue(func);
			}
			else
			{
				// This is a dynamic function call.
				if (!stack_count)
					goto abort_with_exception;
				stack_count--;
				func_token = stack[stack_count];
				func = TokenToObject(*func_token);
				if (!func)
				{
					if (func_token->symbol == SYM_SUPER)
					{
						// Invoke the super-class but pass the current function's "this".
						ASSERT(g->CurrentFunc && g->CurrentFunc->mClass);
						func = g->CurrentFunc->mClass->Base();
						ASSERT(func);
						func_token->SetVar(g->CurrentFunc->mParam[0].var);
						flags |= IF_SUPER;
					}
					else
					{
						// Invoke a substitute object but pass func_token as "this".
						func = Object::ValueBase(*func_token);
						flags |= IF_SUBSTITUTE_THIS;
					}
				}
				if (!func) // Possible for SYM_MISSING in cases that currently can't be detected at load-time, like (a ? unset : b).
					goto abort_with_exception;
			}

			// The following two steps are done for built-in functions inside Func::Call:
			//result_token.symbol = SYM_INTEGER; // Set default return type so that functions don't have to do it if they return INTs.
			//result_token.func = func;          // Inform function of which built-in function called it (allows code sharing/reduction).
			
			// This is done by ResultToken below:
			//result_token.buf = left_buf;       // mBIF() can use this to store a string result, and for other purposes.
			//result_token.mem_to_free = NULL;   // Init to detect whether the called function allocates it.
			
			ResultToken result_token;
			result_token.InitResult(left_buf); // But we'll take charge of its contents INSTEAD of calling Free().

			// Invoke the function or object.
			bool keep_alive = func_token->symbol == SYM_VAR;
			if (keep_alive) // Might help performance to avoid these virtual calls in common cases.
				func->AddRef(); // Ensure the object isn't deleted during the call, by an assignment.
			ResultType invoke_result;
			if (flags & EIF_VARIADIC)
				invoke_result = VariadicCall(func, result_token, flags, member, *func_token, params, param_count);
			else
				invoke_result = func->Invoke(result_token, flags, member, *func_token, params, param_count);
			if (keep_alive)
				func->Release();

			switch (invoke_result)
			{
			case FAIL:
			case EARLY_EXIT:
				aResult = invoke_result;
				result_to_return = NULL; // Use NULL to inform our caller that this thread is finished (whether through normal means such as Exit or a critical error).
				// Above: The callers of this function know that the value of aResult (which already contains the
				// reason for early exit) should be considered valid/meaningful only if result_to_return is NULL.
				goto normal_end_skip_output_var; // output_var is left unchanged in these cases.
			case INVOKE_NOT_HANDLED:
				if (!(flags & EIF_UNSET_PROP))
				{
					result_token.UnknownMemberError(*func_token, flags, member);
					aResult = result_token.Result(); // FAIL to abort, OK if user or OnError requested continuation.
					goto abort_if_result;
				}
				// For something like (a.b?) or (a.b ?? c), INVOKE_NOT_HANDLED is treated as unset.
				result_token.symbol = SYM_MISSING;
			}

#ifdef CONFIG_DEBUGGER
			// See PostExecFunctionCall() itself for comments.
			if (g_Debugger.IsConnected())
				g_Debugger.PostExecFunctionCall(this);
#endif
			g_script.mCurrLine = this; // For error-reporting.
			
			if ((flags & EIF_LEAVE_PARAMS)
				&& (!(flags & EIF_UNSET_RETURN) || result_token.symbol == SYM_MISSING))
				// Leave params on the stack for the next part of a compound assignment.
				// The combination of (EIF_LEAVE_PARAMS | EIF_UNSET_RETURN) implies this is
				// something like the `x.y` in `x.y ??= z`, which needs to take the params
				// off the stack if it's going to short-circuit (i.e. result is unset).
				stack_count = prev_stack_count;
			
			if (flags & IT_SET)
			{
				result_token.Free();
				auto &value = *params[param_count - 1];
				// value came from a previous part of this expression, so it's already in to_free[]
				// if appropriate, and we can just push it back onto the stack.
				this_token.CopyValueFrom(value);
				goto push_this_token;
			}

			if (result_token.symbol != SYM_STRING)
			{
				if (result_token.symbol == SYM_MISSING && !(flags & EIF_UNSET_RETURN))
				{
					result_token.Error(_T("No value was returned.")
						, this_token.error_reporting_marker
						, (flags & IT_BITMASK) == IT_GET && !member ? ErrorPrototype::UnsetItem : ErrorPrototype::Unset);
					aResult = result_token.Result(); // FAIL to abort, OK if user or OnError requested continuation.
					goto abort_if_result;
				}
				// No need for make_result_persistent or early Assign().  Any numeric or object result can
				// be considered final because it's already stored in permanent memory (the token itself).
				// Additionally, this_token.mem_to_free is assumed to be NULL since the result is not
				// a string; i.e. the function would've had no need to return memory to us.
				this_token.value_int64 = result_token.value_int64;
				this_token.symbol = result_token.symbol;
				if (this_token.symbol == SYM_OBJECT)
					to_free[to_free_count++] = &this_token; // A slot was reserved for this SYM_FUNC.
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
					&& stack[stack_count-1]->var->Type() == VAR_NORMAL) // Don't do VAR_VIRTUAL here; it mustn't become a SYM_VAR result, so the result would have to be made persistent anyway.
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
				this_token.SetVar(internal_output_var);
				++this_postfix; // We've fully handled the assignment.
				goto push_this_token;
			}
			// Otherwise, there's no output_var or the expression isn't finished yet, so do normal processing.
				
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
				// Mark it to be freed at the time we return.
				to_free[to_free_count++] = &this_token; // A slot was reserved for this SYM_FUNC.
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

			bool make_result_persistent;
			{
				// Since above didn't goto, the result may need to be copied to a more persistent location.

				// For BIFs, "result" can be any of the following (some of which were handled above):
				//  - mem_to_free:  Allocated by TokenSetResult().
				//  - left_buf:  Copied there by TokenSetResult() or via _f_retval_buf.
				//  - Static memory, such as a literal string.
				//  - Others that might be volatile.

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
					this_token.marker = tmemcpy(target, result, result_length); // Benches slightly faster than strcpy().
					target += result_size; // Point it to the location where the next string would be written.
				}
				else if (result_size < EXPR_SMALL_MEM_LIMIT && alloca_usage < EXPR_ALLOCA_LIMIT) // See comments at EXPR_SMALL_MEM_LIMIT.
				{
					this_token.marker = tmemcpy(talloca(result_size), result, result_length); // Benches slightly faster than strcpy().
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
					if (  !(this_token.marker = tmalloc(result_size))  )
						goto outofmem;
					tmemcpy(this_token.marker, result, result_length); // Benches slightly faster than strcpy().
					to_free[to_free_count++] = &this_token; // A slot was reserved for this SYM_FUNC.
				}
				// Must be null-terminated because some built-in functions (such as MsgBox) still require it.
				// Explicit null-termination here (vs. including it in tmemcpy above) allows SubStr and others
				// to return a non-terminated substring.
				this_token.marker[result_length] = '\0';
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
		if (right.symbol == SYM_MISSING)
		{
			if (this_token.symbol == SYM_OR_MAYBE // SYM_MISSING is to ?? what False is to ||.
				|| this_token.symbol == SYM_COMMA) // The operand of a multi-statement comma is always ignored, so is not required to be a value.
				continue; // Continue on to evaluate the right branch.
			if (this_token.symbol == SYM_MAYBE)
			{
				++stack_count; // Put unset back on the stack.
				this_postfix = this_token.circuit_token;
				continue;
			}
			if (this_token.symbol != SYM_ASSIGN) // Anything other than := is not permitted.
				goto abort_with_exception;
		}

		switch (this_token.symbol)
		{
		case SYM_ASSIGN:        // These don't need "right_is_number" to be resolved. v1.0.48.01: Also avoid
		case SYM_CONCAT:        // resolving right_is_number for CONCAT because TokenIsPureNumeric() will take
		case SYM_ASSIGN_CONCAT: // a long time if the string is very long and consists entirely of digits/whitespace.
		case SYM_IS:
			right_is_pure_number = right_is_number = PURE_NOT_NUMERIC; // Init for convenience/maintainability.
		case SYM_AND:			// v2: These don't need it either since even numeric strings are considered "true".
		case SYM_OR:			// right_is_number isn't used at all in these cases since they are handled early.
		case SYM_OR_MAYBE:		//
		case SYM_IFF_THEN:		//
		case SYM_LOWNOT:		//
		case SYM_HIGHNOT:		//
		case SYM_REF:
			break;
			
		case SYM_COMMA: // This can only be a statement-separator comma, not a function comma, since function commas weren't put into the postfix array.
			// Do nothing other than discarding the operand that was just popped off the stack, which is the
			// result of the comma's left-hand sub-statement.  At this point the right-hand sub-statement
			// has not yet been evaluated.  Like C++ and other languages, but unlike AutoHotkey v1, the
			// rightmost operand is preserved, not the leftmost.
			continue;

		case SYM_MAYBE:
			++stack_count; // right was already confirmed to not be SYM_MISSING, so just put it back on the stack.
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
		case SYM_OR_MAYBE: // This case skips the "if" above, since the false (SYM_MISSING) branch was already handled.
				// The ternary's condition is false or this AND/OR causes a short-circuit.
				// Discard the entire right branch of this AND/OR or "then" branch of this IFF:
				this_postfix = this_token.circuit_token; // The address in any circuit_token always points into the arg's postfix array (never any temporary array or token created here) due to the nature/definition of circuit_token.

				if (this_token.symbol != SYM_IFF_THEN)
				{
					// This will be the final result of this AND/OR because it's right branch was
					// discarded above without having been evaluated nor any of its functions called:
					++stack_count; // It's already at stack[stack_count], so this "puts it back" on the stack.
					// Any SYM_OBJECT on our stack was already put into to_free[], so if this is SYM_OBJECT,
					// there's no need to do anything; we actually MUST NOT AddRef() unless we also put it
					// into to_free[].
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
			else // String.  Seems best to consider the application of unary minus to a string to be a failure.
			{
				error_info = _T("Number");
				error_value = &right;
				goto type_mismatch;
			}
			// Since above didn't "break":
			this_token.symbol = right_is_number;
			break;

		case SYM_POSITIVE: // Added in v2 for symmetry with SYM_NEGATIVE; i.e. if -x produces NaN, so should +x.
			if (right_is_number)
				TokenToDoubleOrInt64(right, this_token);
			else
			{
				error_info = _T("Number");
				error_value = &right;
				goto type_mismatch; // For consistency with unary minus (see above).
			}
			break;

		case SYM_REF:
			if (right.symbol != SYM_VAR) // Syntax error?
				goto abort_with_exception;
			if (this_token.var_usage != VARREF_READ)
			{
				if (this_token.var_usage != VARREF_REF)
				{
					// VARREF_ISSET or VARREF_OUTPUT_VAR -> SYM_VAR.
					this_token.SetVarRef(right.var);
					goto push_this_token;
				}
				Var *target_var = right.var->ResolveAlias();
				if (!target_var->IsNonStaticLocal()
					|| !this_token.object
					|| !((UserFunc *)this_token.object)->mInstances)
				{
					// target_var definitely isn't a local var of the function being called,
					// so it's safe to pass as SYM_VAR.  Pass right.var and not target_var,
					// otherwise GetRef() won't be able to identify the existing VarRef and
					// may create a new VarRef and a circular reference.
					this_token.SetVarRef(right.var);
					goto push_this_token;
				}
			}
			ASSERT(right.var->Type() == VAR_NORMAL);
			this_token.SetValue(right.var->GetRef());
			if (!this_token.object)
				goto outofmem;
			to_free[to_free_count++] = &this_token;
			break;

		case SYM_POST_INCREMENT: // These were added in v1.0.46.  It doesn't seem worth translating them into
		case SYM_POST_DECREMENT: // += and -= at load-time or during the tokenizing phase higher above because 
		case SYM_PRE_INCREMENT:  // it might introduce precedence problems, plus the post-inc/dec's nature is
		case SYM_PRE_DECREMENT:  // unique among all the operators in that it pushes an operand before the evaluation.
			if (right.symbol != SYM_VAR) // Syntax error.
				goto abort_with_exception;
			is_pre_op = SYM_INCREMENT_OR_DECREMENT_IS_PRE(this_token.symbol); // Store this early because its symbol will soon be overwritten.
			if (right_is_number == PURE_NOT_NUMERIC) // Not numeric: invalid operation.
			{
				error_info = _T("Number");
				error_value = &right;
				goto type_mismatch;
			}

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
					this_token.SetVar(right.var);
				}
				else // VAR_VIRTUAL, which is allowed in only when it's the lvalue of an assignment or inc/dec.
				{
					// VAR_VIRTUAL isn't allowed as SYM_VAR beyond this point (to simplify the code and
					// improve maintainability).  So use the new contents of the variable as the result,
					// rather than the variable itself.
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

		case SYM_BITNOT:  // The tilde (~) operator.
			if (right_is_number != PURE_INTEGER) // String.  Seems best to consider the application of '*' or '~' to a non-numeric string to be a failure.
			{
				error_info = _T("Number");
				error_value = &right;
				goto type_mismatch;
			}
			// Since above didn't "goto": right_is_number is PURE_INTEGER.
			right_int64 = TokenToInt64(right); // The slight performance reduction of calling TokenToInt64() is done for brevity.
			
			// Note that it is not legal to perform ~, &, |, or ^ on doubles.  
			// Treat it as a 64-bit signed value, since no other aspects of the program
			// will recognize an unsigned 64 bit number.
			this_token.value_int64 = ~right_int64;
			this_token.symbol = SYM_INTEGER; // Must be done only after its old value was used above. v1.0.36.07: Fixed to be SYM_INTEGER vs. right_is_number for SYM_BITNOT.
			break;

		default: // NON-UNARY OPERATOR.
			// GET THE SECOND (LEFT-SIDE) OPERAND FOR THIS OPERATOR:
			if (!stack_count) // Prevent stack underflow.
				goto abort_with_exception;
			ExprTokenType &left = *STACK_POP; // i.e. the right operand always comes off the stack before the left.
			if (left.symbol == SYM_MISSING)
				goto abort_with_exception;
			
			if (IS_ASSIGNMENT_EXCEPT_POST_AND_PRE(this_token.symbol)) // v1.0.46: Added support for various assignment operators.
			{
				if (left.symbol != SYM_VAR) // Syntax error.
					goto abort_with_exception;

				switch(this_token.symbol)
				{
				case SYM_ASSIGN: // Listed first for performance (it's probably the most common because things like ++ and += aren't expressions when they're by themselves on a line).
					if (!left.var->Assign(right)) // left.var can be VAR_VIRTUAL in this case.
						goto abort;
					if (left.var->Type() != VAR_NORMAL // VAR_VIRTUAL should not yield SYM_VAR (as some sections of the code wouldn't handle it correctly).
						|| right.symbol == SYM_MISSING // Subsequent operators/calls (confirmed at load-time as being able to handle `unset`) need SYM_MISSING,
							&& this_postfix[1].symbol != SYM_REF) // except the reference operator, as in &x:=unset.
						this_token.CopyValueFrom(right); // Doing it this way is more maintainable than other methods, and is unlikely to perform much worse.
					else
						this_token.SetVar(left.var);
					goto push_this_token;
				case SYM_ASSIGN_ADD:					this_token.symbol = SYM_ADD; break;
				case SYM_ASSIGN_SUBTRACT:				this_token.symbol = SYM_SUBTRACT; break;
				case SYM_ASSIGN_MULTIPLY:				this_token.symbol = SYM_MULTIPLY; break;
				case SYM_ASSIGN_DIVIDE:					this_token.symbol = SYM_DIVIDE; break;
				case SYM_ASSIGN_INTEGERDIVIDE:			this_token.symbol = SYM_INTEGERDIVIDE; break;
				case SYM_ASSIGN_BITOR:					this_token.symbol = SYM_BITOR; break;
				case SYM_ASSIGN_BITXOR:					this_token.symbol = SYM_BITXOR; break;
				case SYM_ASSIGN_BITAND:					this_token.symbol = SYM_BITAND; break;
				case SYM_ASSIGN_BITSHIFTLEFT:			this_token.symbol = SYM_BITSHIFTLEFT; break;
				case SYM_ASSIGN_BITSHIFTRIGHT:			this_token.symbol = SYM_BITSHIFTRIGHT; break;
				case SYM_ASSIGN_BITSHIFTRIGHT_LOGICAL:	this_token.symbol = SYM_BITSHIFTRIGHT_LOGICAL; break;
				case SYM_ASSIGN_CONCAT:					this_token.symbol = SYM_CONCAT; break;
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
				|| IS_EQUALITY_OPERATOR(this_token.symbol) && !right_is_pure_number && !left_is_pure_number  ) // i.e. if both are strings, compare them alphabetically if the operator supports it.
			{
				// L31: Handle binary ops supported by objects (= == != !==).
				switch (this_token.symbol)
				{
				case SYM_EQUAL:
				case SYM_EQUALCASE:
				case SYM_NOTEQUAL:
				case SYM_NOTEQUALCASE:
					IObject *right_obj = TokenToObject(right);
					IObject *left_obj = TokenToObject(left);
					// To support a future "implicit default value" feature, both operands must be objects.
					// Otherwise, an object operand will be treated as its default value, currently always "".
					// This is also consistent with unsupported operands such as < and > - i.e. because obj<""
					// and obj>"" are always false and obj<="" and obj>="" are always true, obj must be "".
					// When the default value feature is implemented all operators (excluding =, ==, !== and !=
					// if both operands are objects) may use the default value of any object operand.
					// UPDATE: Above is not done because it seems more intuitive to document the other
					// comparison operators as unsupported than for (obj == "") to evaluate to true.
					if (right_obj || left_obj)
					{
						this_token.SetValue((this_token.symbol != SYM_NOTEQUAL && this_token.symbol != SYM_NOTEQUALCASE) == (right_obj == left_obj));
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
				case SYM_EQUAL:	this_token.value_int64 = !_tcsicmp(left_string, right_string); break; // Case-insensitive string comparison.
				
				case SYM_EQUALCASE: // Case sensitive.  Also supports binary data.
					// Support basic equality checking of binary data by using tmemcmp rather than _tcscmp.
					// The results should be the same for strings, but faster.  Length must be checked first
					// since tmemcmp wouldn't stop at the null-terminator (and that's why we're using it).
					// As a result, the comparison is much faster when the length differs.
					this_token.value_int64 = (left_length == right_length) && !tmemcmp(left_string, right_string, left_length);
					break; 
				case SYM_NOTEQUAL:	this_token.value_int64 = !!_tcsicmp(left_string, right_string); break; // Same as SYM_EQUAL but inverted result, code copied for performance (assumed). Note the double !! instead of no !, it is for handling the strcmp functions returning -1.

				case SYM_NOTEQUALCASE:	this_token.value_int64 = !((left_length == right_length)		// Same as SYM_EQUALCASE but inverted result, code copied for performance (assumed).		
										&& !tmemcmp(left_string, right_string, left_length) ); break;
				
				case SYM_CONCAT:
					if (TokenToObject(left))
					{
						error_info = _T("String");
						error_value = &left;
						goto type_mismatch; // Treat this as an error, especially to catch `new classname`.
					}
					if (TokenToObject(right))
					{
						error_info = _T("String");
						error_value = &right;
						goto type_mismatch; // Treat this as an error, especially to catch `new classname`.
					}
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
						this_token.SetVar(sym_assign_var);
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
						&& stack[stack_count-1]->var->Type() == VAR_NORMAL) // Don't do VAR_VIRTUAL here; it mustn't become a SYM_VAR result, so the result would have to be made persistent anyway.
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
									this_token.SetVar(STACK_POP->var);
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
							temp_var->Close(); // Must be called after Assign(NULL, ...) or when Contents() has been altered because it updates the variable's attributes and properly handles VAR_VIRTUAL.
							if (done_and_have_an_output_var) // Fix for v1.0.48: Checking "temp_var == output_var" would not be enough for cases like v := (v := "a" . "b") . "c".
								goto normal_end_skip_output_var; // Nothing more to do because it has even taken care of output_var already.
							else // temp_var is from look-ahead to a future assignment.
							{
								++this_postfix;
								this_token.SetVar(STACK_POP->var);
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
						if (  !(this_token.marker = tmalloc(result_size))  )
							goto outofmem;
						to_free[to_free_count++] = &this_token; // A slot was reserved for this SYM_CONCAT.
					}
					if (left_length)
						tmemcpy(this_token.marker, left_string, left_length);  // Not +1 because don't need the zero terminator.
					tmemcpy(this_token.marker + left_length, right_string, right_length + 1); // +1 to include its zero terminator.
					this_token.marker_length = result_size - 1;
					result_symbol = SYM_STRING;
					break;

				case SYM_IS:
				{
					if (Object *right_obj = dynamic_cast<Object *>(TokenToObject(right)))
					{
						if (IObject *prototype = right_obj->GetOwnPropObj(_T("Prototype")))
						{
							this_token.value_int64 = Object::HasBase(left, prototype);
							break;
						}
					}
					// Since "break" was not used, "right" is not a valid type object.
					error_info = _T("Class");
					error_value = &right;
					goto type_mismatch;
				}

				default:
					// All other operators do not support non-numeric operands.
					error_info = _T("Number");
					error_value = right_is_number ? &left : &right; // Must use right_is_number since if it's false, left_is_number wasn't set.
					goto type_mismatch;
				}
				this_token.symbol = result_symbol; // Must be done only after the switch() above.
			}

			else if (right_is_number == PURE_INTEGER && left_is_number == PURE_INTEGER && this_token.symbol != SYM_DIVIDE)
			{
				// Because both are integers and the operation isn't division, the result is integer.
				right_int64 = TokenToInt64(right); // It can't be SYM_STRING because in here, both right and
				left_int64 = TokenToInt64(left);    // left are known to be numbers (otherwise an earlier "else if" would have executed instead of this one).
				result_symbol = SYM_INTEGER; // Set default.
				switch(this_token.symbol)
				{
				// The most common cases are kept up top to enhance performance if switch() is implemented as if-else ladder.
				case SYM_ADD:			this_token.value_int64 = left_int64 + right_int64; break;
				case SYM_SUBTRACT:		this_token.value_int64 = left_int64 - right_int64; break;
				case SYM_MULTIPLY:		this_token.value_int64 = left_int64 * right_int64; break;
				// A look at K&R confirms that relational/comparison operations and logical-AND/OR/NOT
				// always yield a one or a zero rather than arbitrary non-zero values:
				case SYM_EQUALCASE: // Same behavior as SYM_EQUAL for numeric operands.
				case SYM_EQUAL:			this_token.value_int64 = left_int64 == right_int64; break;
				case SYM_NOTEQUALCASE: // Same behavior as SYM_NOTEQUAL for numeric operands.
				case SYM_NOTEQUAL:		this_token.value_int64 = left_int64 != right_int64; break;
				case SYM_GT:			this_token.value_int64 = left_int64 > right_int64; break;
				case SYM_LT:			this_token.value_int64 = left_int64 < right_int64; break;
				case SYM_GTOE:			this_token.value_int64 = left_int64 >= right_int64; break;
				case SYM_LTOE:			this_token.value_int64 = left_int64 <= right_int64; break;
				case SYM_BITAND:		this_token.value_int64 = left_int64 & right_int64; break;
				case SYM_BITOR:			this_token.value_int64 = left_int64 | right_int64; break;
				case SYM_BITXOR:		this_token.value_int64 = left_int64 ^ right_int64; break;
				case SYM_BITSHIFTLEFT:
				case SYM_BITSHIFTRIGHT:
				case SYM_BITSHIFTRIGHT_LOGICAL:
					if (right_int64 < 0 || right_int64 > 63)
						goto abort_with_exception;
					if (this_token.symbol == SYM_BITSHIFTRIGHT_LOGICAL)
						this_token.value_int64 = (unsigned __int64)left_int64 >> right_int64;
					else
						this_token.value_int64 = this_token.symbol == SYM_BITSHIFTLEFT
						? left_int64 << right_int64
						: left_int64 >> right_int64;
					break;
				case SYM_INTEGERDIVIDE:
					if (right_int64 == 0)
						goto divide_by_zero;
					this_token.value_int64 = left_int64 / right_int64;
					break;
				case SYM_POWER:
					if (!left_int64 && right_int64 < 0) // In essence, this is divide-by-zero.
					{
						// Throw an exception rather than returning something undefined:
						goto divide_by_zero;
					}
					else if (left_int64 == 0 && right_int64 == 0)	// 0**0, not defined.
					{
						goto abort_with_exception;
					}
					else // We have a valid base and exponent and both are integers, so the calculation will always have a defined result.
					{
						if (right_int64 >= 0)	// result is an integer.
						{
							this_token.value_int64 = pow_ll(left_int64, right_int64);
							break;
						}
						result_symbol = SYM_FLOAT; // Due to negative exponent, override to float.
#ifdef USE_INLINE_ASM	// see qmath.h
						// Note: The function pow() in math.h adds about 28 KB of code size (uncompressed)! That is why it's not used here.
						// v1.0.44.11: With Laszlo's help, negative integer bases are now supported.
						if (left_was_negative = (left_int64 < 0))
							left_int64 = -left_int64; // Force a positive due to the limitations of qmathPow().
						this_token.value_double = qmathPow((double)left_int64, (double)right_int64);
						if (left_was_negative && right_int64 % 2) // Negative base and odd exponent (not zero or even).
							this_token.value_double = -this_token.value_double;
#else
						this_token.value_double = pow((double)left_int64, (double)right_int64);
#endif
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
					if (right_double == 0.0)
						goto divide_by_zero;
					this_token.value_double = left_double / right_double;
					break;
				case SYM_EQUALCASE: // Same behavior as SYM_EQUAL for numeric operands.
				case SYM_EQUAL:    this_token.value_int64 = left_double == right_double; break;
				case SYM_NOTEQUALCASE: // Same behavior as SYM_NOTEQUAL for numeric operands.
				case SYM_NOTEQUAL: this_token.value_int64 = left_double != right_double; break;
				case SYM_GT:       this_token.value_int64 = left_double > right_double; break;
				case SYM_LT:       this_token.value_int64 = left_double < right_double; break;
				case SYM_GTOE:     this_token.value_int64 = left_double >= right_double; break;
				case SYM_LTOE:     this_token.value_int64 = left_double <= right_double; break;
				case SYM_POWER:
					if (left_double == 0.0 && right_double < 0)  // In essence, this is divide-by-zero.
						goto divide_by_zero;
					left_was_negative = (left_double < 0);
					if ((left_was_negative && qmathFmod(right_double, 1.0) != 0.0)	// Negative base, but exponent isn't close enough to being an integer: unsupported (to simplify code).
						|| (left_double == 0.0 && right_double == 0.0))				// 0.0**0.0, not defined.
						goto abort_with_exception;
#ifdef USE_INLINE_ASM	// see qmath.h
					// v1.0.44.11: With Laszlo's help, negative bases are now supported as long as the exponent is not fractional.
					// See the other SYM_POWER higher above for more details about below.
					if (left_was_negative)
						left_double = -left_double; // Force a positive due to the limitations of qmathPow().
					this_token.value_double = qmathPow(left_double, right_double);
					if (left_was_negative && qmathFabs(qmathFmod(right_double, 2.0)) == 1.0) // Negative base and exactly-odd exponent (otherwise, it can only be zero or even because if not it would have returned higher above).
						this_token.value_double = -this_token.value_double;
#else
					this_token.value_double = pow(left_double, right_double);
#endif
					break;
				default:
					if (IS_INTEGER_OPERATOR(this_token.symbol))
					{
						error_info = _T("Integer");
						error_value = left_is_number == PURE_INTEGER ? &right : &left;
						goto type_mismatch; // floats are not supported for the integer operators.
					}
					// this is should not be reachable.
#ifdef _DEBUG
					LineError(_T("Unhandled float operation.")); // To help catch bugs.
					goto abort;
#endif
				} // switch(this_token.symbol)
				this_token.symbol = result_symbol; // Must be done only after the switch() above.
			} // Result is floating point.
		} // switch() operator type

		if (sym_assign_var) // Added in v1.0.46. There are some places higher above that handle sym_assign_var themselves and skip this section via goto.
		{
			if (!sym_assign_var->Assign(this_token)) // Assign the result (based on its type) to the target variable.
				goto abort;
			if (sym_assign_var->Type() == VAR_NORMAL)
				this_token.SetVar(sym_assign_var);
			//else its VAR_VIRTUAL, so just push this_token as-is because after its assignment is done,
			// it should no longer be a SYM_VAR.  This is done to simplify the code, such as BIFs.
			//
			// Now fall through and push this_token onto the stack as an operand for use by future operators.
			// This is because by convention, an assignment like "x+=1" produces a usable operand.
		}

push_this_token:
		ASSERT(stack_count < mArg[aArgIndex].max_stack);
		STACK_PUSH(&this_token);   // Push the result onto the stack for use as an operand by a future operator.
	} // For each item in the postfix array.

	if (stack_count != 1) // Even for multi-statement expressions, the stack should have only one item left on it:
		goto abort_with_exception; // the overall result.  Any conditions that cause this *should* be detected at load time.

	ExprTokenType &result_token = *stack[0];  // For performance and convenience.  Even for multi-statement, the bottommost item on the stack is the final result so that things like var1:=1,var2:=2 work.

	// Although ACT_EXPRESSION was already checked higher above for function calls, there are other ways besides
	// an isolated function call to have ACT_EXPRESSION.  For example: var&=3 (where &= is an operator that lacks
	// a corresponding command).  Another example: true ? fn1() : fn2()
	// Also, there might be ways the function-call section didn't return for ACT_EXPRESSION, such as when somehow
	// there was more than one token on the stack even for the final function call, or maybe other unforeseen ways.
	// It seems best to avoid any chance of looking at the result since it might be invalid due to the above
	// having taken shortcuts (since it knew the result would be discarded).
	if (mActionType == ACT_EXPRESSION)   // A stand-alone expression whose end result doesn't matter.
		goto normal_end_skip_output_var; // Can't be any output_var for this action type. Also, leave result_to_return at its default of "".

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
		case SYM_MISSING: // return unset
			// Return numeric or object result as-is.
			aResultToken->symbol = result_token.symbol;
			aResultToken->value_int64 = result_token.value_int64; // Union copy.
			if (result_token.symbol == SYM_OBJECT)
				result_token.object->AddRef();
			goto normal_end_skip_output_var; // result_to_return is left at its default of "".
		case SYM_VAR:
			if (result_token.var->IsPureNumericOrObject())
			{
				result_token.var->ToToken(*aResultToken);
				goto normal_end_skip_output_var; // result_to_return is left at its default of "".
			}
			// The following "optimizations" are avoided because of hidden complexity, most notably
			// unwanted consequences with try-finally (which allows the script to read or modify the
			// variable after we return):
			//  - Return a local variable's string buffer directly, detaching it from the variable.
			//    The idea was that the variable is about to be freed anyway, but there are exceptions.
			//  - Return a static or global variable's string directly.
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
	
	if (result_token.symbol == SYM_MISSING) // No valid cases permit this as a final result except those already handled above.  Some sections below might not handle it.
		goto abort_with_exception;

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
	default:
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
			result_length = result_token.var->Length();
		}
		else
		{
			result = result_token.marker;
			result_length = result_token.marker_length; // At this stage, marker_length should always be valid, not -1.
		}
		result_size = result_length + 1;

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
				goto outofmem;
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
			tmemcpy(aTarget, result, result_length); // Copy from old location to the newly allocated one.
			aTarget[result_length] = '\0'; // Guarantee null-termination so it doesn't have to be done at an earlier stage.

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
		{
			tmemmove(aTarget, result, result_length); // memmove() vs. memcpy() in this case, since source and dest might overlap (i.e. "target" may have been used to put temporary things into aTarget, but those things are no longer needed and now safe to overwrite).
			aTarget[result_length] = '\0'; // Guarantee null-termination so it doesn't have to be done at an earlier stage.
		}
		if (aResultToken)
		{
			aResultToken->marker = aTarget;
			aResultToken->marker_length = result_length;
		}
		aTarget += result_size;
		goto normal_end_skip_output_var; // output_var was already checked higher above, so no need to consider it again.

	case SYM_OBJECT:
		// At this point we aren't capable of returning an object, otherwise above would have
		// already returned.  So in other words, the caller wants a string, not an object.
		error_info = _T("String");
		error_value = &result_token;
		goto type_mismatch;
	} // switch (result_token.symbol)

// ALL PATHS ABOVE SHOULD "GOTO".  TO CATCH BUGS, ANY THAT DON'T FALL INTO "ABORT" BELOW.
abort_with_exception:
	aResult = LineError(error_msg, FAIL_OR_OK, error_info);
	// FALL THROUGH:
abort_if_result:
	if (aResult != FAIL)
	{
		if (aResultToken)
			aResultToken->symbol = SYM_MISSING;
		goto normal_end_skip_output_var;
	}
	// FALL THROUGH:
abort:
	// The callers of this function know that the value of aResult (which contains the reason
	// for early exit) should be considered valid/meaningful only if result_to_return is NULL.
	result_to_return = NULL; // Use NULL to inform our caller that this entire thread is to be terminated.
	aResult = FAIL; // Indicate reason to caller.
	goto normal_end_skip_output_var; // output_var is skipped as part of standard abort behavior.

type_mismatch:
	{
		ResultToken temp_result;
		temp_result.SetResult(OK);
		temp_result.TypeError(error_info, *error_value);
		aResult = temp_result.Result(); // FAIL to abort, OK if user or OnError requested continuation.
		goto abort_if_result;
	}
divide_by_zero:
	aResult = g_script.RuntimeError(ERR_DIVIDEBYZERO, nullptr, FAIL_OR_OK, this, ErrorPrototype::ZeroDivision);
	goto abort_if_result;
outofmem:
	aResult = MemoryError();
	goto abort_if_result;
unset_var:
	aResult = g_script.VarUnsetError(error_value->var);
	goto abort_if_result;

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



ResultType Line::ExpandSingleArg(int aArgIndex, ResultToken &aResultToken, LPTSTR &aDerefBuf, size_t &aDerefBufSize)
{
	ExprTokenType *postfix = mArg[aArgIndex].postfix;
	if (postfix->symbol < SYM_DYNAMIC // i.e. any other operand type.
		&& postfix->symbol != SYM_VAR // Variables must be dereferenced.
		&& postfix[1].symbol == SYM_INVALID) // Exactly one token.
	{
		aResultToken.symbol = postfix->symbol;
		aResultToken.value_int64 = postfix->value_int64;
#ifdef _WIN64
		aResultToken.marker_length = postfix->marker_length;
#endif
		if (aResultToken.symbol == SYM_OBJECT) // This can happen for VAR_CONSTANT.
			aResultToken.object->AddRef();
		return OK;
	}

	size_t space_needed = EXPR_BUF_SIZE(mArg[aArgIndex].length);

	if (aDerefBufSize < space_needed)
	{
		if (aDerefBuf)
		{
			free(aDerefBuf);
			if (aDerefBufSize > LARGE_DEREF_BUF_SIZE)
				--sLargeDerefBufs;
		}
		if ( !(aDerefBuf = tmalloc(space_needed)) )
		{
			aDerefBufSize = 0;
			return MemoryError();
		}
		aDerefBufSize = space_needed;
		if (aDerefBufSize > LARGE_DEREF_BUF_SIZE)
			++sLargeDerefBufs;
	}

	// Use the whole buf:
	LPTSTR buf_marker = aDerefBuf;
	size_t extra_size = aDerefBufSize - space_needed;

	// None of the previous args (if any) are needed, so pass an array of NULLs:
	LPTSTR arg_deref[MAX_ARGS];
	for (int i = 0; i < aArgIndex; i++)
		arg_deref[i] = NULL;

	// Initialize aResultToken so we can detect when ExpandExpression() uses it:
	aResultToken.symbol = SYM_INVALID;
	
	ResultType result;
	LPTSTR string_result = ExpandExpression(aArgIndex, result, &aResultToken, buf_marker, aDerefBuf, aDerefBufSize, arg_deref, extra_size);
	if (!string_result)
		return result; // Should be EARLY_EXIT or FAIL.

	if (aResultToken.symbol == SYM_INVALID) // It wasn't set by ExpandExpression().
	{
		aResultToken.symbol = SYM_STRING;
		aResultToken.marker = string_result;
	}
	return OK;
}



ResultType VariadicCall(IObject *aObj, IObject_Invoke_PARAMS_DECL)
{
	IObject *param_obj = nullptr; // Vararg object passed by caller.
	Array *param_array = nullptr; // Array of parameters, either the same as param_obj or the result of enumeration.
	ExprTokenType* token = nullptr; // Used for variadic calls.

	ExprTokenType *rvalue = IS_INVOKE_SET ? aParam[--aParamCount] : nullptr;

	--aParamCount; // Exclude param_obj from aParamCount, so it's the count of normal params.
	param_obj = TokenToObject(*aParam[aParamCount]);
	// It might be more correct to use the enumerator even for Array, but that could be slow.
	// Future changes might enable efficient detection of a custom __Enum method, allowing
	// us to take the more efficient path most times, but still support custom enumeration.
	if (param_array = dynamic_cast<Array *>(param_obj))
		param_array->AddRef();
	else
		if (!(param_array = Array::FromEnumerable(*aParam[aParamCount])))
			return aResultToken.SetExitResult(FAIL);
	int extra_params = param_array->Length();
	if (extra_params > 0)
	{
		// Calculate space required for ...
		size_t space_needed = extra_params * sizeof(ExprTokenType) // ... new param tokens
			+ (aParamCount + extra_params) * sizeof(ExprTokenType *); // ... existing and new param pointers
		if (rvalue)
			space_needed += sizeof(rvalue); // ... extra slot for aRValue
		// Allocate new param list and tokens; tokens first for convenience.
		token = (ExprTokenType *)_malloca(space_needed);
		if (!token)
			return aResultToken.MemoryError();
		ExprTokenType **param_list = (ExprTokenType **)(token + extra_params);
		// Since built-in functions don't have variables we can directly assign to,
		// we need to expand the param object's contents into an array of tokens:
		param_array->ToParams(token, param_list, aParam, aParamCount);
		aParam = param_list;
		aParamCount += extra_params;
	}
	if (rvalue)
		aParam[aParamCount++] = rvalue; // In place of the variadic param.

#ifdef ENABLE_HALF_BAKED_NAMED_PARAMS
	aResultToken.named_params = param_obj;
#endif

	auto result = aObj->Invoke(aResultToken, aFlags, aName, aThisToken, aParam, aParamCount);

	if (param_array)
		param_array->Release();
	if (token)
		_freea(token);
	
	return result;
}



bool Func::Call(ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	if (aParamCount > mParamCount && !mIsVariadic) // v2 policy.
	{
		for ( ; aParamCount && aParam[aParamCount - 1]->symbol == SYM_MISSING; --aParamCount);
		if (aParamCount > mParamCount)
		{
			aResultToken.Error(ERR_TOO_MANY_PARAMS, mName);
			return false;
		}
	}
	return true;
}

bool NativeFunc::Call(ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	if (!Func::Call(aResultToken, aParam, aParamCount))
		return false;

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
		
		for (int i = 0; i < mMinParams; ++i)
		{
			if (aParam[i]->symbol == SYM_MISSING
				&& !ArgIsOptional(i)) // BuiltInMethod requires this exception for some setters.
			{
				aResultToken.Error(ERR_PARAM_REQUIRED);
				return false; // Abort expression.
			}
		}

		if (mOutputVars)
		{
			// Verify that each output parameter is either a valid var or completely omitted.
			for (int i = 0; i < MAX_FUNC_OUTPUT_VAR && mOutputVars[i]; ++i)
			{
				if (mOutputVars[i] <= aParamCount
					&& aParam[mOutputVars[i]-1]->symbol != SYM_MISSING
					&& !TokenToOutputVar(*aParam[mOutputVars[i]-1]))
				{
					aResultToken.ParamError(mOutputVars[i]-1, aParam[mOutputVars[i]-1], _T("variable reference"), mName);
					return false; // Abort expression.
				}
			}
		}

	return true;
}

bool BuiltInFunc::Call(ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	if (!NativeFunc::Call(aResultToken, aParam, aParamCount))
		return false;

	aResultToken.func = this; // Inform function of which built-in function called it (allows code sharing/reduction).

		// Push an entry onto the debugger's stack.  This has two purposes:
		//  1) Allow CreateRuntimeException() to know which function is throwing an exception.
		//  2) If a UDF is called before the BIF returns, it will show on the call stack.
		//     e.g. DllCall(RegisterCallback("F")) will show DllCall while F is running.
		DEBUGGER_STACK_PUSH(this)

		aResultToken.symbol = SYM_INTEGER; // Set default return type so that functions don't have to do it if they return INTs.
		mBIF(aResultToken, aParam, aParamCount);

		DEBUGGER_STACK_POP()
		
		// There shouldn't be any need to check g->ThrownToken since built-in functions
		// currently throw exceptions via aResultToken.Error():
		//if (g->ThrownToken)
		//	aResultToken.SetExitResult(FAIL); // Abort thread.

	return !aResultToken.Exited();
}

bool BuiltInMethod::Call(ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	if (!NativeFunc::Call(aResultToken, aParam, aParamCount))
		return false;

	DEBUGGER_STACK_PUSH(this) // See comments in BuiltInFunc::Call.

	auto obj = TokenToObject(*aParam[0]);
	// mClass is the prototype object of the class which this method is for, such as
	// Object.Prototype or Map.Prototype.  IsOfType() takes care of ruling out derived
	// prototype objects, which are always really just Object.
	if (!obj || !obj->IsOfType(mClass))
	{
		LPCTSTR expected_type;
		ExprTokenType value;
		if (mClass->GetOwnProp(value, _T("__Class")) && value.symbol == SYM_STRING)
			expected_type = value.marker;
		else
			expected_type = _T("?"); // Script may have tampered with the prototype.
		aResultToken.TypeError(expected_type, *aParam[0]);
	}
	else
		(obj->*mBIM)(aResultToken, mMID, mMIT, aParam + 1, aParamCount - 1);

	DEBUGGER_STACK_POP()

	return !aResultToken.Exited();
}

bool UserFunc::Call(ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	return Call(aResultToken, aParam, aParamCount, nullptr);
}

bool UserFunc::Call(ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount, FreeVars *aUpVars)
{
	if (!Func::Call(aResultToken, aParam, aParamCount))
		return false;

		UDFCallInfo recurse(this);

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
			for (j = 0; j < aParamCount; ++j) // For each actual parameter.
			{
				ExprTokenType &this_param_token = *aParam[j];
				if (this_param_token.symbol != SYM_VAR
					|| VARREF_IS_WRITE(this_param_token.var_usage)) // VARREF_REF indicates SYM_VAR is being passed ByRef.
					continue;
				// Since this SYM_VAR is being passed by value, convert it to a non-var to allow
				// the variables to be backed up and reset further below without corrupting any
				// SYM_VARs that happen to be locals or params of this very same function.
				// Skip AddRef() if this is an object because Release() won't be called, and
				// AddRef() will be called when the object is assigned to a parameter.
				this_param_token.var->ToTokenSkipAddRef(this_param_token);
			}
			// BackupFunctionVars() will also clear each local variable and formal parameter so that
			// if that parameter or local var is assigned a value by any other means during our call
			// to it, new memory will be allocated to hold that value rather than overwriting the
			// underlying recursed/interrupted instance's memory, which it will need intact when it's resumed.
			if (!Var::BackupFunctionVars(*this, recurse.backup, recurse.backup_count)) // Out of memory.
			{
				aResultToken.MemoryError();
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
		
		FreeVars *caller_free_vars = sFreeVars;
		if (sFreeVars && mOuterFunc && !aUpVars)
			aUpVars = sFreeVars->ForFunc(mOuterFunc);

		if (mDownVarCount)
		{
			// These local vars need to persist after the function returns (and be independent of
			// any other instances of this function).  Since we really only have one set of local
			// vars for the lifetime of the script, make them aliases for newly allocated vars:
			sFreeVars = FreeVars::Alloc(*this, mDownVarCount, aUpVars);
			for (int i = 0; i < mDownVarCount; ++i)
				mDownVar[i]->SetAliasDirect(sFreeVars->mVar + i);
		}
		else
			sFreeVars = NULL;
		
		if (mUpVarCount)
		{
			if (!aUpVars)
			{
				// No aUpVars, so it must be a direct call, and mOuterFunc wasn't found in the sFreeVars
				// linked list, so it's probably a direct call from something which doesn't support closures,
				// occurring after mOuterFunc returned.
				aResultToken.Error(_T("Func out of scope."), mName); // Keep it short since this shouldn't be possible once the implementation is complete.
				goto free_and_return;
			}
			for (int i = 0; i < mUpVarCount; ++i)
			{
				Var *outer_free_var = aUpVars->mVar + mUpVarIndex[i];
				if (mUpVar[i]->Scope() & VAR_DOWNVAR) // This is both an upvar and a downvar.
				{
					Var *inner_free_var = mUpVar[i]->ResolveAlias(); // Retrieve the alias which was just set above.
					inner_free_var->UpdateAlias(outer_free_var); // Point the free var of our layer to the outer one for use by closures within this function.
					// mUpVar[i] is now a two-level alias (mUpVar[i] -> inner_free_var -> outer_free_var),
					// but that will be corrected below.  Technically outer_free_var might also be an alias,
					// in which case inner_free_var is now an alias for outer_free_var->mAliasFor.
				}
				mUpVar[i]->UpdateAlias(outer_free_var);
			}
		}

		if (mClosureCount)
		{
			ASSERT(sFreeVars);
			for (int i = 0; i < mClosureCount; ++i)
			{
				auto closure = new Closure(mClosure[i].func, sFreeVars
					, mClosure[i].var->Scope() & VAR_DOWNVAR); // Closures in downvars have lifetime tied to sFreeVars.
				Var *var = mClosure[i].var->ResolveAlias();
				var->AssignSkipAddRef(closure);
				var->MakeReadOnly();
			}
		}

		int default_expr = mParamCount;
		for (j = 0; j < mParamCount; ++j) // For each formal parameter.
		{
			FuncParam &this_formal_param = mParam[j]; // For performance and convenience.

			// Assignments below rely on ByRef parameters having already been reset to VAR_NORMAL
			// by Free() or Backup(), except when it's a downvar, which should be VAR_ALIAS.
			ASSERT((this_formal_param.var->Scope() & VAR_DOWNVAR) ? (this_formal_param.var->ResolveAlias()->Scope() & ~VAR_VARREF) == 0
				: !this_formal_param.var->IsAlias());

			if (j >= aParamCount || aParam[j]->symbol == SYM_MISSING)
			{
#ifdef ENABLE_HALF_BAKED_NAMED_PARAMS
				if (aResultToken.named_params)
				{
					FuncResult rt_item;
					ExprTokenType t_this(aResultToken.named_params);
					auto r = aResultToken.named_params->Invoke(rt_item, IT_GET, this_formal_param.var->mName, t_this, nullptr, 0);
					if (r == FAIL || r == EARLY_EXIT)
					{
						aResultToken.SetExitResult(r);
						goto free_and_return;
					}
					if (r != INVOKE_NOT_HANDLED)
					{
						this_formal_param.var->Assign(rt_item);
						rt_item.Free();
						continue;
					}
				}
#endif
			
				switch(this_formal_param.default_type)
				{
				case PARAM_DEFAULT_STR:   this_formal_param.var->Assign(this_formal_param.default_str);    break;
				case PARAM_DEFAULT_INT:   this_formal_param.var->Assign(this_formal_param.default_int64);  break;
				case PARAM_DEFAULT_FLOAT: this_formal_param.var->Assign(this_formal_param.default_double); break;
				case PARAM_DEFAULT_UNSET: this_formal_param.var->MarkUninitialized(); break;
				case PARAM_DEFAULT_EXPR:
					this_formal_param.var->MarkUninitialized();
					if (default_expr > j)
						default_expr = j; // Take note of the first param with a default expr.
					break;
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
				if (token.symbol == SYM_VAR && VARREF_IS_WRITE(token.var_usage)) // An optimized &var ref.
				{
					if (this_formal_param.var->Scope() & VAR_DOWNVAR) // This parameter's var is referenced by one or more closures.
					{
						// Make sure the caller's var (token.var) points to a ref-counted freevar.
						auto ref = token.var->GetRef();
						if (!ref)
							goto free_and_return;
						ref->Release(); // token.var retains a reference; release ours.
						ASSERT(this_formal_param.var->IsAlias());
						// Point our freevar to the caller's freevar, for use by our closures.
						this_formal_param.var->GetAliasFor()->UpdateAlias(token.var);
						// Also update our local alias below.
					}
					this_formal_param.var->UpdateAlias(token.var); // Set mAliasFor.
					continue;
				}
				else if (auto ref = dynamic_cast<VarRef *>(TokenToObject(token)))
				{
					if (this_formal_param.var->Scope() & VAR_DOWNVAR) // This parameter's var is referenced by one or more closures.
					{
						ASSERT(this_formal_param.var->IsAlias());
						// Point our freevar to the caller's freevar, for use by our closures.
						this_formal_param.var->GetAliasFor()->UpdateAlias(ref);
						// Also update our local alias below.
					}
					this_formal_param.var->UpdateAlias(ref); // Set mAliasFor and mObject.
					continue;
				}
				aResultToken.ParamError(j - (mClass ? 1 : 0), &token, _T("variable reference"), mName);
				goto free_and_return;
			}
			//else // This parameter is passed "by value".
			// Assign actual parameter's value to the formal parameter (which is itself a
			// local variable in the function).  
			// token.var's Type() is always VAR_NORMAL (never a built-in virtual variable).
			// A SYM_VAR token can still happen because the previous loop's conversion of all
			// by-value SYM_VAR operands into the appropriate operand symbol would not have
			// happened if no backup was needed for this function (which is usually the case).
			if (!this_formal_param.var->Assign(token))
			{
				aResultToken.SetExitResult(FAIL); // Abort thread.
				goto free_and_return;
			}
		} // for each formal parameter.
		
		if (mIsVariadic && mParam[mParamCount].var) // i.e. this function is capable of accepting excess params via an object/array.
		{
			// Unused named parameters in param_obj are currently discarded, pending completion of the
			// object redesign.  Ultimately the named parameters (either key-value pairs or properties)
			// would be enumerated and added to vararg_obj or passed to the function some other way.
			auto vararg_obj = Array::Create();
			if (!vararg_obj)
			{
				aResultToken.MemoryError();
				goto free_and_return;
			}
			if (j < aParamCount)
				// Insert the excess parameters from the actual parameter list.
				vararg_obj->InsertAt(0, aParam + j, aParamCount - j);
			// Assign to the "param*" var:
			mParam[mParamCount].var->AssignSkipAddRef(vararg_obj);
		}

		DEBUGGER_STACK_PUSH(&recurse)

		ResultType result = OK;
		// Execute any default initializers that weren't simple constants.  This is not done in
		// the loop above for two reasons:
		//  1) It needs to be after DEBUGGER_STACK_PUSH (which isn't moved because it probably
		//     doesn't make sense for the other errors to include this function in the stack trace).
		//  2) To preserve the pre-v2.0.8 behaviour, which allows an initializer to refer to later
		//     parameters if they are simple values.
		for (j = default_expr; j < mParamCount; ++j)
		{
			if (j < aParamCount && aParam[j]->symbol != SYM_MISSING || mParam[j].default_type != PARAM_DEFAULT_EXPR)
				continue;
			result = mParam[j].default_expr->ExecUntil(ONLY_ONE_LINE);
			if (result != OK)
			{
				aResultToken.SetExitResult(result);
				break;
			}
		}

		if (result == OK)
			result = Execute(&aResultToken); // Execute the body of the function.

		DEBUGGER_STACK_POP()
		
		// Setting this unconditionally isn't likely to perform any worse than checking for EXIT/FAIL,
		// and likely produces smaller code.  Execute() takes care of translating EARLY_RETURN to OK.
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
		Var::FreeAndRestoreFunctionVars(*this, recurse.backup, recurse.backup_count);

		// mInstances must remain non-zero until this point to ensure that any recursive calls by an
		// object's __Delete meta-function receive fresh variables, and none partially-destructed.
		--mInstances;

		if (sFreeVars)
			sFreeVars->Release();
		sFreeVars = caller_free_vars;
	return !aResultToken.Exited(); // i.e. aResultToken.SetExitResult() or aResultToken.Error() was not called.
}

FreeVars *UserFunc::sFreeVars = nullptr;



ResultType Line::ExpandArgs(ResultToken *aResultTokens)
// aResultTokens is non-null if the caller wants results that aren't strings.  Caller is
// responsible for freeing the tokens either after successful use or before aborting, on failure.
// Returns OK, FAIL, or EARLY_EXIT.  EARLY_EXIT occurs when a function-call inside an expression
// used the EXIT command to terminate the thread.
{
	// The counterpart of sArgDeref kept on our stack to protect it from recursion caused by
	// the calling of functions in the script:
	LPTSTR arg_deref[MAX_ARGS];
	int i;

	// Make two passes through this line's arg list.  This is done because the performance of
	// realloc() is worse than doing a free() and malloc() because the former often does a memcpy()
	// in addition to the latter's steps.  In addition, realloc() as much as doubles the memory
	// load on the system during the brief time that both the old and the new blocks of memory exist.
	// First pass: determine how much space will be needed to do all the args and allocate
	// more memory if needed.  Second pass: dereference the args into the buffer.

	// First pass. It takes into account the same things as 2nd pass.
	size_t space_needed = GetExpandedArgSize();
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
			return MemoryError();
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

			if (this_arg.type != ARG_TYPE_INPUT_VAR)
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
			// aResultTokens should currently always be NULL for ARG_TYPE_INPUT_VAR, so isn't filled.

			if (VAR(mArg[i])->IsUninitializedNormalVar())
			{
				result_to_return = g_script.VarUnsetError(VAR(mArg[i]));
				goto end;
			}

			// Some comments below might be obsolete.  There was previously some logic here deciding
			// whether the variable needed to be dereferenced, but that's no longer ever necessary.
			// Previous stages have ensured the_only_var_of_this_arg is of type VAR_NORMAL.
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
				) ? _T("") : NULL; // See "Update #2" and later comments above.
			
		} // for each arg.

		// See "Update #3" comment above.  This must be done separately to the loop below since Contents()
		// may cause a warning dialog, which in turn may cause a new thread to launch, thus potentially
		// corrupting sArgDeref.
		for (i = 0; i < mArgc; ++i)
			if (arg_deref[i] == NULL)
				arg_deref[i] = VAR(mArg[i])->Contents();

		// IT'S NOT SAFE to do the following until the above loops FULLY complete because any calls made above to
		// ExpandExpression() might call functions, which in turn might result in a recursive call to ExpandArgs(),
		// which in turn might change the values in the static array sArgDeref.
		// Also, only when the loop ends normally is the following needed, since otherwise it's a failure condition.
		// Now that any recursive calls to ExpandArgs() above us on the stack have collapsed back to us, it's
		// safe to set the args of this command for use by our caller, to whom we're about to return.
		for (i = 0; i < mArgc; ++i) // Copying actual/used elements is probably faster than using memcpy to copy both entire arrays.
			sArgDeref[i] = arg_deref[i];
	} // redundant block

	// v1.0.40.02: The following loop was added to avoid the need for the ARGn macros to provide an empty
	// string when mArgc was too small (indicating that the parameter is absent).  This saves quite a bit
	// of code size.  Also, the slight performance loss caused by it is partially made up for by the fact
	// that all the other sections don't need to check mArgc anymore.
	// Benchmarks show that it doesn't help performance to try to tweak this with a pre-check such as
	// "if (mArgc < max_params)":
	int max_params = g_act[mActionType].MaxParams; // Resolve once for performance.
	for (i = mArgc; i < max_params; ++i) // START AT mArgc.  For performance, this only does the actual max args for THIS command, not MAX_ARGS.
		sArgDeref[i] = _T("");

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
	{
		// SetTimer has a cost that adds up very quickly if ExpandArgs() is called in a tight loop
		// (potentially thousands or millions of times per second).  There's no need for the timer
		// to be precise, so don't reset it more often than twice every second.  (Even checking
		// now != sLastTimerReset is sufficient.)
		static DWORD sLastTimerReset = 0;
		DWORD now = GetTickCount();
		if (now - sLastTimerReset > 500 || !g_DerefTimerExists)
		{
			sLastTimerReset = now;
			SET_DEREF_TIMER(10000) // Reset the timer right before the deref buf is possibly about to become idle.
		}
	}

	return result_to_return;
}

	

VarSizeType Line::GetExpandedArgSize()
// Returns the size, or VARSIZE_ERROR if there was a problem.
// This function can return a size larger than what winds up actually being needed
// (e.g. caused by ScriptGetCursor()), so our callers should be aware that that can happen.
{
	int i;
	VarSizeType space_needed;
	
	// Note: the below loop is similar to the one in ExpandArgs(), so the two should be maintained together:
	for (i = 0, space_needed = 0; i < mArgc; ++i) // FOR EACH ARG:
	{
		ArgStruct &this_arg = mArg[i]; // For performance and convenience.
		
		if (this_arg.is_expression)
		{
			// The length used below is more room than is strictly necessary, but given how little
			// space is typically wasted (and that only while the expression is being evaluated),
			// it doesn't seem worth worrying about it.  See other comments at macro definition.
			space_needed += EXPR_BUF_SIZE(this_arg.length);
		}
		// Since is_expression is false, it must be plain text or a non-dynamic input/output var.
	}

	return space_needed;
}


