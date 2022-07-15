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

#include "stdafx.h" // pre-compiled headers
#include <winioctl.h> // For PREVENT_MEDIA_REMOVAL and CD lock/unlock.
#include "qmath.h" // Used by Transform() [math.h incurs 2k larger code size just for ceil() & floor()]
#include "script.h"
#include "window.h" // for IF_USE_FOREGROUND_WINDOW
#include "application.h" // for MsgSleep()
#include "resources/resource.h"  // For InputBox.
#include "TextIO.h"
#include <Psapi.h> // for GetModuleBaseName.
#include "shlwapi.h" // StrCmpLogicalW

#include <mmdeviceapi.h> // for SoundSet/SoundGet.
#include <endpointvolume.h> // for SoundSet/SoundGet.
#include <functiondiscoverykeys.h>

#define PCRE_STATIC             // For RegEx. PCRE_STATIC tells PCRE to declare its functions for normal, static
#include "lib_pcre/pcre/pcre.h" // linkage rather than as functions inside an external DLL.

#include "script_func_impl.h"



////////////////////
// Math Functions //
////////////////////


BIF_DECL(BIF_Round)
// For simplicity, this always yields something numeric (or a string that's numeric).
// Even Round(empty_or_unintialized_var) is zero rather than "" or "NaN".
{
	// In the future, a string conversion algorithm might be better to avoid the loss
	// of 64-bit integer precision that is currently caused by the use of doubles in
	// the calculation:
	int param2;
	double multiplier;
	if (aParamCount > 1)
	{
		Throw_if_Param_NaN(1);
		param2 = ParamIndexToInt(1);
		multiplier = qmathPow(10, param2);
	}
	else // Omitting the parameter is the same as explicitly specifying 0 for it.
	{
		param2 = 0;
		multiplier = 1;
	}
	Throw_if_Param_NaN(0);
	double value = ParamIndexToDouble(0);
	value = (value >= 0.0 ? qmathFloor(value * multiplier + 0.5)
		: qmathCeil(value * multiplier - 0.5)) / multiplier;

	// If incoming value is an integer, it seems best for flexibility to convert it to a
	// floating point number whenever the second param is >0.  That way, it can be used
	// to "cast" integers into floats.  Conversely, it seems best to yield an integer
	// whenever the second param is <=0 or omitted.
	if (param2 > 0)
	{
		// v1.0.44.01: Since Round (in its param2>0 mode) is almost always used to facilitate some kind of
		// display or output of the number (hardly ever for intentional reducing the precision of a floating
		// point math operation), it seems best by default to omit only those trailing zeroes that are beyond
		// the specified number of decimal places.  This is done by converting the result into a string here,
		// which will cause the expression evaluation to write out the final result as this very string as long
		// as no further floating point math is done on it (such as Round(3.3333, 2)+0).  Also note that not
		// all trailing zeros are removed because it is often the intent that exactly the number of decimal
		// places specified should be *shown* (for column alignment, etc.).  For example, Round(3.5, 2) should
		// be 3.50 not 3.5.  Similarly, Round(1, 2) should be 1.00 not 1 (see above comment about "casting" for
		// why.
		// Performance: This method is about twice as slow as the old method (which did merely the line
		// "aResultToken.symbol = SYM_FLOAT" in place of the below).  However, that might be something
		// that can be further optimized in the caller (its calls to _tcslen, memcpy, etc. might be optimized
		// someday to omit certain calls when very simply situations allow it).  In addition, twice as slow is
		// not going to impact the vast majority of scripts since as mentioned above, Round (in its param2>0
		// mode) is almost always used for displaying data, not for intensive operations within a expressions.
		// AS DOCUMENTED: Round(..., positive_number) displays exactly positive_number decimal places, which
		// might be inconsistent with normal float->string conversion.  If it wants, the script can force the
		// result to be formatted the normal way (omitting trailing 0s) by adding 0 to the result.
		// Also, a new parameter an be added someday to trim excess trailing zeros from param2>0's result
		// (e.g. Round(3.50, 2, true) can be 3.5 rather than 3.50), but this seems less often desired due to
		// column alignment and other goals where consistency is important.
		LPTSTR buf = _f_retval_buf;
		int len = _stprintf(buf, _T("%0.*f"), param2, value); // %f can handle doubles in MSVC++.
		_f_return_p(buf, len);
	}
	else
		// Fix for v1.0.47.04: See BIF_FloorCeil() for explanation of this fix.  Currently, the only known example
		// of when the fix is necessary is the following script in release mode (not debug mode):
		//   myNumber  := 1043.22  ; Bug also happens with -1043.22 (negative).
		//   myRounded1 := Round( myNumber, -1 )  ; Stores 1040 (correct).
		//   ChartModule := DllCall("LoadLibrary", "str", "rmchart.dll")
		//   myRounded2 := Round( myNumber, -1 )  ; Stores 1039 (wrong).
		_f_return_i((__int64)(value + (value > 0 ? 0.2 : -0.2)));
		// Formerly above simply returned (__int64)value.
}



BIF_DECL(BIF_FloorCeil)
// Probably saves little code size to merge extremely short/fast functions, hence FloorCeil.
// Floor() rounds down to the nearest integer; that is, to the integer that lies to the left on the
// number line (this is not the same as truncation because Floor(-1.2) is -2, not -1).
// Ceil() rounds up to the nearest integer; that is, to the integer that lies to the right on the number line.
//
// For simplicity and backward compatibility, a numeric result is always returned (even if the input
// is non-numeric or an empty string).
{
	Throw_if_Param_NaN(0);
	// The qmath routines are used because Floor() and Ceil() are deceptively difficult to implement in a way
	// that gives the correct result in all permutations of the following:
	// 1) Negative vs. positive input.
	// 2) Whether or not the input is already an integer.
	// Therefore, do not change this without conducting a thorough test.
	double x = ParamIndexToDouble(0);
	x = (_f_callee_id == FID_Floor) ? qmathFloor(x) : qmathCeil(x);
	// Fix for v1.0.40.05: For some inputs, qmathCeil/Floor yield a number slightly to the left of the target
	// integer, while for others they yield one slightly to the right.  For example, Ceil(62/61) and Floor(-4/3)
	// yield a double that would give an incorrect answer if it were simply truncated to an integer via
	// type casting.  The below seems to fix this without breaking the answers for other inputs (which is
	// surprisingly harder than it seemed).  There is a similar fix in BIF_Round().
	_f_return_i((__int64)(x + (x > 0 ? 0.2 : -0.2)));
}



BIF_DECL(BIF_Mod)
{
	// Load-time validation has already ensured there are exactly two parameters.
	// "Cast" each operand to Int64/Double depending on whether it has a decimal point.
	ExprTokenType param0, param1;
	if (ParamIndexToNumber(0, param0) && ParamIndexToNumber(1, param1)) // Both are numeric.
	{
		if (param0.symbol == SYM_INTEGER && param1.symbol == SYM_INTEGER) // Both are integers.
		{
			if (param1.value_int64 == 0)
				_f_throw(ERR_DIVIDEBYZERO, ErrorPrototype::ZeroDivision);
			// For performance, % is used vs. qmath for integers.
			_f_return_i(param0.value_int64 % param1.value_int64);
		}
		else // At least one is a floating point number.
		{
			double dividend = TokenToDouble(param0);
			double divisor = TokenToDouble(param1);
			if (divisor == 0.0)
				_f_throw(ERR_DIVIDEBYZERO, ErrorPrototype::ZeroDivision);
			_f_return(qmathFmod(dividend, divisor));
		}
	}
	// Since above didn't return, one or both parameters were invalid.
	_f_throw(ERR_PARAM_INVALID, ErrorPrototype::Type);
}



BIF_DECL(BIF_MinMax)
{
	// Supports one or more parameters.
	// Load-time validation has already ensured there is at least one parameter.
	ExprTokenType param;
	int index, ib_index = 0, db_index = 0;
	bool isMin = _f_callee_id == FID_Min;
	__int64 ia, ib = 0; double da, db = 0;
	bool ib_empty = TRUE, db_empty = TRUE;
	for (int i = 0; i < aParamCount; ++i)
	{
		ParamIndexToNumber(i, param);
		switch (param.symbol)
		{
			case SYM_INTEGER: // Compare only integers.
				ia = param.value_int64;
				if ((ib_empty) || (isMin ? ia < ib : ia > ib))
				{
					ib_empty = FALSE;
					ib = ia;
					ib_index = i;
				}
				break;
			case SYM_FLOAT: // Compare only floats.
				da = param.value_double;
				if ((db_empty) || (isMin ? da < db : da > db))
				{
					db_empty = FALSE;
					db = da;
					db_index = i;
				}
				break;
			default: // Non-operand or non-numeric string.
				_f_throw_param(i, _T("Number"));
		}
	}
	// Compare found integer with found float:
	index = (db_empty || !ib_empty && (isMin ? ib < db : ib > db)) ? ib_index : db_index;
	ParamIndexToNumber(index, param);
	aResultToken.symbol = param.symbol;
	aResultToken.value_int64 = param.value_int64;
}



BIF_DECL(BIF_Abs)
{
	if (!TokenToDoubleOrInt64(*aParam[0], aResultToken)) // "Cast" token to Int64/Double depending on whether it has a decimal point.
		_f_throw_param(0, _T("Number")); // Non-operand or non-numeric string.
	if (aResultToken.symbol == SYM_INTEGER)
	{
		// The following method is used instead of __abs64() to allow linking against the multi-threaded
		// DLLs (vs. libs) if that option is ever used (such as for a minimum size AutoHotkeySC.bin file).
		// It might be somewhat faster than __abs64() anyway, unless __abs64() is a macro or inline or something.
		if (aResultToken.value_int64 < 0)
			aResultToken.value_int64 = -aResultToken.value_int64;
	}
	else // Must be SYM_FLOAT due to the conversion above.
		aResultToken.value_double = qmathFabs(aResultToken.value_double);
}



BIF_DECL(BIF_Sin)
// For simplicity and backward compatibility, a numeric result is always returned (even if the input
// is non-numeric or an empty string).
{
	Throw_if_Param_NaN(0);
	_f_return(qmathSin(ParamIndexToDouble(0)));
}



BIF_DECL(BIF_Cos)
// For simplicity and backward compatibility, a numeric result is always returned (even if the input
// is non-numeric or an empty string).
{
	Throw_if_Param_NaN(0);
	_f_return(qmathCos(ParamIndexToDouble(0)));
}



BIF_DECL(BIF_Tan)
// For simplicity and backward compatibility, a numeric result is always returned (even if the input
// is non-numeric or an empty string).
{
	Throw_if_Param_NaN(0);
	_f_return(qmathTan(ParamIndexToDouble(0)));
}



BIF_DECL(BIF_ASinACos)
{
	Throw_if_Param_NaN(0);
	double value = ParamIndexToDouble(0);
	if (value > 1 || value < -1) // ASin and ACos aren't defined for such values.
	{
		_f_throw_param(0);
	}
	else
	{
		// For simplicity and backward compatibility, a numeric result is always returned in this case (even if
		// the input is non-numeric or an empty string).
		_f_return((_f_callee_id == FID_ASin) ? qmathAsin(value) : qmathAcos(value));
	}
}



BIF_DECL(BIF_ATan)
// For simplicity and backward compatibility, a numeric result is always returned (even if the input
// is non-numeric or an empty string).
{
	Throw_if_Param_NaN(0);
	_f_return(qmathAtan(ParamIndexToDouble(0)));
}



BIF_DECL(BIF_Exp)
{
	Throw_if_Param_NaN(0);
	_f_return(qmathExp(ParamIndexToDouble(0)));
}



BIF_DECL(BIF_SqrtLogLn)
{
	Throw_if_Param_NaN(0);
	double value = ParamIndexToDouble(0);
	if (value < 0) // Result is undefined in these cases.
	{
		_f_throw_param(0);
	}
	else
	{
		// For simplicity and backward compatibility, a numeric result is always returned in this case (even if
		// the input is non-numeric or an empty string).
		switch (_f_callee_id)
		{
		case FID_Sqrt:	_f_return(qmathSqrt(value));
		case FID_Log:	_f_return(qmathLog10(value));
		//case FID_Ln:
		default:		_f_return(qmathLog(value));
		}
	}
}



BIF_DECL(BIF_Random)
{
	UINT64 rand = 0;
	if (!GenRandom(&rand, sizeof(rand)))
		_f_throw(ERR_INTERNAL_CALL);

	SymbolType arg1type = ParamIndexIsOmitted(0) ? SYM_MISSING : ParamIndexIsNumeric(0);
	SymbolType arg2type = ParamIndexIsOmitted(1) ? SYM_MISSING : ParamIndexIsNumeric(1);
	if (arg1type == PURE_NOT_NUMERIC)
		_f_throw_param(0, _T("Number"));
	if (arg2type == PURE_NOT_NUMERIC)
		_f_throw_param(1, _T("Number"));

	bool use_float = arg1type == PURE_FLOAT || arg2type == PURE_FLOAT || !aParamCount; // Let Random() be Random(0.0, 1.0).
	if (use_float)
	{
		double target_min = arg1type != SYM_MISSING ? ParamIndexToDouble(0) : 0.0;
		double target_max = arg2type != SYM_MISSING ? ParamIndexToDouble(1) : arg1type == SYM_MISSING ? 1.0 : 0.0;
		// Be permissive about the order of parameters, and convert Random(n) to Random(0, n).
		if (target_min > target_max)
			swap(target_min, target_max);
		// The first part below produces a 53-bit integer, and from that a value between
		// 0.0 (inclusive) and 1.0 (exclusive) with the maximum precision for a double.
		_f_return((((rand >> 11) / 9007199254740992.0) * (target_max - target_min)) + target_min);
	}
	else
	{
		INT64 target_min = arg1type != SYM_MISSING ? ParamIndexToInt64(0) : 0;
		INT64 target_max = arg2type != SYM_MISSING ? ParamIndexToInt64(1) : 0;
		// Be permissive about the order of parameters, and convert Random(n) to Random(0, n).
		if (target_min > target_max)
			swap(target_min, target_max);
		// Do NOT use floating-point to generate random integers because of cases like
		// min=0 and max=1: we want an even distribution of 1's and 0's in that case, not
		// something skewed that might result due to rounding/truncation issues caused by
		// the float method used above.
		// Furthermore, the simple modulo approach is biased when the target range does not
		// divide cleanly into rand.  Suppose that rand ranges from 0..7, and the target range
		// is 0..2.  By using modulo, rand is effectively divided into sets {0..2, 3..5, 6..7}
		// and each value in the set is mapped to the target range 0..2.  Because the last set
		// maps to 0..1, 0 and 1 have greater chance of appearing than 2.  However, since rand
		// is 64-bit, this isn't actually a problem for small ranges such as 1..100 due to the
		// vanishingly small chance of rand falling within the defective set.
		UINT64 u_max = (UINT64)(target_max - target_min);
		if (u_max < UINT64_MAX)
		{
			// What we actually want is (UINT64_MAX + 1) % (u_max + 1), but without overflow.
			UINT64 error_margin = UINT64_MAX % (u_max + 1);
			if (error_margin != u_max) // i.e. ((error_margin + 1) % (u_max + 1)) != 0.
			{
				++error_margin;
				// error_margin is now the remainder after dividing (UINT64_MAX+1) by (u_max+1).
				// This is also the size of the incomplete number set which must be excluded to
				// ensure even distribution of random numbers.  For simplicity we just take the
				// number set starting at 0, since use of modulo below will auto-correct.
				// For example, with a target range of 1..100, the error_margin should be 16,
				// which gives a mere 16 in 2**64 chance that a second iteration will be needed.
				while (rand < error_margin)
					// To ensure even distribution, keep trying until outside the error margin.
					GenRandom(&rand, sizeof(rand));
			}
			rand %= (u_max + 1);
		}
		_f_return((INT64)(rand + (UINT64)target_min));
	}
}



BIF_DECL(BIF_DateAdd)
{
	FILETIME ft;
	if (!YYYYMMDDToFileTime(ParamIndexToString(0, _f_number_buf), ft))
		_f_throw_value(ERR_PARAM1_INVALID);

	// Use double to support a floating point value for days, hours, minutes, etc:
	double nUnits;
	nUnits = TokenToDouble(*aParam[1]);

	// Convert to 10ths of a microsecond (the units of the FILETIME struct):
	switch (ctoupper(*TokenToString(*aParam[2])))
	{
	case 'S': // Seconds
		nUnits *= (double)10000000;
		break;
	case 'M': // Minutes
		nUnits *= ((double)10000000 * 60);
		break;
	case 'H': // Hours
		nUnits *= ((double)10000000 * 60 * 60);
		break;
	case 'D': // Days
		nUnits *= ((double)10000000 * 60 * 60 * 24);
		break;
	default: // Invalid
		_f_throw_param(2);
	}
	// Convert ft struct to a 64-bit variable (maybe there's some way to avoid these conversions):
	ULARGE_INTEGER ul;
	ul.LowPart = ft.dwLowDateTime;
	ul.HighPart = ft.dwHighDateTime;
	// Add the specified amount of time to the result value:
	ul.QuadPart += (__int64)nUnits;  // Seems ok to cast/truncate in light of the *=10000000 above.
	// Convert back into ft struct:
	ft.dwLowDateTime = ul.LowPart;
	ft.dwHighDateTime = ul.HighPart;
	_f_return_p(FileTimeToYYYYMMDD(_f_retval_buf, ft, false));
}



BIF_DECL(BIF_DateDiff)
{
	LPTSTR error_message;
	// If either parameter is blank, it will default to the current time:
	__int64 time_until; // Declaring separate from initializing avoids compiler warning when not inside a block.
	TCHAR number_buf[MAX_NUMBER_SIZE]; // Additional buf used in case both parameters are pure numbers.
	time_until = YYYYMMDDSecondsUntil(
		ParamIndexToString(1, _f_number_buf),
		ParamIndexToString(0, number_buf),
		error_message);
	if (error_message) // Usually caused by an invalid component in the date-time string.
		_f_throw(error_message);
	switch (ctoupper(*ParamIndexToString(2)))
	{
	case 'S': break;
	case 'M': time_until /= 60; break; // Minutes
	case 'H': time_until /= 60 * 60; break; // Hours
	case 'D': time_until /= 60 * 60 * 24; break; // Days
	default: // Invalid
		_f_throw_param(2);
	}
	_f_return_i(time_until);
}



BIF_DECL(BIF_Hotkey)
{
	_f_param_string_opt(aParam0, 0);
	_f_param_string_opt(aParam1, 1);
	_f_param_string_opt(aParam2, 2);
	
	ResultType result = OK;
	IObject *functor = nullptr;

	switch (_f_callee_id) 
	{
	case FID_Hotkey:
	{
		HookActionType hook_action = 0;
		if (!ParamIndexIsOmitted(1))
		{
			if (  !(functor = ParamIndexToObject(1)) && *aParam1
				&& !(hook_action = Hotkey::ConvertAltTab(aParam1, true))  )
			{
				// Search for a match in the hotkey variants' "original callbacks".
				// I.e., find the function implicitly defined by "x::action".
				for (int i = 0; i < Hotkey::sHotkeyCount; ++i)
				{
					if (_tcscmp(Hotkey::shk[i]->mName, aParam1))
						continue;
					
					for (HotkeyVariant* v = Hotkey::shk[i]->mFirstVariant; v; v = v->mNextVariant)
						if (v->mHotCriterion == g->HotCriterion)
						{
							functor = v->mOriginalCallback.ToFunc();
							goto break_twice;
						}
				}
			break_twice:;
				if (!functor)
					_f_throw_param(1);
			}
			if (!functor)
				hook_action = Hotkey::ConvertAltTab(aParam1, true);
		}
		result = Hotkey::Dynamic(aParam0, aParam2, functor, hook_action, aResultToken);
		break;
	}
	case FID_HotIf:
		functor = ParamIndexToOptionalObject(0);
		result = Hotkey::IfExpr(aParam0, functor, aResultToken);
		break;
	
	default: // HotIfWinXXX
		result = SetHotkeyCriterion(_f_callee_id, aParam0, aParam1); // Currently, it only fails upon out-of-memory.
	
	}
	
	if (!result)
		_f_return_FAIL;
	_f_return_empty;
}



BIF_DECL(BIF_SetTimer)
{
	IObject *callback;
	// Note that only one timer per callback is allowed because the callback is the
	// unique identifier that allows us to update or delete an existing timer.
	if (ParamIndexIsOmitted(0)) // Fully omitted, not an empty string.
	{
		if (g->CurrentTimer)
			// Default to the timer which launched the current thread.
			callback = g->CurrentTimer->mCallback.ToObject();
		else
			callback = NULL;
		if (!callback)
			// Either the thread was not launched by a timer or the timer has been deleted.
			_f_throw_value(ERR_PARAM1_MUST_NOT_BE_BLANK);
	}
	else
	{
		callback = ParamIndexToObject(0);
		if (!callback)
			_f_throw_param(0, _T("object"));
		if (!ValidateFunctor(callback, 0, aResultToken))
			return;
	}
	__int64 period = DEFAULT_TIMER_PERIOD;
	int priority = 0;
	bool update_period = false, update_priority = false;
	if (!ParamIndexIsOmitted(1))
	{
		Throw_if_Param_NaN(1);
		period = ParamIndexToInt64(1);
		if (!period)
		{
			g_script.DeleteTimer(callback);
			_f_return_empty;
		}
		update_period = true;
	}
	if (!ParamIndexIsOmitted(2))
	{
		priority = ParamIndexToInt(2);
		update_priority = true;
	}
	g_script.UpdateOrCreateTimer(callback, update_period, period, update_priority, priority);
	_f_return_empty;
}


//////////////////////////////
// Event Handling Functions //
//////////////////////////////

BIF_DECL(BIF_OnMessage)
// Returns: An empty string.
// Parameters:
// 1: Message number to monitor.
// 2: Name of the function that will monitor the message.
// 3: Maximum threads and "register first" flag.
{
	// Currently OnMessage (in v2) has no return value.
	_f_set_retval_p(_T(""), 0);

	// Prior validation has ensured there's at least two parameters:
	UINT specified_msg = (UINT)ParamIndexToInt64(0); // Parameter #1

	// Set defaults:
	bool mode_is_delete = false;
	int max_instances = 1;
	bool call_it_last = true;

	if (!ParamIndexIsOmitted(2)) // Parameter #3 is present.
	{
		max_instances = (int)ParamIndexToInt64(2);
		// For backward-compatibility, values between MAX_INSTANCES+1 and SHORT_MAX must be supported.
		if (max_instances > MsgMonitorStruct::MAX_INSTANCES) // MAX_INSTANCES >= MAX_THREADS_LIMIT.
			max_instances = MsgMonitorStruct::MAX_INSTANCES;
		if (max_instances < 0) // MaxThreads < 0 is a signal to assign this monitor the lowest priority.
		{
			call_it_last = false; // Call it after any older monitors.  No effect if already registered.
			max_instances = -max_instances; // Convert to positive.
		}
		else if (max_instances == 0) // It would never be called, so this is used as a signal to delete the item.
			mode_is_delete = true;
	}

	// Parameter #2: The callback to add or remove.  Must be an object.
	IObject *callback = TokenToObject(*aParam[1]);
	if (!callback)
		_f_throw_param(1, _T("object"));

	// Check if this message already exists in the array:
	MsgMonitorStruct *pmonitor = g_MsgMonitor.Find(specified_msg, callback);
	bool item_already_exists = (pmonitor != NULL);
	if (!item_already_exists)
	{
		if (mode_is_delete) // Delete a non-existent item.
			_f_return_retval; // Yield the default return value set earlier (an empty string).
		if (!ValidateFunctor(callback, 4, aResultToken))
			return;
		// From this point on, it is certain that an item will be added to the array.
		pmonitor = g_MsgMonitor.Add(specified_msg, callback, call_it_last);
		if (!pmonitor)
			_f_throw_oom;
	}

	MsgMonitorStruct &monitor = *pmonitor;

	if (item_already_exists)
	{
		if (mode_is_delete)
		{
			// The msg-monitor is deleted from the array for two reasons:
			// 1) It improves performance because every incoming message for the app now needs to be compared
			//    to one less filter. If the count will now be zero, performance is improved even more because
			//    the overhead of the call to MsgMonitor() is completely avoided for every incoming message.
			// 2) It conserves space in the array in a situation where the script creates hundreds of
			//    msg-monitors and then later deletes them, then later creates hundreds of filters for entirely
			//    different message numbers.
			// The main disadvantage to deleting message filters from the array is that the deletion might
			// occur while the monitor is currently running, which requires more complex handling within
			// MsgMonitor() (see its comments for details).
			g_MsgMonitor.Delete(pmonitor);
			_f_return_retval;
		}
		if (aParamCount < 2) // Single-parameter mode: Report existing item's function name.
			_f_return_retval; // Everything was already set up above to yield the proper return value.
		// Otherwise, an existing item is being assigned a new function or MaxThreads limit.
		// Continue on to update this item's attributes.
	}
	else // This message was newly added to the array.
	{
		// The above already verified that callback is not NULL and there is room in the array.
		monitor.instance_count = 0; // Reset instance_count only for new items since existing items might currently be running.
		// Continue on to the update-or-create logic below.
	}

	// Update those struct attributes that get the same treatment regardless of whether this is an update or creation.
	if (!item_already_exists || !ParamIndexIsOmitted(2))
		monitor.max_instances = max_instances;
	// Otherwise, the parameter was omitted so leave max_instances at its current value.
	_f_return_retval;
}


MsgMonitorStruct *MsgMonitorList::Find(UINT aMsg, IObject *aCallback, UCHAR aMsgType)
{
	for (int i = 0; i < mCount; ++i)
		if (mMonitor[i].msg == aMsg
			&& mMonitor[i].func == aCallback // No need to check is_method, since it's impossible for an object and string to exist at the same address.
			&& mMonitor[i].msg_type == aMsgType) // Checked last because it's nearly always true.
			return mMonitor + i;
	return NULL;
}


MsgMonitorStruct *MsgMonitorList::Find(UINT aMsg, LPTSTR aMethodName, UCHAR aMsgType)
{
	for (int i = 0; i < mCount; ++i)
		if (mMonitor[i].msg == aMsg
			&& mMonitor[i].is_method && !_tcsicmp(aMethodName, mMonitor[i].method_name)
			&& mMonitor[i].msg_type == aMsgType) // Checked last because it's nearly always true.
			return mMonitor + i;
	return NULL;
}


MsgMonitorStruct *MsgMonitorList::AddInternal(UINT aMsg, bool aAppend)
{
	if (mCount == mCountMax)
	{
		int new_count = mCountMax ? mCountMax * mCountMax : 16;
		void *new_array = realloc(mMonitor, new_count * sizeof(MsgMonitorStruct));
		if (!new_array)
			return NULL;
		mMonitor = (MsgMonitorStruct *)new_array;
		mCountMax = new_count;
	}
	MsgMonitorStruct *new_mon;
	if (!aAppend)
	{
		for (MsgMonitorInstance *inst = mTop; inst; inst = inst->previous)
		{
			inst->index++; // Correct the index of each running monitor.
			inst->count++; // Iterate the same set of items which existed before.
			// By contrast, count isn't adjusted when adding at the end because we do not
			// want new items to be called by messages received before they were registered.
		}
		// Shift existing items to make room.
		memmove(mMonitor + 1, mMonitor, mCount * sizeof(MsgMonitorStruct));
		new_mon = mMonitor;
	}
	else
		new_mon = mMonitor + mCount;

	++mCount;
	new_mon->msg = aMsg;
	new_mon->msg_type = 0; // Must be initialised to 0 for all callers except GUI.
	// These are initialised by OnMessage, since OnExit and OnClipboardChange don't use them:
	//new_mon->instance_count = 0;
	//new_mon->max_instances = 1;
	return new_mon;
}


MsgMonitorStruct *MsgMonitorList::Add(UINT aMsg, IObject *aCallback, bool aAppend)
{
	MsgMonitorStruct *new_mon = AddInternal(aMsg, aAppend);
	if (new_mon)
	{
		aCallback->AddRef();
		new_mon->func = aCallback;
		new_mon->is_method = false;
	}
	return new_mon;
}


MsgMonitorStruct *MsgMonitorList::Add(UINT aMsg, LPTSTR aMethodName, bool aAppend)
{
	if (  !(aMethodName = _tcsdup(aMethodName))  )
		return NULL;
	MsgMonitorStruct *new_mon = AddInternal(aMsg, aAppend);
	if (new_mon)
	{
		new_mon->method_name = aMethodName;
		new_mon->is_method = true;
	}
	else
		free(aMethodName);
	return new_mon;
}


void MsgMonitorList::Delete(MsgMonitorStruct *aMonitor)
{
	ASSERT(aMonitor >= mMonitor && aMonitor < mMonitor + mCount);

	int mon_index = int(aMonitor - mMonitor);
	// Adjust the index of any active message monitors affected by this deletion.  This allows a
	// message monitor to delete older message monitors while still allowing any remaining monitors
	// of that message to be called (when there are multiple).
	for (MsgMonitorInstance *inst = mTop; inst; inst = inst->previous)
	{
		inst->Delete(mon_index);
	}
	// Remove the item from the array.
	--mCount;  // Must be done prior to the below.
	LPVOID release_me = aMonitor->union_value;
	bool is_method = aMonitor->is_method;
	if (mon_index < mCount) // An element other than the last is being removed. Shift the array to cover/delete it.
		memmove(aMonitor, aMonitor + 1, (mCount - mon_index) * sizeof(MsgMonitorStruct));
	if (is_method)
		free(release_me);
	else
		reinterpret_cast<IObject *>(release_me)->Release(); // Must be called last in case it calls a __delete() meta-function.
}


BOOL MsgMonitorList::IsMonitoring(UINT aMsg, UCHAR aMsgType)
{
	for (int i = 0; i < mCount; ++i)
		if (mMonitor[i].msg == aMsg && mMonitor[i].msg_type == aMsgType)
			return TRUE;
	return FALSE;
}


BOOL MsgMonitorList::IsRunning(UINT aMsg, UCHAR aMsgType)
// Returns true if there are any monitors for a message currently executing.
{
	for (MsgMonitorInstance *inst = mTop; inst; inst = inst->previous)
		if (!inst->deleted && mMonitor[inst->index].msg == aMsg && mMonitor[inst->index].msg_type == aMsgType)
			return TRUE;
	//if (!mTop)
	//	return FALSE;
	//for (int i = 0; i < mCount; ++i)
	//	if (mMonitor[i].msg == aMsg && mMonitor[i].instance_count)
	//		return TRUE;
	return FALSE;
}


void MsgMonitorList::Dispose()
{
	// Although other action taken by GuiType::Destroy() ensures the event list isn't
	// reachable from script once destruction begins, we take the careful approach and
	// decrement mCount at each iteration to ensure that if Release() executes external
	// code, this list is always in a valid state.
	while (mCount)
	{
		--mCount;
		if (mMonitor[mCount].is_method)
			free(mMonitor[mCount].method_name);
		else
			mMonitor[mCount].func->Release();
	}
	free(mMonitor);
	mMonitor = nullptr;
	mCountMax = 0;
	// Dispose all iterator instances to ensure Call() does not continue iterating:
	for (MsgMonitorInstance *inst = mTop; inst; inst = inst->previous)
		inst->Dispose();
}


BIF_DECL(BIF_On)
{
	_f_set_retval_p(_T("")); // In all cases there is no return value.
	auto event_type = _f_callee_id;
	MsgMonitorList *phandlers;
	switch (event_type)
	{
	case FID_OnError: phandlers = &g_script.mOnError; break;
	case FID_OnClipboardChange: phandlers = &g_script.mOnClipboardChange; break;
	default: phandlers = &g_script.mOnExit; break;
	}
	MsgMonitorList &handlers = *phandlers;


	IObject *callback = ParamIndexToObject(0);
	if (!callback)
		_f_throw_param(0, _T("object"));
	if (!ValidateFunctor(callback, event_type == FID_OnClipboardChange ? 1 : 2, aResultToken))
		return;
	
	int mode = 1; // Default.
	if (!ParamIndexIsOmitted(1))
		mode = ParamIndexToInt(1);

	MsgMonitorStruct *existing = handlers.Find(0, callback);

	switch (mode)
	{
	case  1:
	case -1:
		if (existing)
			return;
		if (event_type == FID_OnClipboardChange)
		{
			// Do this before adding the handler so that it won't be called as a result of the
			// SetClipboardViewer() call on Windows XP.  This won't cause existing handlers to
			// be called because in that case the clipboard listener is already enabled.
			g_script.EnableClipboardListener(true);
		}
		if (!handlers.Add(0, callback, mode == 1))
			_f_throw_oom;
		break;
	case  0:
		if (existing)
			handlers.Delete(existing);
		break;
	default:
		_f_throw_param(1);
	}
	// In case the above enabled the clipboard listener but failed to add the handler,
	// do this even if mode != 0:
	if (event_type == FID_OnClipboardChange && !handlers.Count())
		g_script.EnableClipboardListener(false);
}


///////////////////////
// Interop Functions //
///////////////////////


#ifdef ENABLE_REGISTERCALLBACK
struct RCCallbackFunc // Used by BIF_CallbackCreate() and related.
{
#ifdef WIN32_PLATFORM
	ULONG data1;	//E8 00 00 00
	ULONG data2;	//00 8D 44 24
	ULONG data3;	//08 50 FF 15
	UINT_PTR (CALLBACK **callfuncptr)(UINT_PTR*, char*);
	ULONG data4;	//59 84 C4 nn
	USHORT data5;	//FF E1
#endif
#ifdef _WIN64
	UINT64 data1; // 0xfffffffff9058d48
	UINT64 data2; // 0x9090900000000325
	void (*stub)();
	UINT_PTR (CALLBACK *callfuncptr)(UINT_PTR*, char*);
#endif
	//code ends
	UCHAR actual_param_count; // This is the actual (not formal) number of parameters passed from the caller to the callback. Kept adjacent to the USHORT above to conserve memory due to 4-byte struct alignment.
#define CBF_CREATE_NEW_THREAD	1
#define CBF_PASS_PARAMS_POINTER	2
	UCHAR flags; // Kept adjacent to above to conserve memory due to 4-byte struct alignment in 32-bit builds.
	IObject *func; // The function object to be called whenever the callback's caller calls callfuncptr.
};

#ifdef _WIN64
extern "C" void RegisterCallbackAsmStub();
#endif


UINT_PTR CALLBACK RegisterCallbackCStub(UINT_PTR *params, char *address) // Used by BIF_RegisterCallback().
// JGR: On Win32 parameters are always 4 bytes wide. The exceptions are functions which work on the FPU stack
// (not many of those). Win32 is quite picky about the stack always being 4 byte-aligned, (I have seen only one
// application which defied that and it was a patched ported DOS mixed mode application). The Win32 calling
// convention assumes that the parameter size equals the pointer size. 64 integers on Win32 are passed on
// pointers, or as two 32 bit halves for some functions...
{
#ifdef WIN32_PLATFORM
	RCCallbackFunc &cb = *((RCCallbackFunc*)(address-5)); //second instruction is 5 bytes after start (return address pushed by call)
#else
	RCCallbackFunc &cb = *((RCCallbackFunc*) address);
#endif

	BOOL pause_after_execute;

	// NOTES ABOUT INTERRUPTIONS / CRITICAL:
	// An incoming call to a callback is considered an "emergency" for the purpose of determining whether
	// critical/high-priority threads should be interrupted because there's no way easy way to buffer or
	// postpone the call.  Therefore, NO check of the following is done here:
	// - Current thread's priority (that's something of a deprecated feature anyway).
	// - Current thread's status of Critical (however, Critical would prevent us from ever being called in
	//   cases where the callback is triggered indirectly via message/dispatch due to message filtering
	//   and/or Critical's ability to pump messes less often).
	// - INTERRUPTIBLE_IN_EMERGENCY (which includes g_MenuIsVisible and g_AllowInterruption), which primarily
	//   affects SLEEP_WITHOUT_INTERRUPTION): It's debatable, but to maximize flexibility it seems best to allow
	//   callbacks during the display of a menu and during SLEEP_WITHOUT_INTERRUPTION.  For most callers of
	//   SLEEP_WITHOUT_INTERRUPTION, interruptions seem harmless.  For some it could be a problem, but when you
	//   consider how rare such callbacks are (mostly just subclassing of windows/controls) and what those
	//   callbacks tend to do, conflicts seem very rare.
	// Of course, a callback can also be triggered through explicit script action such as a DllCall of
	// EnumWindows, in which case the script would want to be interrupted unconditionally to make the call.
	// However, in those cases it's hard to imagine that INTERRUPTIBLE_IN_EMERGENCY wouldn't be true anyway.
	if (cb.flags & CBF_CREATE_NEW_THREAD)
	{
		if (g_nThreads >= g_MaxThreadsTotal) // To avoid array overflow, g_MaxThreadsTotal must not be exceeded except where otherwise documented.
			return 0;
		// See MsgSleep() for comments about the following section.
		InitNewThread(0, false, true);
		DEBUGGER_STACK_PUSH(_T("Callback"))
	}
	else
	{
		if (pause_after_execute = g->IsPaused) // Assign.
		{
			// v1.0.48: If the current thread is paused, this threadless callback would get stuck in
			// ExecUntil()'s pause loop (keep in mind that this situation happens only when a fast-mode
			// callback has been created without a script thread to control it, which goes against the
			// advice in the documentation). To avoid that, it seems best to temporarily unpause the
			// thread until the callback finishes.  But for performance, tray icon color isn't updated.
			g->IsPaused = false;
			--g_nPausedThreads; // See below.
			// If g_nPausedThreads isn't adjusted here, g_nPausedThreads could become corrupted if the
			// callback (or some thread that interrupts it) uses the Pause command/menu-item because
			// those aren't designed to deal with g->IsPaused being out-of-sync with g_nPausedThreads.
			// However, if --g_nPausedThreads reduces g_nPausedThreads to 0, timers would allowed to
			// run during the callback.  But that seems like the lesser evil, especially given that
			// this whole situation is very rare, and the documentation advises against doing it.
		}
		//else the current thread wasn't paused, which is usually the case.
		// TRAY ICON: g_script.UpdateTrayIcon() is not called because it's already in the right state
		// except when pause_after_execute==true, in which case it seems best not to change the icon
		// because it's likely to hurt any callback that's performance-sensitive.
	}

	g_script.mLastPeekTime = GetTickCount(); // Somewhat debatable, but might help minimize interruptions when the callback is called via message (e.g. subclassing a control; overriding a WindowProc).

	__int64 number_to_return;
	FuncResult result_token;
	ExprTokenType *param, one_param;
	int param_count;

	if (cb.flags & CBF_PASS_PARAMS_POINTER)
	{
		param_count = 1;
		param = &one_param;
		one_param.SetValue((UINT_PTR)params);
	}
	else
	{
		param_count = cb.actual_param_count;
		param = (ExprTokenType *)_alloca(param_count * sizeof(ExprTokenType));
		for (int i = 0; i < param_count; ++i)
			param[i].SetValue((UINT_PTR)params[i]);
	}
	
	CallMethod(cb.func, cb.func, nullptr, param, param_count, &number_to_return);
	// CallMethod()'s own return value is ignored because it wouldn't affect the handling below.
	
	if (cb.flags & CBF_CREATE_NEW_THREAD)
	{
		DEBUGGER_STACK_POP()
		ResumeUnderlyingThread();
	}
	else
	{
		if (g == g_array && !g_script.mAutoExecSectionIsRunning)
			// If the function just called used thread #0 and the AutoExec section isn't running, that means
			// the AutoExec section definitely didn't launch or control the callback (even if it is running,
			// it's not 100% certain it launched the callback). This can happen when a fast-mode callback has
			// been invoked via message, though the documentation advises against the fast mode when there is
			// no script thread controlling the callback.
			global_maximize_interruptibility(*g); // In case the script function called above used commands like Critical or "Thread Interrupt", ensure the idle thread is interruptible.  This avoids having to treat the idle thread as special in other places.
		//else never alter the interruptibility of AutoExec while it's running because it has its own method to do that.
		if (pause_after_execute) // See comments where it's defined.
		{
			g->IsPaused = true;
			++g_nPausedThreads;
		}
	}

	return (INT_PTR)number_to_return;
}



BIF_DECL(BIF_CallbackCreate)
// Returns: Address of callback procedure, or empty string on failure.
// Parameters:
// 1: Name of the function to be called when the callback routine is executed.
// 2: Options.
// 3: Number of parameters of callback.
//
// Author: Original x86 RegisterCallback() was created by Jonathan Rennison (JGR).
//   x64 support by fincs.  Various changes by Lexikos.
{
	IObject *func = ParamIndexToObject(0);
	if (!func)
		_f_throw_param(0, _T("object"));

	LPTSTR options = ParamIndexToOptionalString(1);
	bool pass_params_pointer = _tcschr(options, '&'); // Callback wants the address of the parameter list instead of their values.
#ifdef WIN32_PLATFORM
	bool use_cdecl = StrChrAny(options, _T("Cc")); // Recognize "C" as the "CDecl" option.
	bool require_param_count = !use_cdecl; // Param count must be specified for x86 stdcall.
#else
	bool require_param_count = false;
#endif

	bool params_specified = !ParamIndexIsOmittedOrEmpty(2);
	if (pass_params_pointer && require_param_count && !params_specified)
		_f_throw_value(ERR_PARAM3_MUST_NOT_BE_BLANK);

	int actual_param_count = params_specified ? ParamIndexToInt(2) : 0;
	if (!ValidateFunctor(func
		, pass_params_pointer ? 1 : actual_param_count // Count of script parameters being passed.
		, aResultToken
		// Use MinParams as actual_param_count if unspecified and no & option.
		, params_specified || pass_params_pointer ? nullptr : &actual_param_count))
	{
		return;
	}
	
#ifdef WIN32_PLATFORM
	if (!use_cdecl && actual_param_count > 31) // The ASM instruction currently used limits parameters to 31 (which should be plenty for any realistic use).
	{
		func->Release();
		_f_throw_param(2);
	}
#endif

	// GlobalAlloc() and dynamically-built code is the means by which a script can have an unlimited number of
	// distinct callbacks. On Win32, GlobalAlloc is the same function as LocalAlloc: they both point to
	// RtlAllocateHeap on the process heap. For large chunks of code you would reserve a 64K section with
	// VirtualAlloc and fill that, but for the 32 bytes we use here that would be overkill; GlobalAlloc is
	// much more efficient. MSDN says about GlobalAlloc: "All memory is created with execute access; no
	// special function is required to execute dynamically generated code. Memory allocated with this function
	// is guaranteed to be aligned on an 8-byte boundary." 
	// ABOVE IS OBSOLETE/INACCURATE: Systems with DEP enabled (and some without) require a VirtualProtect call
	// to allow the callback to execute.  MSDN currently says only this about the topic in the documentation
	// for GlobalAlloc:  "To execute dynamically generated code, use the VirtualAlloc function to allocate
	//						memory and the VirtualProtect function to grant PAGE_EXECUTE access."
	RCCallbackFunc *callbackfunc=(RCCallbackFunc*) GlobalAlloc(GMEM_FIXED,sizeof(RCCallbackFunc));	//allocate structure off process heap, automatically RWE and fixed.
	if (!callbackfunc)
		_f_throw_oom;
	RCCallbackFunc &cb = *callbackfunc; // For convenience and possible code-size reduction.

#ifdef WIN32_PLATFORM
	cb.data1=0xE8;       // call +0 -- E8 00 00 00 00 ;get eip, stays on stack as parameter 2 for C function (char *address).
	cb.data2=0x24448D00; // lea eax, [esp+8] -- 8D 44 24 08 ;eax points to params
	cb.data3=0x15FF5008; // push eax -- 50 ;eax pushed on stack as parameter 1 for C stub (UINT *params)
                         // call [xxxx] (in the lines below) -- FF 15 xx xx xx xx ;call C stub __stdcall, so stack cleaned up for us.

	// Comments about the static variable below: The reason for using the address of a pointer to a function,
	// is that the address is passed as a fixed address, whereas a direct call is passed as a 32-bit offset
	// relative to the beginning of the next instruction, which is more fiddly than it's worth to calculate
	// for dynamic code, as a relative call is designed to allow position independent calls to within the
	// same memory block without requiring dynamic fixups, or other such inconveniences.  In essence:
	//    call xxx ; is relative
	//    call [ptr_xxx] ; is position independent
	// Typically the latter is used when calling imported functions, etc., as only the pointers (import table),
	// need to be adjusted, not the calls themselves...

	static UINT_PTR (CALLBACK *funcaddrptr)(UINT_PTR*, char*) = RegisterCallbackCStub; // Use fixed absolute address of pointer to function, instead of varying relative offset to function.
	cb.callfuncptr = &funcaddrptr; // xxxx: Address of C stub.

	cb.data4=0xC48359 // pop ecx -- 59 ;return address... add esp, xx -- 83 C4 xx ;stack correct (add argument to add esp, nn for stack correction).
		+ (use_cdecl ? 0 : actual_param_count<<26);

	cb.data5=0xE1FF; // jmp ecx -- FF E1 ;return
#endif

#ifdef _WIN64
	/* Adapted from http://www.dyncall.org/
		lea rax, (rip)  # copy RIP (=p?) to RAX and use address in
		jmp [rax+16]    # 'entry' (stored at RIP+16) for jump
		nop
		nop
		nop
	*/
	cb.data1 = 0xfffffffff9058d48ULL;
	cb.data2 = 0x9090900000000325ULL;
	cb.stub = RegisterCallbackAsmStub;
	cb.callfuncptr = RegisterCallbackCStub;
#endif

	func->AddRef();
	cb.func = func;
	cb.actual_param_count = actual_param_count;
	cb.flags = 0;
	if (!StrChrAny(options, _T("Ff"))) // Recognize "F" as the "fast" mode that avoids creating a new thread.
		cb.flags |= CBF_CREATE_NEW_THREAD;
	if (pass_params_pointer)
		cb.flags |= CBF_PASS_PARAMS_POINTER;

	// If DEP is enabled (and sometimes when DEP is apparently "disabled"), we must change the
	// protection of the page of memory in which the callback resides to allow it to execute:
	DWORD dwOldProtect;
	VirtualProtect(callbackfunc, sizeof(RCCallbackFunc), PAGE_EXECUTE_READWRITE, &dwOldProtect);

	_f_return_i((__int64)callbackfunc); // Yield the callable address as the result.
}

BIF_DECL(BIF_CallbackFree)
{
	INT_PTR address = ParamIndexToIntPtr(0);
	if (address < 65536 && address >= 0) // Basic sanity check to catch incoming raw addresses that are zero or blank.  On Win32, the first 64KB of address space is always invalid.
		_f_throw_param(0);
	RCCallbackFunc *callbackfunc = (RCCallbackFunc *)address;
	callbackfunc->func->Release();
	callbackfunc->func = NULL; // To help detect bugs.
	GlobalFree(callbackfunc);
	_f_return_empty;
}

#endif



BIF_DECL(BIF_MenuFromHandle)
{
	auto *menu = g_script.FindMenu((HMENU)ParamIndexToInt64(0));
	if (menu)
	{
		menu->AddRef();
		_f_return(menu);
	}
	_f_return_empty;
}



//////////////////////
// Gui & GuiControl //
//////////////////////


void GuiControlType::StatusBar(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	GuiType& gui = *this->gui;
	HWND control_hwnd = this->hwnd;
	LPTSTR buf = _f_number_buf;

	HICON hicon;
	switch (aID)
	{
	case FID_SB_SetText:
		_o_return(SendMessage(control_hwnd, SB_SETTEXT
			, (WPARAM)((ParamIndexIsOmitted(1) ? 0 : ParamIndexToInt64(1) - 1) // The Part# param is present.
				     | (ParamIndexIsOmitted(2) ? 0 : ParamIndexToInt64(2) << 8)) // The uType parameter is present.
			, (LPARAM)ParamIndexToString(0, buf))); // Caller has ensured that there's at least one param in this mode.

	case FID_SB_SetParts:
		LRESULT old_part_count, new_part_count;
		int edge, part[256]; // Load-time validation has ensured aParamCount is under 255, so it shouldn't overflow.
		for (edge = 0, new_part_count = 0; new_part_count < aParamCount; ++new_part_count)
		{
			edge += gui.Scale(ParamIndexToInt(new_part_count)); // For code simplicity, no check for negative (seems fairly harmless since the bar will simply show up with the wrong number of parts to indicate the problem).
			part[new_part_count] = edge;
		}
		// For code simplicity, there is currently no means to have the last part of the bar use less than
		// all of the bar's remaining width.  The desire to do so seems rare, especially since the script can
		// add an extra/unused part at the end to achieve nearly (or perhaps exactly) the same effect.
		part[new_part_count++] = -1; // Make the last part use the remaining width of the bar.

		old_part_count = SendMessage(control_hwnd, SB_GETPARTS, 0, NULL); // MSDN: "This message always returns the number of parts in the status bar [regardless of how it is called]".
		if (old_part_count > new_part_count) // Some parts are being deleted, so destroy their icons.  See other notes in GuiType::Destroy() for explanation.
			for (LRESULT i = new_part_count; i < old_part_count; ++i) // Verified correct.
				if (hicon = (HICON)SendMessage(control_hwnd, SB_GETICON, i, 0))
					DestroyIcon(hicon);

		_o_return(SendMessage(control_hwnd, SB_SETPARTS, new_part_count, (LPARAM)part)
			? (__int64)control_hwnd : 0); // Return HWND to provide an easy means for the script to get the bar's HWND.

	//case FID_SB_SetIcon:
	default:
		int unused, icon_number;
		icon_number = ParamIndexToOptionalInt(1, 1);
		if (icon_number == 0) // Must be != 0 to tell LoadPicture that "icon must be loaded, never a bitmap".
			icon_number = 1;
		if (hicon = (HICON)LoadPicture(ParamIndexToString(0, buf) // Load-time validation has ensured there is at least one parameter.
			, GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON) // Apparently the bar won't scale them for us.
			, unused, icon_number, false)) // Defaulting to "false" for "use GDIplus" provides more consistent appearance across multiple OSes.
		{
			WPARAM part_index = ParamIndexIsOmitted(2) ? 0 : (WPARAM)ParamIndexToInt64(2) - 1;
			HICON hicon_old = (HICON)SendMessage(control_hwnd, SB_GETICON, part_index, 0); // Get the old one before setting the new one.
			// For code simplicity, the script is responsible for destroying the hicon later, if it ever destroys
			// the window.  Though in practice, most people probably won't do this, which is usually okay (if the
			// script doesn't load too many) since they're all destroyed by the system upon program termination.
			if (SendMessage(control_hwnd, SB_SETICON, part_index, (LPARAM)hicon))
			{
				if (hicon_old)
					// Although the old icon is automatically destroyed here, the script can call SendMessage(SB_SETICON)
					// itself if it wants to work with HICONs directly (for performance reasons, etc.)
					DestroyIcon(hicon_old);
			}
			else
			{
				DestroyIcon(hicon);
				hicon = NULL;
			}
		}
		//else can't load icon, so return 0.
		_o_return((size_t)hicon); // This allows the script to destroy the HICON later when it doesn't need it (see comments above too).
	// SB_SetTipText() not implemented (though can be done via SendMessage in the script) because the conditions
	// under which tooltips are displayed don't seem like something a script would want very often:
	// This ToolTip text is displayed in two situations: 
	// When the corresponding pane in the status bar contains only an icon. 
	// When the corresponding pane in the status bar contains text that is truncated due to the size of the pane.
	// In spite of the above, SB_SETTIPTEXT doesn't actually seem to do anything, even when the text is too long
	// to fit in a narrowed part, tooltip text has been set, and the user hovers the cursor over the bar.  Maybe
	// I'm not doing it right or maybe this feature is somehow disabled under certain service packs or conditions.
	//case 'T': // SB.SetTipText()
	//	break;
	} // switch(mode)
}


void GuiControlType::LV_GetNextOrCount(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
// LV.GetNext:
// Returns: The index of the found item, or 0 on failure.
// Parameters:
// 1: Starting index (one-based when it comes in).  If absent, search starts at the top.
// 2: Options string.
// 3: (FUTURE): Possible for use with LV.FindItem (though I think it can only search item text, not subitem text).
{
	HWND control_hwnd = hwnd;

	LPTSTR options;
	if (aID == FID_LV_GetCount)
	{
		options = (aParamCount > 0) ? omit_leading_whitespace(ParamIndexToString(0, _f_number_buf)) : _T("");
		if (*options)
		{
			if (ctoupper(*options) == 'S')
				_o_return(SendMessage(control_hwnd, LVM_GETSELECTEDCOUNT, 0, 0));
			else if (!_tcsnicmp(options, _T("Col"), 3)) // "Col" or "Column". Don't allow "C" by itself, so that "Checked" can be added in the future.
				_o_return(union_lv_attrib->col_count);
			else
				_o_throw_param(0);
		}
		_o_return(SendMessage(control_hwnd, LVM_GETITEMCOUNT, 0, 0));
	}
	// Since above didn't return, this is GetNext() mode.

	int index = -1;
	if (!ParamIndexIsOmitted(0))
	{
		if (!ParamIndexIsNumeric(0))
			_o_throw_param(0, _T("Number"));
		index = ParamIndexToInt(0) - 1; // -1 to convert to zero-based.
		// For flexibility, allow index to be less than -1 to avoid first-iteration complications in script loops
		// (such as when deleting rows, which shifts the row index upward, require the search to resume at
		// the previously found index rather than the row after it).  However, reset it to -1 to ensure
		// proper return values from the API in the "find checked item" mode used below.
		if (index < -1)
			index = -1;  // Signal it to start at the top.
	}

	// For performance, decided to always find next selected item when the "C" option hasn't been specified,
	// even when the checkboxes style is in effect.  Otherwise, would have to fetch and check checkbox style
	// bit for each call, which would slow down this heavily-called function.

	options = ParamIndexToOptionalString(1, _f_number_buf);
	TCHAR first_char = ctoupper(*omit_leading_whitespace(options));
	// To retain compatibility in the future, also allow "Check(ed)" and "Focus(ed)" since any word that
	// starts with C or F is already supported.

	switch(first_char)
	{
	case '\0': // Listed first for performance.
	case 'F':
		_o_return(ListView_GetNextItem(control_hwnd, index
			, first_char ? LVNI_FOCUSED : LVNI_SELECTED) + 1); // +1 to convert to 1-based.
	case 'C': // Checkbox: Find checked items. For performance assume that the control really has checkboxes.
	  {
		int item_count = ListView_GetItemCount(control_hwnd);
		for (int i = index + 1; i < item_count; ++i) // Start at index+1 to omit the first item from the search (for consistency with the other mode above).
			if (ListView_GetCheckState(control_hwnd, i)) // Item's box is checked.
				_o_return(i + 1); // +1 to convert from zero-based to one-based.
		// Since above didn't return, no match found.
		_o_return(0);
	  }
	default:
		_o_throw_param(1);
	}
}



void GuiControlType::LV_GetText(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
// Returns: Text on success.
// Throws on failure.
// Parameters:
// 1: Row index (one-based when it comes in).
// 2: Column index (one-based when it comes in).
{
	// Above sets default result in case of early return.  For code reduction, a zero is returned for all
	// the following conditions:
	// Item not found in ListView.
	// And others.

	int row_index = ParamIndexToInt(0) - 1; // -1 to convert to zero-based.
	if (row_index < -1) // row_index==-1 is reserved to mean "get column heading's text".
		_o_throw_param(0);
	// If parameter 2 is omitted, default to the first column (index 0):
	int col_index = ParamIndexIsOmitted(1) ? 0 : ParamIndexToInt(1) - 1; // -1 to convert to zero-based.
	if (col_index < 0)
		_o_throw_param(1);

	TCHAR buf[LV_TEXT_BUF_SIZE];

	if (row_index == -1) // Special mode to get column's text.
	{
		LVCOLUMN lvc;
		lvc.cchTextMax = LV_TEXT_BUF_SIZE - 1;  // See notes below about -1.
		lvc.pszText = buf;
		lvc.mask = LVCF_TEXT;
		if (SendMessage(hwnd, LVM_GETCOLUMN, col_index, (LPARAM)&lvc)) // Assign.
			_o_return(lvc.pszText); // See notes below about why pszText is used instead of buf (might apply to this too).
		else // On failure, it seems best to throw.
			_o_throw(ERR_FAILED);
	}
	else // Get row's indicated item or subitem text.
	{
		LVITEM lvi;
		// Subtract 1 because of that nagging doubt about size vs. length. Some MSDN examples subtract one, such as
		// TabCtrl_GetItem()'s cchTextMax:
		lvi.iItem = row_index;
		lvi.iSubItem = col_index; // Which field to fetch.  If it's zero, the item vs. subitem will be fetched.
		lvi.mask = LVIF_TEXT;
		lvi.pszText = buf;
		lvi.cchTextMax = LV_TEXT_BUF_SIZE - 1; // Note that LVM_GETITEM doesn't update this member to reflect the new length.
		// Unlike LVM_GETITEMTEXT, LVM_GETITEM indicates success or failure, which seems more useful/preferable
		// as a return value since a text length of zero would be ambiguous: could be an empty field or a failure.
		if (SendMessage(hwnd, LVM_GETITEM, 0, (LPARAM)&lvi)) // Assign
			// Must use lvi.pszText vs. buf because MSDN says: "Applications should not assume that the text will
			// necessarily be placed in the specified buffer. The control may instead change the pszText member
			// of the structure to point to the new text rather than place it in the buffer."
			_o_return(lvi.pszText);
		else // On failure, it seems best to throw.
			_o_throw(ERR_FAILED);
	}
}



void GuiControlType::LV_AddInsertModify(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
// Returns: 1 on success and 0 on failure.
// Parameters:
// 1: For Add(), this is the options.  For Insert/Modify, it's the row index (one-based when it comes in).
// 2: For Add(), this is the first field's text.  For Insert/Modify, it's the options.
// 3 and beyond: Additional field text.
// In Add/Insert mode, if there are no text fields present, a blank for is appended/inserted.
{
	GuiControlType &control = *this;
	auto mode = BuiltInFunctionID(aID);
	LPTSTR buf = _f_number_buf; // Resolve macro early for maintainability.

	int index;
	if (mode == FID_LV_Add) // For Add mode, use INT_MAX as a signal to append the item rather than inserting it.
	{
		index = INT_MAX;
		mode = FID_LV_Insert; // Add has now been set up to be the same as insert, so change the mode to simplify other things.
	}
	else // Insert or Modify: the target row-index is their first parameter, which load-time has ensured is present.
	{
		index = ParamIndexToInt(0) - 1; // -1 to convert to zero-based.
		if (index < -1 || (mode != FID_LV_Modify && index < 0)) // Allow -1 to mean "all rows" when in modify mode.
			_o_throw_param(0);
		++aParam;  // Remove the first parameter from further consideration to make Insert/Modify symmetric with Add.
		--aParamCount;
	}

	LPTSTR options = ParamIndexToOptionalString(0, buf);
	bool ensure_visible = false, is_checked = false;  // Checkmark.
	int col_start_index = 0;
	LVITEM lvi;
	lvi.mask = LVIF_STATE; // LVIF_STATE: state member is valid, but only to the extent that corresponding bits are set in stateMask (the rest will be ignored).
	lvi.stateMask = 0;
	lvi.state = 0;

	// Parse list of space-delimited options:
	TCHAR *next_option, *option_end, orig_char;
	bool adding; // Whether this option is being added (+) or removed (-).

	for (next_option = options; *next_option; next_option = omit_leading_whitespace(option_end))
	{
		if (*next_option == '-')
		{
			adding = false;
			// omit_leading_whitespace() is not called, which enforces the fact that the option word must
			// immediately follow the +/- sign.  This is done to allow the flexibility to have options
			// omit the plus/minus sign, and also to reserve more flexibility for future option formats.
			++next_option;  // Point it to the option word itself.
		}
		else
		{
			// Assume option is being added in the absence of either sign.  However, when we were
			// called by GuiControl(), the first option in the list must begin with +/- otherwise the cmd
			// would never have been properly detected as GUICONTROL_CMD_OPTIONS in the first place.
			adding = true;
			if (*next_option == '+')
				++next_option;  // Point it to the option word itself.
			//else do not increment, under the assumption that the plus has been omitted from a valid
			// option word and is thus an implicit plus.
		}

		if (!*next_option) // In case the entire option string ends in a naked + or -.
			break;
		// Find the end of this option item:
		if (   !(option_end = StrChrAny(next_option, _T(" \t")))   )  // Space or tab.
			option_end = next_option + _tcslen(next_option); // Set to position of zero terminator instead.
		if (option_end == next_option)
			continue; // i.e. the string contains a + or - with a space or tab after it, which is intentionally ignored.

		// Temporarily terminate to help eliminate ambiguity for words contained inside other words,
		// such as "Checked" inside of "CheckedGray":
		orig_char = *option_end;
		*option_end = '\0';

		if (!_tcsnicmp(next_option, _T("Select"), 6)) // Could further allow "ed" suffix by checking for that inside, but "Selected" is getting long so it doesn't seem something many would want to use.
		{
			next_option += 6;
			// If it's Select0, invert the mode to become "no select". This allows a boolean variable
			// to be more easily applied, such as this expression: "Select" . VarContainingState
			if (*next_option && !ATOI(next_option))
				adding = !adding;
			// Another reason for not having "Select" imply "Focus" by default is that it would probably
			// reduce performance when selecting all or a large number of rows.
			// Because a row might or might not have focus, the script may wish to retain its current
			// focused state.  For this reason, "select" does not imply "focus", which allows the
			// LVIS_FOCUSED bit to be omitted from the stateMask, which in turn retains the current
			// focus-state of the row rather than disrupting it.
			lvi.stateMask |= LVIS_SELECTED;
			if (adding)
				lvi.state |= LVIS_SELECTED;
			//else removing, so the presence of LVIS_SELECTED in the stateMask above will cause it to be de-selected.
		}
		else if (!_tcsnicmp(next_option, _T("Focus"), 5))
		{
			next_option += 5;
			if (*next_option && !ATOI(next_option)) // If it's Focus0, invert the mode to become "no focus".
				adding = !adding;
			lvi.stateMask |= LVIS_FOCUSED;
			if (adding)
				lvi.state |= LVIS_FOCUSED;
			//else removing, so the presence of LVIS_FOCUSED in the stateMask above will cause it to be de-focused.
		}
		else if (!_tcsnicmp(next_option, _T("Check"), 5))
		{
			// The rationale for not checking for an optional "ed" suffix here and incrementing next_option by 2
			// is that: 1) It would be inconsistent with the lack of support for "selected" (see reason above);
			// 2) Checkboxes in a ListView are fairly rarely used, so code size reduction might be more important.
			next_option += 5;
			if (*next_option && !ATOI(next_option)) // If it's Check0, invert the mode to become "unchecked".
				adding = !adding;
			if (mode == FID_LV_Modify) // v1.0.46.10: Do this section only for Modify, not Add/Insert, to avoid generating an extra "unchecked" notification when a row is added/inserted with an initial state of "checked".  In other words, the script now receives only a "checked" notification, not an "unchecked+checked". Search on is_checked for more comments.
			{
				lvi.stateMask |= LVIS_STATEIMAGEMASK;
				lvi.state |= adding ? 0x2000 : 0x1000; // The #1 image is "unchecked" and the #2 is "checked".
			}
			is_checked = adding;
		}
		else if (!_tcsnicmp(next_option, _T("Col"), 3))
		{
			if (adding)
			{
				col_start_index = ATOI(next_option + 3) - 1; // The ability to start at a column other than 1 (i.e. subitem vs. item).
				if (col_start_index < 0)
					col_start_index = 0;
			}
		}
		else if (!_tcsnicmp(next_option, _T("Icon"), 4))
		{
			// Testing shows that there is no way to avoid having an item icon in report view if the
			// ListView has an associated small-icon ImageList (well, perhaps you could have it show
			// a blank square by specifying an invalid icon index, but that doesn't seem useful).
			// If LVIF_IMAGE is entirely omitted when adding and item/row, the item will take on the
			// first icon in the list.  This is probably by design because the control wants to make
			// each item look consistent by indenting its first field by a certain amount for the icon.
			if (adding)
			{
				lvi.mask |= LVIF_IMAGE;
				lvi.iImage = ATOI(next_option + 4) - 1;  // -1 to convert to zero-based.
			}
			//else removal of icon currently not supported (see comment above), so do nothing in order
			// to reserve "-Icon" in case a future way can be found to do it.
		}
		else if (!_tcsicmp(next_option, _T("Vis"))) // v1.0.44
		{
			// Since this option much more typically used with LV.Modify than LV.Add/Insert, the technique of
			// Vis%VarContainingOneOrZero% isn't supported, to reduce code size.
			ensure_visible = adding; // Ignored by modes other than LV.Modify(), since it's not really appropriate when adding a row (plus would add code complexity).
		}
		else
		{
			aResultToken.ValueError(ERR_INVALID_OPTION, next_option);
			*option_end = orig_char; // See comment below.
			return;
		}

		*option_end = orig_char; // Undo the temporary termination because the caller needs aOptions to be unaltered.
	}

	// Suppress any events raised by the changes made below:
	control.attrib |= GUI_CONTROL_ATTRIB_SUPPRESS_EVENTS;

	// More maintainable and performs better to have a separate struct for subitems vs. items.
	LVITEM lvi_sub;
	// Ensure mask is pure to avoid giving it any excuse to fail due to the fact that
	// "You cannot set the state or lParam members for subitems."
	lvi_sub.mask = LVIF_TEXT;

	int i, j, rows_to_change;
	if (index == -1) // Modify all rows (above has ensured that this is only happens in modify-mode).
	{
		rows_to_change = ListView_GetItemCount(control.hwnd);
		lvi.iItem = 0;
		ensure_visible = false; // Not applicable when operating on all rows.
	}
	else // Modify or insert a single row.  Set it up for the loop to perform exactly one iteration.
	{
		rows_to_change = 1;
		lvi.iItem = index; // Which row to operate upon.  This can be a huge number such as 999999 if the caller wanted to append vs. insert.
	}
	lvi.iSubItem = 0;  // Always zero to operate upon the item vs. sub-item (subitems have their own LVITEM struct).
	int result = 1; // Set default from this point forward to be true/success. It will be overridden in insert mode to be the index of the new row.

	for (j = 0; j < rows_to_change; ++j, ++lvi.iItem) // ++lvi.iItem because if the loop has more than one iteration, by definition it is modifying all rows starting at 0.
	{
		if (!ParamIndexIsOmitted(1) && col_start_index == 0) // 2nd parameter: item's text (first field) is present, so include that when setting the item.
		{
			lvi.pszText = ParamIndexToString(1, buf); // Fairly low-overhead, so called every iteration for simplicity (so that buf can be used for both items and subitems).
			lvi.mask |= LVIF_TEXT;
		}
		if (mode == FID_LV_Insert) // Insert or Add.
		{
			// Note that ListView_InsertItem() will append vs. insert if the index is too large, in which case
			// it returns the items new index (which will be the last item in the list unless the control has
			// auto-sort style).
			if (   -1 == (lvi_sub.iItem = ListView_InsertItem(control.hwnd, &lvi))   )
			{
				control.attrib &= ~GUI_CONTROL_ATTRIB_SUPPRESS_EVENTS; // Re-enable events.
				_o_return(0); // Since item can't be inserted, no reason to try attaching any subitems to it.
			}
			// Update iItem with the actual index assigned to the item, which might be different than the
			// specified index if the control has an auto-sort style in effect.  This new iItem value
			// is used for ListView_SetCheckState() and for the attaching of any subitems to this item.
			result = lvi_sub.iItem + 1; // Convert from zero-based to 1-based.
			// For add/insert (but not modify), testing shows that checkmark must be added only after
			// the item has been inserted rather than provided in the lvi.state/stateMask fields.
			// MSDN confirms this by saying "When an item is added with [LVS_EX_CHECKBOXES],
			// it will always be set to the unchecked state [ignoring any value placed in bits
			// 12 through 15 of the state member]."
			if (is_checked)
				ListView_SetCheckState(control.hwnd, lvi_sub.iItem, TRUE); // TRUE = Check the row's checkbox.
				// Note that 95/NT4 systems that lack comctl32.dll 4.70+ distributed with MSIE 3.x
				// do not support LVS_EX_CHECKBOXES, so the above will have no effect for them.
		}
		else // Modify.
		{
			// Rather than trying to detect if anything was actually changed, this is called
			// unconditionally to simplify the code (ListView_SetItem() is probably very fast if it
			// discovers that lvi.mask==LVIF_STATE and lvi.stateMask==0).
			// By design (to help catch script bugs), a failure here does not revert to append mode.
			if (!ListView_SetItem(control.hwnd, &lvi)) // Returns TRUE/FALSE.
				result = 0; // Indicate partial failure, but attempt to continue in case it failed for reason other than "row doesn't exist".
			lvi_sub.iItem = lvi.iItem; // In preparation for modifying any subitems that need it.
			if (ensure_visible) // Seems best to do this one prior to "select" below.
				SendMessage(control.hwnd, LVM_ENSUREVISIBLE, lvi.iItem, FALSE); // PartialOK==FALSE is somewhat arbitrary.
		}

		// For each remaining parameter, assign its text to a subitem.
		// Testing shows that if the control has too few columns for all of the fields/parameters
		// present, the ones at the end are automatically ignored: they do not consume memory nor
		// do they significantly impact performance (at least on Windows XP).  For this reason, there
		// is no code above the for-loop above to reduce aParamCount if it's "too large" because
		// it might reduce flexibility (in case future/past OSes allow non-existent columns to be
		// populated, or in case current OSes allow the contents of recently removed columns to be modified).
		for (lvi_sub.iSubItem = (col_start_index > 1) ? col_start_index : 1 // Start at the first subitem unless we were told to start at or after the third column.
			// "i" starts at 2 (the third parameter) unless col_start_index is greater than 0, in which case
			// it starts at 1 (the second parameter) because that parameter has not yet been assigned to anything:
			, i = 2 - (col_start_index > 0)
			; i < aParamCount
			; ++i, ++lvi_sub.iSubItem)
		{
			if (aParam[i]->symbol == SYM_MISSING) // Omitted, such as LV.Modify(1,Opt,"One",,"Three").
				continue;
			lvi_sub.pszText = ParamIndexToString(i, buf); // Done every time through the outer loop since it's not high-overhead, and for code simplicity.
			if (!ListView_SetItem(control.hwnd, &lvi_sub) && mode != FID_LV_Insert) // Relies on short-circuit. Seems best to avoid loss of item's index in insert mode, since failure here should be rare.
				result = 0; // Indicate partial failure, but attempt to continue in case it failed for reason other than "row doesn't exist".
		}
	} // outer for()

	// When the control has no rows, work around the fact that LVM_SETITEMCOUNT delivers less than 20%
	// of its full benefit unless done after the first row is added (at least on XP SP1).  A non-zero
	// row_count_hint tells us that this message should be sent after the row has been inserted/appended:
	if (control.union_lv_attrib->row_count_hint > 0 && mode == FID_LV_Insert)
	{
		SendMessage(control.hwnd, LVM_SETITEMCOUNT, control.union_lv_attrib->row_count_hint, 0); // Last parameter should be 0 for LVS_OWNERDATA (verified if you look at the definition of ListView_SetItemCount macro).
		control.union_lv_attrib->row_count_hint = 0; // Reset so that it only gets set once per request.
	}

	control.attrib &= ~GUI_CONTROL_ATTRIB_SUPPRESS_EVENTS; // Re-enable events.
	_o_return(result);
}



void GuiControlType::LV_Delete(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
// Returns: 1 on success and 0 on failure.
// Parameters:
// 1: Row index (one-based when it comes in).
{
	if (ParamIndexIsOmitted(0))
		_o_return(SendMessage(hwnd, LVM_DELETEALLITEMS, 0, 0)); // Returns TRUE/FALSE.

	// Since above didn't return, there is a first parameter present.
	int index = ParamIndexToInt(0) - 1; // -1 to convert to zero-based.
	if (index > -1)
		_o_return(SendMessage(hwnd, LVM_DELETEITEM, index, 0)); // Returns TRUE/FALSE.
	else
		// Even if index==0, for safety, it seems best not to do a delete-all.
		_o_throw_param(0);
}



void GuiControlType::LV_InsertModifyDeleteCol(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
// Returns: 1 on success and 0 on failure.
// Parameters:
// 1: Column index (one-based when it comes in).
// 2: String of options
// 3: New text of column
// There are also some special modes when only zero or one parameter is present, see below.
{
	auto mode = BuiltInFunctionID(aID);
	LPTSTR buf = _f_number_buf; // Resolve macro early for maintainability.
	aResultToken.SetValue(0); // Set default return value.

	GuiControlType &control = *this;
	GuiType &gui = *control.gui;
	lv_attrib_type &lv_attrib = *control.union_lv_attrib;
	DWORD view_mode = mode != 'D' ? ListView_GetView(control.hwnd) : 0;

	int index;
	if (!ParamIndexIsOmitted(0))
		index = ParamIndexToInt(0) - 1; // -1 to convert to zero-based.
	else // Zero parameters.  Load-time validation has ensured that the 'D' (delete) mode cannot have zero params.
	{
		if (mode == FID_LV_ModifyCol)
		{
			if (view_mode != LV_VIEW_DETAILS)
				_o_return_retval; // Return 0 to indicate failure.
			// Otherwise:
			// v1.0.36.03: Don't attempt to auto-size the columns while the view is not report-view because
			// that causes any subsequent switch to the "list" view to be corrupted (invisible icons and items):
			for (int i = 0; ; ++i) // Don't limit it to lv_attrib.col_count in case script added extra columns via direct API calls.
				if (!ListView_SetColumnWidth(control.hwnd, i, LVSCW_AUTOSIZE)) // Failure means last column has already been processed.
					break;
			_o_return(1); // Always successful, regardless of what happened in the loop above.
		}
		// Since above didn't return, mode must be 'I' (insert).
		index = lv_attrib.col_count; // When no insertion index was specified, append to the end of the list.
	}

	// Do this prior to checking if index is in bounds so that it can support columns beyond LV_MAX_COLUMNS:
	if (mode == FID_LV_DeleteCol) // Delete a column.  In this mode, index parameter was made mandatory via load-time validation.
	{
		if (ListView_DeleteColumn(control.hwnd, index))  // Returns TRUE/FALSE.
		{
			// It's important to note that when the user slides columns around via drag and drop, the
			// column index as seen by the script is not changed.  This is fortunate because otherwise,
			// the lv_attrib.col array would get out of sync with the column indices.  Testing shows that
			// all of the following operations respect the original column index, regardless of where the
			// user may have moved the column physically: InsertCol, DeleteCol, ModifyCol.  Insert and Delete
			// shifts the indices of those columns that *originally* lay to the right of the affected column.
			if (lv_attrib.col_count > 0) // Avoid going negative, which would otherwise happen if script previously added columns by calling the API directly.
				--lv_attrib.col_count; // Must be done prior to the below.
			if (index < lv_attrib.col_count) // When a column other than the last was removed, adjust the array so that it stays in sync with actual columns.
				MoveMemory(lv_attrib.col+index, lv_attrib.col+index+1, sizeof(lv_col_type)*(lv_attrib.col_count-index));
			_f_set_retval_i(1);
		}
		_o_return_retval;
	}
	// Do this prior to checking if index is in bounds so that it can support columns beyond LV_MAX_COLUMNS:
	if (mode == FID_LV_ModifyCol && aParamCount < 2) // A single parameter is a special modify-mode to auto-size that column.
	{
		// v1.0.36.03: Don't attempt to auto-size the columns while the view is not report-view because
		// that causes any subsequent switch to the "list" view to be corrupted (invisible icons and items):
		if (view_mode == LV_VIEW_DETAILS)
			_f_set_retval_i(ListView_SetColumnWidth(control.hwnd, index, LVSCW_AUTOSIZE));
		//else leave retval set to 0.
		_o_return_retval;
	}
	if (mode == FID_LV_InsertCol)
	{
		if (lv_attrib.col_count >= LV_MAX_COLUMNS) // No room to insert or append.
			_o_return_retval;
		if (index >= lv_attrib.col_count) // For convenience, fall back to "append" when index too large.
			index = lv_attrib.col_count;
	}
	//else do nothing so that modification and deletion of columns that were added via script's
	// direct calls to the API can sort-of work (it's documented in the help file that it's not supported,
	// since col-attrib array can get out of sync with actual columns that way).

	if (index < 0 || index >= LV_MAX_COLUMNS) // For simplicity, do nothing else if index out of bounds.
		_o_return_retval; // Avoid array under/overflow below.

	// In addition to other reasons, must convert any numeric value to a string so that an isolated width is
	// recognized, e.g. LV.SetCol(1, old_width + 10):
	LPTSTR options = ParamIndexToOptionalString(1, buf);

	// It's done the following way so that when in insert-mode, if the column fails to be inserted, don't
	// have to remove the inserted array element from the lv_attrib.col array:
	lv_col_type temp_col = {0}; // Init unconditionally even though only needed for mode=='I'.
	lv_col_type &col = (mode == FID_LV_InsertCol) ? temp_col : lv_attrib.col[index]; // Done only after index has been confirmed in-bounds.

	LVCOLUMN lvc;
	lvc.mask = LVCF_FMT;
	if (mode == FID_LV_ModifyCol) // Fetch the current format so that it's possible to leave parts of it unaltered.
		ListView_GetColumn(control.hwnd, index, &lvc);
	else // Mode is "insert".
		lvc.fmt = 0;

	// Init defaults prior to parsing options:
	bool sort_now = false;
	int do_auto_size = (mode == FID_LV_InsertCol) ? LVSCW_AUTOSIZE_USEHEADER : 0;  // Default to auto-size for new columns.
	TCHAR sort_now_direction = 'A'; // Ascending.
	int new_justify = lvc.fmt & LVCFMT_JUSTIFYMASK; // Simplifies the handling of the justification bitfield.
	//lvc.iSubItem = 0; // Not necessary if the LVCF_SUBITEM mask-bit is absent.

	// Parse list of space-delimited options:
	TCHAR *next_option, *option_end, orig_char;
	bool adding; // Whether this option is being added (+) or removed (-).

	for (next_option = options; *next_option; next_option = omit_leading_whitespace(option_end))
	{
		if (*next_option == '-')
		{
			adding = false;
			// omit_leading_whitespace() is not called, which enforces the fact that the option word must
			// immediately follow the +/- sign.  This is done to allow the flexibility to have options
			// omit the plus/minus sign, and also to reserve more flexibility for future option formats.
			++next_option;  // Point it to the option word itself.
		}
		else
		{
			// Assume option is being added in the absence of either sign.  However, when we were
			// called by GuiControl(), the first option in the list must begin with +/- otherwise the cmd
			// would never have been properly detected as GUICONTROL_CMD_OPTIONS in the first place.
			adding = true;
			if (*next_option == '+')
				++next_option;  // Point it to the option word itself.
			//else do not increment, under the assumption that the plus has been omitted from a valid
			// option word and is thus an implicit plus.
		}

		if (!*next_option) // In case the entire option string ends in a naked + or -.
			break;
		// Find the end of this option item:
		if (   !(option_end = StrChrAny(next_option, _T(" \t")))   )  // Space or tab.
			option_end = next_option + _tcslen(next_option); // Set to position of zero terminator instead.
		if (option_end == next_option)
			continue; // i.e. the string contains a + or - with a space or tab after it, which is intentionally ignored.

		// Temporarily terminate to help eliminate ambiguity for words contained inside other words,
		// such as "Checked" inside of "CheckedGray":
		orig_char = *option_end;
		*option_end = '\0';

		// For simplicity, the value of "adding" is ignored for this and the other number/alignment options.
		if (!_tcsicmp(next_option, _T("Integer")))
		{
			// For simplicity, changing the col.type dynamically (since it's so rarely needed)
			// does not try to set up col.is_now_sorted_ascending so that the next click on the column
			// puts it into default starting order (which is ascending unless the Desc flag was originally
			// present).
			col.type = LV_COL_INTEGER;
			new_justify = LVCFMT_RIGHT;
		}
		else if (!_tcsicmp(next_option, _T("Float")))
		{
			col.type = LV_COL_FLOAT;
			new_justify = LVCFMT_RIGHT;
		}
		else if (!_tcsicmp(next_option, _T("Text"))) // Seems more approp. name than "Str" or "String"
			// Since "Text" is so general, it seems to leave existing alignment (Center/Right) as it is.
			col.type = LV_COL_TEXT;

		// The following can exist by themselves or in conjunction with the above.  They can also occur
		// *after* one of the above words so that alignment can be used to override the default for the type;
		// e.g. "Integer Left" to have left-aligned integers.
		else if (!_tcsicmp(next_option, _T("Right")))
			new_justify = adding ? LVCFMT_RIGHT : LVCFMT_LEFT;
		else if (!_tcsicmp(next_option, _T("Center")))
			new_justify = adding ? LVCFMT_CENTER : LVCFMT_LEFT;
		else if (!_tcsicmp(next_option, _T("Left"))) // Supported so that existing right/center column can be changed back to left.
			new_justify = LVCFMT_LEFT; // The value of "adding" seems inconsequential so is ignored.

		else if (!_tcsicmp(next_option, _T("Uni"))) // Unidirectional sort (clicking the column will not invert to the opposite direction).
			col.unidirectional = adding;
		else if (!_tcsicmp(next_option, _T("Desc"))) // Make descending order the default order (applies to uni and first click of col for non-uni).
			col.prefer_descending = adding; // So that the next click will toggle to the opposite direction.
		else if (!_tcsnicmp(next_option, _T("Case"), 4))
		{
			if (adding)
				col.case_sensitive = !_tcsicmp(next_option + 4, _T("Locale")) ? SCS_INSENSITIVE_LOCALE : SCS_SENSITIVE;
			else
				col.case_sensitive = SCS_INSENSITIVE;
		}
		else if (!_tcsicmp(next_option, _T("Logical"))) // v1.0.44.12: Supports StrCmpLogicalW() method of sorting.
			col.case_sensitive = SCS_INSENSITIVE_LOGICAL;

		else if (!_tcsnicmp(next_option, _T("Sort"), 4)) // This is done as an option vs. LV.SortCol/LV.Sort so that the column's options can be changed simultaneously with a "sort now" to refresh.
		{
			// Defer the sort until after all options have been parsed and applied.
			sort_now = true;
			if (!_tcsicmp(next_option + 4, _T("Desc")))
				sort_now_direction = 'D'; // Descending.
		}
		else if (!_tcsicmp(next_option, _T("NoSort"))) // Called "NoSort" so that there's a way to enable and disable the setting via +/-.
			col.sort_disabled = adding;

		else if (!_tcsnicmp(next_option, _T("Auto"), 4)) // No separate failure result is reported for this item.
			// In case the mode is "insert", defer auto-width of column until col exists.
			do_auto_size = _tcsicmp(next_option + 4, _T("Hdr")) ? LVSCW_AUTOSIZE : LVSCW_AUTOSIZE_USEHEADER;

		else if (!_tcsnicmp(next_option, _T("Icon"), 4))
		{
			next_option += 4;
			if (!_tcsicmp(next_option, _T("Right")))
			{
				if (adding)
					lvc.fmt |= LVCFMT_BITMAP_ON_RIGHT;
				else
					lvc.fmt &= ~LVCFMT_BITMAP_ON_RIGHT;
			}
			else // Assume its an icon number or the removal of the icon via -Icon.
			{
				if (adding)
				{
					lvc.mask |= LVCF_IMAGE;
					lvc.fmt |= LVCFMT_IMAGE; // Flag this column as displaying an image.
					lvc.iImage = ATOI(next_option) - 1; // -1 to convert to zero based.
				}
				else
					lvc.fmt &= ~LVCFMT_IMAGE; // Flag this column as NOT displaying an image.
			}
		}

		else // Handle things that are more general than the above, such as single letter options and pure numbers.
		{
			// Width does not have a W prefix to permit a naked expression to be used as the entirely of
			// options.  For example: LV.SetCol(1, old_width + 10)
			// v1.0.37: Fixed to allow floating point (although ATOI below will convert it to integer).
			if (IsNumeric(next_option, true, false, true)) // Above has already verified that *next_option can't be whitespace.
			{
				lvc.mask |= LVCF_WIDTH;
				int width = gui.Scale(ATOI(next_option));
				// Specifying a width when the column is initially added prevents the scrollbar from
				// updating on Windows 7 and 10 (but not XP).  As a workaround, initialise the width
				// to 0 and then resize it afterward.  do_auto_size is overloaded for this purpose
				// since it's already passed to ListView_SetColumnWidth().
				if (mode == 'I' && view_mode == LV_VIEW_DETAILS)
				{
					lvc.cx = 0; // Must be zero; if width is zero, ListView_SetColumnWidth() won't be called.
					do_auto_size = width; // If non-zero, this is passed to ListView_SetColumnWidth().
				}
				else
				{
					lvc.cx = width;
					do_auto_size = 0; // Turn off any auto-sizing that may have been put into effect (explicitly or by default).
				}
			}
			else
			{
				aResultToken.ValueError(ERR_INVALID_OPTION, next_option);
				*option_end = orig_char; // See comment below.
				return;
			}
		}

		// If the item was not handled by the above, ignore it because it is unknown.
		*option_end = orig_char; // Undo the temporary termination because the caller needs aOptions to be unaltered.
	}

	// Apply any changed justification/alignment to the fmt bit field:
	lvc.fmt = (lvc.fmt & ~LVCFMT_JUSTIFYMASK) | new_justify;

	if (!ParamIndexIsOmitted(2)) // Parameter #3 (text) is present.
	{
		lvc.pszText = ParamIndexToString(2, buf);
		lvc.mask |= LVCF_TEXT;
	}

	if (mode == FID_LV_ModifyCol) // Modify vs. Insert (Delete was already returned from, higher above).
		// For code simplicity, this is called unconditionally even if nothing internal the control's column
		// needs updating.  This seems justified given how rarely columns are modified.
		_f_set_retval_i(ListView_SetColumn(control.hwnd, index, &lvc)); // Returns TRUE/FALSE.
	else // Insert
	{
		// It's important to note that when the user slides columns around via drag and drop, the
		// column index as seen by the script is not changed.  This is fortunate because otherwise,
		// the lv_attrib.col array would get out of sync with the column indices.  Testing shows that
		// all of the following operations respect the original column index, regardless of where the
		// user may have moved the column physically: InsertCol, DeleteCol, ModifyCol.  Insert and Delete
		// shifts the indices of those columns that *originally* lay to the right of the affected column.
		// Doesn't seem to do anything -- not even with respect to inserting a new first column with it's
		// unusual behavior of inheriting the previously column's contents -- so it's disabled for now.
		// Testing shows that it also does not seem to cause a new column to inherit the indicated subitem's
		// text, even when iSubItem is set to index + 1 vs. index:
		//lvc.mask |= LVCF_SUBITEM;
		//lvc.iSubItem = index;
		// Testing shows that the following serve to set the column's physical/display position in the
		// heading to iOrder without affecting the specified index.  This concept is very similar to
		// when the user drags and drops a column heading to a new position: it's index doesn't change,
		// only it's displayed position:
		//lvc.mask |= LVCF_ORDER;
		//lvc.iOrder = index + 1;
		if (   -1 == (index = ListView_InsertColumn(control.hwnd, index, &lvc))   )
			_o_return_retval; // Since column could not be inserted, return so that below, sort-now, etc. are not done.
		_f_set_retval_i(index + 1); // +1 to convert the new index to 1-based.
		if (index < lv_attrib.col_count) // Since col is not being appended to the end, make room in the array to insert this column.
			MoveMemory(lv_attrib.col+index+1, lv_attrib.col+index, sizeof(lv_col_type)*(lv_attrib.col_count-index));
			// Above: Shift columns to the right by one.
		lv_attrib.col[index] = col; // Copy temp struct's members to the correct element in the array.
		// The above is done even when index==0 because "col" may contain attributes set via the Options
		// parameter.  Therefore, for code simplicity and rarity of real-world need, no attempt is made
		// to make the following idea work:
		// When index==0, retain the existing attributes due to the unique behavior of inserting a new first
		// column: The new first column inherit's the old column's values (fields), so it seems best to also have it
		// inherit the old column's attributes.
		++lv_attrib.col_count; // New column successfully added.  Must be done only after the MoveMemory() above.
	}

	// Auto-size is done only at this late a stage, in case column was just created above.
	// Note that ListView_SetColumn() apparently does not support LVSCW_AUTOSIZE_USEHEADER for it's "cx" member.
	// do_auto_size contains the actual column width if mode == 'I' and a width was passed by the caller.
	if (do_auto_size && view_mode == LV_VIEW_DETAILS)
		ListView_SetColumnWidth(control.hwnd, index, do_auto_size); // retval was previously set to the more important result above.
	//else v1.0.36.03: Don't attempt to auto-size the columns while the view is not report-view because
	// that causes any subsequent switch to the "list" view to be corrupted (invisible icons and items).

	if (sort_now)
		GuiType::LV_Sort(control, index, false, sort_now_direction);

	_o_return_retval;
}



void GuiControlType::LV_SetImageList(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
// Returns (MSDN): "handle to the image list previously associated with the control if successful; NULL otherwise."
// Parameters:
// 1: HIMAGELIST obtained from somewhere such as IL_Create().
// 2: Optional: Type of list.
{
	// Caller has ensured that there is at least one incoming parameter:
	HIMAGELIST himl = (HIMAGELIST)ParamIndexToInt64(0);
	int list_type;
	if (!ParamIndexIsOmitted(1))
		list_type = ParamIndexToInt(1);
	else // Auto-detect large vs. small icons based on the actual icon size in the image list.
	{
		int cx, cy;
		ImageList_GetIconSize(himl, &cx, &cy);
		list_type = (cx > GetSystemMetrics(SM_CXSMICON)) ? LVSIL_NORMAL : LVSIL_SMALL;
	}
	_o_return((size_t)ListView_SetImageList(hwnd, himl, list_type));
}



void GuiControlType::TV_AddModifyDelete(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
// TV.Add():
// Returns the HTREEITEM of the item on success, zero on failure.
// Parameters:
//    1: Text/name of item.
//    2: Parent of item.
//    3: Options.
// TV.Modify():
// Returns the HTREEITEM of the item on success (to allow nested calls in script, zero on failure or partial failure.
// Parameters:
//    1: ID of item to modify.
//    2: Options.
//    3: New name.
// Parameters for TV.Delete():
//    1: ID of item to delete (if omitted, all items are deleted).
{
	GuiControlType &control = *this;
	auto mode = BuiltInFunctionID(aID);
	LPTSTR buf = _f_number_buf; // Resolve macro early for maintainability.

	if (mode == FID_TV_Delete)
	{
		// If param #1 is present but is zero, for safety it seems best not to do a delete-all (in case a
		// script bug is so rare that it is never caught until the script is distributed).  Another reason
		// is that a script might do something like TV.Delete(TV.GetSelection()), which would be desired
		// to fail not delete-all if there's ever any way for there to be no selection.
		_o_return(SendMessage(control.hwnd, TVM_DELETEITEM, 0
			, ParamIndexIsOmitted(0) ? NULL : (LPARAM)ParamIndexToInt64(0)));
	}

	// Since above didn't return, this is TV.Add() or TV.Modify().
	TVINSERTSTRUCT tvi; // It contains a TVITEMEX, which is okay even if MSIE pre-4.0 on Win95/NT because those OSes will simply never access the new/bottommost item in the struct.
	bool add_mode = (mode == FID_TV_Add); // For readability & maint.
	HTREEITEM retval;
	
	// Suppress any events raised by the changes made below:
	control.attrib |= GUI_CONTROL_ATTRIB_SUPPRESS_EVENTS;

	LPTSTR options;
	if (add_mode) // TV.Add()
	{
		tvi.hParent = ParamIndexIsOmitted(1) ? NULL : (HTREEITEM)ParamIndexToInt64(1);
		tvi.hInsertAfter = TVI_LAST; // i.e. default is to insert the new item underneath the bottommost sibling.
		options = ParamIndexToOptionalString(2, buf);
		retval = 0; // Set default return value.
	}
	else // TV.Modify()
	{
		// NOTE: Must allow hitem==0 for TV.Modify, at least for the Sort option, because otherwise there would
		// be no way to sort the root-level items.
		tvi.item.hItem = (HTREEITEM)ParamIndexToInt64(0); // Load-time validation has ensured there is a first parameter for TV.Modify().
		// For modify-mode, set default return value to be "success" from this point forward.  Note that
		// in the case of sorting the root-level items, this will set it to zero, but since that almost
		// always succeeds and the script rarely cares whether it succeeds or not, adding code size for that
		// doesn't seem worth it:
		retval = tvi.item.hItem; // Set default return value.
		if (aParamCount < 2) // In one-parameter mode, simply select the item.
		{
			if (!TreeView_SelectItem(control.hwnd, tvi.item.hItem))
				retval = 0; // Override the HTREEITEM default value set above.
			control.attrib &= ~GUI_CONTROL_ATTRIB_SUPPRESS_EVENTS; // Re-enable events.
			_o_return((size_t)retval);
		}
		// Otherwise, there's a second parameter (even if it's 0 or "").
		options = ParamIndexToString(1, buf);
	}

	// Set defaults prior to options-parsing, to cover all omitted defaults:
	tvi.item.mask = TVIF_STATE; // TVIF_STATE: The state and stateMask members are valid (all other members are ignored).
	tvi.item.stateMask = 0; // All bits in "state" below are ignored unless the corresponding bit is present here in the mask.
	tvi.item.state = 0;
	// It seems tvi.item.cChildren is typically maintained by the control, though one exception is I_CHILDRENCALLBACK
	// and TVN_GETDISPINFO as mentioned at MSDN.

	DWORD select_flag = 0;
	bool ensure_visible = false, ensure_visible_first = false;

	// Parse list of space-delimited options:
	TCHAR *next_option, *option_end, orig_char;
	bool adding; // Whether this option is being added (+) or removed (-).

	for (next_option = options; *next_option; next_option = omit_leading_whitespace(option_end))
	{
		if (*next_option == '-')
		{
			adding = false;
			// omit_leading_whitespace() is not called, which enforces the fact that the option word must
			// immediately follow the +/- sign.  This is done to allow the flexibility to have options
			// omit the plus/minus sign, and also to reserve more flexibility for future option formats.
			++next_option;  // Point it to the option word itself.
		}
		else
		{
			// Assume option is being added in the absence of either sign.  However, when we were
			// called by GuiControl(), the first option in the list must begin with +/- otherwise the cmd
			// would never have been properly detected as GUICONTROL_CMD_OPTIONS in the first place.
			adding = true;
			if (*next_option == '+')
				++next_option;  // Point it to the option word itself.
			//else do not increment, under the assumption that the plus has been omitted from a valid
			// option word and is thus an implicit plus.
		}

		if (!*next_option) // In case the entire option string ends in a naked + or -.
			break;
		// Find the end of this option item:
		if (   !(option_end = StrChrAny(next_option, _T(" \t")))   )  // Space or tab.
			option_end = next_option + _tcslen(next_option); // Set to position of zero terminator instead.
		if (option_end == next_option)
			continue; // i.e. the string contains a + or - with a space or tab after it, which is intentionally ignored.

		// Temporarily terminate to help eliminate ambiguity for words contained inside other words,
		// such as "Checked" inside of "CheckedGray":
		orig_char = *option_end;
		*option_end = '\0';

		if (!_tcsicmp(next_option, _T("Select"))) // Could further allow "ed" suffix by checking for that inside, but "Selected" is getting long so it doesn't seem something many would want to use.
		{
			// Selection of an item apparently needs to be done via message for the control to update itself
			// properly.  Otherwise, single-select isn't enforced via de-selecting previous item and the newly
			// selected item isn't revealed/shown.  There may be other side-effects.
			if (adding)
				select_flag = TVGN_CARET;
			//else since "de-select" is not a supported action, no need to support "-Select".
			// Furthermore, since a TreeView is by its nature has only one item selected at a time, it seems
			// unnecessary to support Select%VarContainingOneOrZero%.  This is because it seems easier for a
			// script to simply load the Tree then select the desired item afterward.
		}
		else if (!_tcsnicmp(next_option, _T("Vis"), 3))
		{
			// Since this option much more typically used with TV.Modify than TV.Add, the technique of
			// Vis%VarContainingOneOrZero% isn't supported, to reduce code size.
			next_option += 3;
			if (!_tcsicmp(next_option, _T("First"))) // VisFirst
				ensure_visible_first = adding;
			else if (!*next_option)
				ensure_visible = adding;
		}
		else if (!_tcsnicmp(next_option, _T("Bold"), 4))
		{
			next_option += 4;
			if (*next_option && !ATOI(next_option)) // If it's Bold0, invert the mode.
				adding = !adding;
			tvi.item.stateMask |= TVIS_BOLD;
			if (adding)
				tvi.item.state |= TVIS_BOLD;
			//else removing, so the fact that this TVIS flag has just been added to the stateMask above
			// but is absent from item.state should remove this attribute from the item.
		}
		else if (!_tcsnicmp(next_option, _T("Expand"), 6))
		{
			next_option += 6;
			if (*next_option && !ATOI(next_option)) // If it's Expand0, invert the mode to become "collapse".
				adding = !adding;
			if (add_mode)
			{
				if (adding)
				{
					// Don't expand via msg because it won't work: since the item is being newly added
					// now, by definition it doesn't have any children, and testing shows that sending
					// the expand message has no effect, but setting the state bit does:
					tvi.item.stateMask |= TVIS_EXPANDED;
					tvi.item.state |= TVIS_EXPANDED;
					// Since the script is deliberately expanding the item, it seems best not to send the
					// TVN_ITEMEXPANDING/-ED messages because:
					// 1) Sending TVN_ITEMEXPANDED without first sending a TVN_ITEMEXPANDING message might
					//    decrease maintainability, and possibly even produce unwanted side-effects.
					// 2) Code size and performance (avoids generating extra message traffic).
				}
				//else removing, so nothing needs to be done because "collapsed" is the default state
				// of a TV item upon creation.
			}
			else // TV.Modify(): Expand and collapse both require a message to work properly on an existing item.
				// Strangely, this generates a notification sometimes (such as the first time) but not for subsequent
				// expands/collapses of that same item.  Also, TVE_TOGGLE is not currently supported because it seems
				// like it would be too rarely used.
				if (!TreeView_Expand(control.hwnd, tvi.item.hItem, adding ? TVE_EXPAND : TVE_COLLAPSE))
					retval = 0; // Indicate partial failure by overriding the HTREEITEM return value set earlier.
					// It seems that despite what MSDN says, failure is returned when collapsing and item that is
					// already collapsed, but not when expanding an item that is already expanded.  For performance
					// reasons and rarity of script caring, it seems best not to try to adjust/fix this.
		}
		else if (!_tcsnicmp(next_option, _T("Check"), 5))
		{
			// The rationale for not checking for an optional "ed" suffix here and incrementing next_option by 2
			// is that: 1) It would be inconsistent with the lack of support for "selected" (see reason above);
			// 2) Checkboxes in a ListView are fairly rarely used, so code size reduction might be more important.
			next_option += 5;
			if (*next_option && !ATOI(next_option)) // If it's Check0, invert the mode to become "unchecked".
				adding = !adding;
			//else removing, so the fact that this TVIS flag has just been added to the stateMask above
			// but is absent from item.state should remove this attribute from the item.
			tvi.item.stateMask |= TVIS_STATEIMAGEMASK;  // Unlike ListViews, Tree checkmarks can be applied in the same step as creating a Tree item.
			tvi.item.state |= adding ? 0x2000 : 0x1000; // The #1 image is "unchecked" and the #2 is "checked".
		}
		else if (!_tcsnicmp(next_option, _T("Icon"), 4))
		{
			if (adding)
			{
				// To me, having a different icon for when the item is selected seems rarely used.  After all,
				// its obvious the item is selected because it's highlighted (unless it lacks a name?)  So this
				// policy makes things easier for scripts that don't want to distinguish.  If ever it is needed,
				// new options such as IconSel and IconUnsel can be added.
				tvi.item.mask |= TVIF_IMAGE|TVIF_SELECTEDIMAGE;
				tvi.item.iSelectedImage = tvi.item.iImage = ATOI(next_option + 4) - 1;  // -1 to convert to zero-based.
			}
			//else removal of icon currently not supported (see comment above), so do nothing in order
			// to reserve "-Icon" in case a future way can be found to do it.
		}
		else if (!_tcsicmp(next_option, _T("Sort")))
		{
			if (add_mode)
				tvi.hInsertAfter = TVI_SORT; // For simplicity, the value of "adding" is ignored.
			else
				// Somewhat debatable, but it seems best to report failure via the return value even though
				// failure probably only occurs when the item has no children, and the script probably
				// doesn't often care about such failures.  It does result in the loss of the HTREEITEM return
				// value, but even if that call is nested in another, the zero should produce no effect in most cases.
				if (!TreeView_SortChildren(control.hwnd, tvi.item.hItem, FALSE)) // Best default seems no-recurse, since typically this is used after a user edits merely a single item.
					retval = 0; // Indicate partial failure by overriding the HTREEITEM return value set earlier.
		}
		// MUST BE LISTED LAST DUE TO "ELSE IF": Options valid only for TV.Add().
		else if (add_mode && !_tcsicmp(next_option, _T("First")))
		{
			tvi.hInsertAfter = TVI_FIRST; // For simplicity, the value of "adding" is ignored.
		}
		else if (add_mode && IsNumeric(next_option, false, false, false))
		{
			tvi.hInsertAfter = (HTREEITEM)ATOI64(next_option); // ATOI64 vs. ATOU avoids need for extra casting to avoid compiler warning.
		}
		else
		{
			control.attrib &= ~GUI_CONTROL_ATTRIB_SUPPRESS_EVENTS; // Re-enable events.
			aResultToken.ValueError(ERR_INVALID_OPTION, next_option);
			*option_end = orig_char; // See comment below.
			return;
		}

		// If the item was not handled by the above, ignore it because it is unknown.
		*option_end = orig_char; // Undo the temporary termination because the caller needs aOptions to be unaltered.
	}

	if (add_mode) // TV.Add()
	{
		tvi.item.pszText = ParamIndexToString(0, buf);
		tvi.item.mask |= TVIF_TEXT;
		tvi.item.hItem = TreeView_InsertItem(control.hwnd, &tvi); // Update tvi.item.hItem for convenience/maint. It's for use in later sections because retval is overridden to be zero for partial failure in modify-mode.
		retval = tvi.item.hItem; // Set return value.
	}
	else // TV.Modify()
	{
		if (!ParamIndexIsOmitted(2)) // An explicit empty string is allowed, which sets it to a blank value.  By contrast, if the param is omitted, the name is left changed.
		{
			tvi.item.pszText = ParamIndexToString(2, buf); // Reuse buf now that options (above) is done with it.
			tvi.item.mask |= TVIF_TEXT;
		}
		//else name/text parameter has been omitted, so don't change the item's name.
		if (tvi.item.mask != LVIF_STATE || tvi.item.stateMask) // An item's property or one of the state bits needs changing.
			if (!TreeView_SetItem(control.hwnd, &tvi.itemex))
				retval = 0; // Indicate partial failure by overriding the HTREEITEM return value set earlier.
	}

	if (ensure_visible) // Seems best to do this one prior to "select" below.
		SendMessage(control.hwnd, TVM_ENSUREVISIBLE, 0, (LPARAM)tvi.item.hItem); // Return value is ignored in this case, since its definition seems a little weird.
	if (ensure_visible_first) // Seems best to do this one prior to "select" below.
		TreeView_Select(control.hwnd, tvi.item.hItem, TVGN_FIRSTVISIBLE); // Return value is also ignored due to rarity, code size, and because most people wouldn't care about a failure even if for some reason it failed.
	if (select_flag)
		if (!TreeView_Select(control.hwnd, tvi.item.hItem, select_flag) && !add_mode) // Relies on short-circuit boolean order.
			retval = 0; // When not in add-mode, indicate partial failure by overriding the return value set earlier (add-mode should always return the new item's ID).

	control.attrib &= ~GUI_CONTROL_ATTRIB_SUPPRESS_EVENTS; // Re-enable events.
	_o_return((size_t)retval);
}



HTREEITEM GetNextTreeItem(HWND aTreeHwnd, HTREEITEM aItem)
// Helper function for others below.
// If aItem is NULL, caller wants topmost ROOT item returned.
// Otherwise, the next child, sibling, or parent's sibling is returned in a manner that allows the caller
// to traverse every item in the tree easily.
{
	if (!aItem)
		return TreeView_GetRoot(aTreeHwnd);
	// Otherwise, do depth-first recursion.  Must be done in the following order to allow full traversal:
	// Children first.
	// Then siblings.
	// Then parent's sibling(s).
	HTREEITEM hitem;
	if (hitem = TreeView_GetChild(aTreeHwnd, aItem))
		return hitem;
	if (hitem = TreeView_GetNextSibling(aTreeHwnd, aItem))
		return hitem;
	// The last stage is trickier than the above: parent's next sibling, or if none, its parent's parent's sibling, etc.
	for (HTREEITEM hparent = aItem;;)
	{
		if (   !(hparent = TreeView_GetParent(aTreeHwnd, hparent))   ) // No parent, so this is a root-level item.
			return NULL; // There is no next item.
		// Now it's known there is a parent.  It's not necessary to check that parent's children because that
		// would have been done by a prior iteration in the script.
		if (hitem = TreeView_GetNextSibling(aTreeHwnd, hparent))
			return hitem;
		// Otherwise, parent has no sibling, but does its parent (and so on)? Continue looping to find out.
	}
}



void GuiControlType::TV_GetRelatedItem(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
// TV.GetParent/Child/Selection/Next/Prev(hitem):
// The above all return the HTREEITEM (or 0 on failure).
// When TV.GetNext's second parameter is present, the search scope expands to include not just siblings,
// but also children and parents, which allows a tree to be traversed from top to bottom without the script
// having to do something fancy.
{
	GuiControlType &control = *this;
	HWND control_hwnd = control.hwnd;

	HTREEITEM hitem = (HTREEITEM)ParamIndexToOptionalIntPtr(0, NULL);

	if (ParamIndexIsOmitted(1))
	{
		WPARAM flag;
		switch (aID)
		{
		case FID_TV_GetSelection: flag = TVGN_CARET; break; // TVGN_CARET is the focused item.
		case FID_TV_GetParent: flag = TVGN_PARENT; break;
		case FID_TV_GetChild: flag = TVGN_CHILD; break;
		case FID_TV_GetPrev: flag = TVGN_PREVIOUS; break;
		case FID_TV_GetNext: flag = !hitem ? TVGN_ROOT : TVGN_NEXT; break; // TV_GetNext(no-parameters) yields very first item in Tree (TVGN_ROOT).
		// Above: It seems best to treat hitem==0 as "get root", even though it sacrifices some error detection,
		// because not doing so would be inconsistent with the fact that TV.GetNext(0, "Full") does get the root
		// (which needs to be retained to make script loops to traverse entire tree easier).
		//case FID_TV_GetCount:
		default:
			// There's a known bug mentioned at MSDN that a TreeView might report a negative count when there
			// are more than 32767 items in it (though of course HTREEITEM values are never negative since they're
			// defined as unsigned pseudo-addresses).  But apparently, that bug only applies to Visual Basic and/or
			// older OSes, because testing shows that SendMessage(TVM_GETCOUNT) returns 32800+ when there are more
			// than 32767 items in the tree, even without casting to unsigned.  So I'm not sure exactly what the
			// story is with this, so for now just casting to UINT rather than something less future-proof like WORD:
			// Older note, apparently unneeded at least on XP SP2: Cast to WORD to convert -1 through -32768 to the
			// positive counterparts.
			_o_return((UINT)SendMessage(control_hwnd, TVM_GETCOUNT, 0, 0));
		}
		// Apparently there's no direct call to get the topmost ancestor of an item, presumably because it's rarely
		// needed.  Therefore, no such mode is provide here yet (the syntax TV.GetParent(hitem, true) could be supported
		// if it's ever needed).
		_o_return(SendMessage(control_hwnd, TVM_GETNEXTITEM, flag, (LPARAM)hitem));
	}

	// Since above didn't return, this TV.GetNext's 2-parameter mode, which has an expanded scope that includes
	// not just siblings, but also children and parents.  This allows a tree to be traversed from top to bottom
	// without the script having to do something fancy.
	TCHAR first_char_upper = ctoupper(*omit_leading_whitespace(ParamIndexToString(1, _f_number_buf))); // Resolve parameter #2.
	bool search_checkmark;
	if (first_char_upper == 'C')
		search_checkmark = true;
	else if (first_char_upper == 'F')
		search_checkmark = false;
	else // Reserve other option letters/words for future use by being somewhat strict.
		_o_throw_param(1);

	// When an actual item was specified, search begins at the item *after* it.  Otherwise (when NULL):
	// It's a special mode that always considers the root node first.  Otherwise, there would be no way
	// to start the search at the very first item in the tree to find out whether it's checked or not.
	hitem = GetNextTreeItem(control_hwnd, hitem); // Handles the comment above.
	if (!search_checkmark) // Simple tree traversal, so just return the next item (if any).
		_o_return((size_t)hitem); // OK if NULL.

	// Otherwise, search for the next item having a checkmark. For performance, it seems best to assume that
	// the control has the checkbox style (the script would realistically never call it otherwise, so the
	// control's style isn't checked.
	for (; hitem; hitem = GetNextTreeItem(control_hwnd, hitem))
		if (TreeView_GetCheckState(control_hwnd, hitem) == 1) // 0 means unchecked, -1 means "no checkbox image".
			_o_return((size_t)hitem); // OK if NULL.
	// Since above didn't return, the entire tree starting at the specified item has been searched,
	// with no match found.
	_o_return(0);
}



void GuiControlType::TV_Get(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
// TV.Get()
// Returns: Varies depending on param #2.
// Parameters:
//    1: HTREEITEM.
//    2: Name of attribute to get.
// TV.GetText()
// Returns: Text on success.
// Throws on failure.
// Parameters:
//    1: HTREEITEM.
{
	bool get_text = aID == FID_TV_GetText;

	HWND control_hwnd = hwnd;

	if (!get_text)
	{
		// Loadtime validation has ensured that param #1 and #2 are present for all these cases.
		HTREEITEM hitem = (HTREEITEM)ParamIndexToInt64(0);
		UINT state_mask;
		switch (ctoupper(*omit_leading_whitespace(ParamIndexToString(1, _f_number_buf))))
		{
		case 'E': state_mask = TVIS_EXPANDED; break; // Expanded
		case 'C': state_mask = TVIS_STATEIMAGEMASK; break; // Checked
		case 'B': state_mask = TVIS_BOLD; break; // Bold
		//case 'S' for "Selected" is not provided because TV.GetSelection() seems to cover that well enough.
		//case 'P' for "is item a parent?" is not provided because TV.GetChild() seems to cover that well enough.
		// (though it's possible that retrieving TVITEM's cChildren would perform a little better).
		}
		// Below seems to be need a bit-AND with state_mask to work properly, at least on XP SP2.  Otherwise,
		// extra bits are present such as 0x2002 for "expanded" when it's supposed to be either 0x00 or 0x20.
		UINT result = state_mask & (UINT)SendMessage(control_hwnd, TVM_GETITEMSTATE, (WPARAM)hitem, state_mask);
		if (state_mask == TVIS_STATEIMAGEMASK)
		{
			if (result != 0x2000) // It doesn't have a checkmark state image.
				hitem = 0;
		}
		else // For all others, anything non-zero means the flag is present.
            if (!result) // Flag not present.
				hitem = 0;
		_o_return((size_t)hitem);
	}

	TCHAR text_buf[LV_TEXT_BUF_SIZE]; // i.e. uses same size as ListView.
	TVITEM tvi;
	tvi.hItem = (HTREEITEM)ParamIndexToInt64(0);
	tvi.mask = TVIF_TEXT;
	tvi.pszText = text_buf;
	tvi.cchTextMax = LV_TEXT_BUF_SIZE - 1; // -1 because of nagging doubt about size vs. length. Some MSDN examples subtract one), such as TabCtrl_GetItem()'s cchTextMax.

	if (SendMessage(control_hwnd, TVM_GETITEM, 0, (LPARAM)&tvi))
	{
		// Must use tvi.pszText vs. text_buf because MSDN says: "Applications should not assume that the text will
		// necessarily be placed in the specified buffer. The control may instead change the pszText member
		// of the structure to point to the new text rather than place it in the buffer."
		_o_return(tvi.pszText);
	}
	else
	{
		// On failure, it seems best to throw an exception.
		_o_throw(ERR_FAILED);
	}
}



void GuiControlType::TV_SetImageList(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
// Returns (MSDN): "handle to the image list previously associated with the control if successful; NULL otherwise."
// Parameters:
// 1: HIMAGELIST obtained from somewhere such as IL_Create().
// 2: Optional: Type of list.
{
	// Caller has ensured that there is at least one incoming parameter:
	HIMAGELIST himl = (HIMAGELIST)ParamIndexToInt64(0);
	int list_type;
	list_type = ParamIndexToOptionalInt(1, TVSIL_NORMAL);
	_o_return((size_t)TreeView_SetImageList(hwnd, himl, list_type));
}



BIF_DECL(BIF_IL_Create)
// Returns: Handle to the new image list, or 0 on failure.
// Parameters:
// 1: Initial image count (ImageList_Create() ignores values <=0, so no need for error checking).
// 2: Grow count (testing shows it can grow multiple times, even when this is set <=0, so it's apparently only a performance aid)
// 3: Width of each image (overloaded to mean small icon size when omitted or false, large icon size otherwise).
// 4: Future: Height of each image [if this param is present and >0, it would mean param 3 is not being used in its TRUE/FALSE mode)
// 5: Future: Flags/Color depth
{
	// The following old comment makes no sense because large icons are only used if param3 is NON-ZERO,
	// and there was never a distinction between passing zero and omitting the param:
	// So that param3 can be reserved as a future "specified width" param, to go along with "specified height"
	// after it, only when the parameter is both present and numerically zero are large icons used.  Otherwise,
	// small icons are used.
	int param3 = ParamIndexToOptionalBOOL(2, FALSE);
	_f_return_i((size_t)ImageList_Create(GetSystemMetrics(param3 ? SM_CXICON : SM_CXSMICON)
		, GetSystemMetrics(param3 ? SM_CYICON : SM_CYSMICON)
		, ILC_MASK | ILC_COLOR32  // ILC_COLOR32 or at least something higher than ILC_COLOR is necessary to support true-color icons.
		, ParamIndexToOptionalInt(0, 2)    // cInitial. 2 seems a better default than one, since it might be common to have only two icons in the list.
		, ParamIndexToOptionalInt(1, 5)));  // cGrow.  Somewhat arbitrary default.
}



BIF_DECL(BIF_IL_Destroy)
// Returns: 1 on success and 0 on failure.
// Parameters:
// 1: HIMAGELIST obtained from somewhere such as IL_Create().
{
	// Load-time validation has ensured there is at least one parameter.
	// Returns nonzero if successful, or zero otherwise, so force it to conform to TRUE/FALSE for
	// better consistency with other functions:
	_f_return_i(ImageList_Destroy((HIMAGELIST)ParamIndexToInt64(0)) ? 1 : 0);
}



BIF_DECL(BIF_IL_Add)
// Returns: the one-based index of the newly added icon, or zero on failure.
// Parameters:
// 1: HIMAGELIST: Handle of an existing ImageList.
// 2: Filename from which to load the icon or bitmap.
// 3: Icon number within the filename (or mask color for non-icon images).
// 4: The mere presence of this parameter indicates that param #3 is mask RGB-color vs. icon number.
//    This param's value should be "true" to resize the image to fit the image-list's size or false
//    to divide up the image into a series of separate images based on its width.
//    (this parameter could be overloaded to be the filename containing the mask image, or perhaps an HBITMAP
//    provided directly by the script)
// 5: Future: can be the scaling height to go along with an overload of #4 as the width.  However,
//    since all images in an image list are of the same size, the use of this would be limited to
//    only those times when the imagelist would be scaled prior to dividing it into separate images.
// The parameters above (at least #4) can be overloaded in the future calling ImageList_GetImageInfo() to determine
// whether the imagelist has a mask.
{
	HIMAGELIST himl = (HIMAGELIST)ParamIndexToInt64(0); // Load-time validation has ensured there is a first parameter.
	if (!himl)
		_f_throw_param(0);

	int param3 = ParamIndexToOptionalInt(2, 0);
	int icon_number, width = 0, height = 0; // Zero width/height causes image to be loaded at its actual width/height.
	if (!ParamIndexIsOmitted(3)) // Presence of fourth parameter switches mode to be "load a non-icon image".
	{
		icon_number = 0; // Zero means "load icon or bitmap (doesn't matter)".
		if (ParamIndexToBOOL(3)) // A value of True indicates that the image should be scaled to fit the imagelist's image size.
			ImageList_GetIconSize(himl, &width, &height); // Determine the width/height to which it should be scaled.
		//else retain defaults of zero for width/height, which loads the image at actual size, which in turn
		// lets ImageList_AddMasked() divide it up into separate images based on its width.
	}
	else
	{
		icon_number = param3; // LoadPicture() properly handles any wrong/negative value that might be here.
		ImageList_GetIconSize(himl, &width, &height); // L19: Determine the width/height of images in the image list to ensure icons are loaded at the correct size.
	}

	int image_type;
	HBITMAP hbitmap = LoadPicture(ParamIndexToString(1, _f_number_buf) // Caller has ensured there are at least two parameters.
		, width, height, image_type, icon_number, false); // Defaulting to "false" for "use GDIplus" provides more consistent appearance across multiple OSes.
	if (!hbitmap)
		_f_return_i(0);

	int index;
	if (image_type == IMAGE_BITMAP) // In this mode, param3 is always assumed to be an RGB color.
	{
		// Return the index of the new image or 0 on failure.
		index = ImageList_AddMasked(himl, hbitmap, rgb_to_bgr((int)param3)) + 1; // +1 to convert to one-based.
		DeleteObject(hbitmap);
	}
	else // ICON or CURSOR.
	{
		// Return the index of the new image or 0 on failure.
		index = ImageList_AddIcon(himl, (HICON)hbitmap) + 1; // +1 to convert to one-based.
		DestroyIcon((HICON)hbitmap); // Works on cursors too.  See notes in LoadPicture().
	}
	_f_return_i(index);
}



////////////////////
// Misc Functions //
////////////////////

BIF_DECL(BIF_LoadPicture)
{
	// h := LoadPicture(filename [, options, ByRef image_type])
	LPTSTR filename = ParamIndexToString(0, aResultToken.buf);
	LPTSTR options = ParamIndexToOptionalString(1);
	Var *image_type_var = ParamIndexToOutputVar(2);

	int width = -1;
	int height = -1;
	int icon_number = 0;
	bool use_gdi_plus = false;

	for (LPTSTR cp = options; cp; cp = StrChrAny(cp, _T(" \t")))
	{
		cp = omit_leading_whitespace(cp);
		if (tolower(*cp) == 'w')
			width = ATOI(cp + 1);
		else if (tolower(*cp) == 'h')
			height = ATOI(cp + 1);
		else if (!_tcsnicmp(cp, _T("Icon"), 4))
			icon_number = ATOI(cp + 4);
		else if (!_tcsnicmp(cp, _T("GDI+"), 4))
			// GDI+ or GDI+1 to enable, GDI+0 to disable.
			use_gdi_plus = cp[4] != '0';
	}

	if (width == -1 && height == -1)
		width = 0;

	int image_type;
	HBITMAP hbm = LoadPicture(filename, width, height, image_type, icon_number, use_gdi_plus);
	if (image_type_var)
		image_type_var->Assign(image_type);
	else if (image_type != IMAGE_BITMAP && hbm)
		// Always return a bitmap when the ImageType output var is omitted.
		hbm = IconToBitmap32((HICON)hbm, true); // Also works for cursors.
	aResultToken.value_int64 = (__int64)(UINT_PTR)hbm;
}



//////////////////////
// String Functions //
//////////////////////

BIF_DECL(BIF_Trim) // L31
{
	BuiltInFunctionID trim_type = _f_callee_id;

	LPTSTR buf = _f_retval_buf;
	size_t extract_length;
	LPTSTR str = ParamIndexToString(0, buf, &extract_length);

	LPTSTR result = str; // Prior validation has ensured at least 1 param.

	TCHAR omit_list_buf[MAX_NUMBER_SIZE]; // Support SYM_INTEGER/SYM_FLOAT even though it doesn't seem likely to happen.
	LPTSTR omit_list = ParamIndexIsOmitted(1) ? _T(" \t") : ParamIndexToString(1, omit_list_buf); // Default: space and tab.

	if (trim_type != FID_RTrim) // i.e. it's Trim() or LTrim()
	{
		result = omit_leading_any(result, omit_list, extract_length);
		extract_length -= (result - str); // Adjust for omitted characters.
	}
	if (extract_length && trim_type != FID_LTrim) // i.e. it's Trim() or RTrim();  THE LINE BELOW REQUIRES extract_length >= 1.
		extract_length = omit_trailing_any(result, omit_list, result + extract_length - 1);

	_f_return_p(result, extract_length);
}



BIF_DECL(BIF_VerCompare)
{
	_f_param_string(a, 0);
	_f_param_string(b, 1);
	_f_return(CompareVersion(a, b));
}



////////////////////
// Misc Functions //
////////////////////


BIF_DECL(BIF_Type)
{
	_f_return_p(TokenTypeString(*aParam[0]));
}

// Returns the type name of the given value.
LPTSTR TokenTypeString(ExprTokenType &aToken)
{
	switch (TypeOfToken(aToken))
	{
	case SYM_STRING: return STRING_TYPE_STRING;
	case SYM_INTEGER: return INTEGER_TYPE_STRING;
	case SYM_FLOAT: return FLOAT_TYPE_STRING;
	case SYM_OBJECT: return TokenToObject(aToken)->Type();
	default: return _T(""); // For maintainability.
	}
}



void Object::Error__New(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	LPTSTR message;
	TCHAR what_buf[MAX_NUMBER_SIZE], extra_buf[MAX_NUMBER_SIZE];
	LPCTSTR what = ParamIndexToOptionalString(1, what_buf);
	Line *line = g_script.mCurrLine;

	if (aID == M_OSError__New && (ParamIndexIsOmitted(0) || ParamIndexIsNumeric(0)))
	{
		DWORD error = ParamIndexIsOmitted(0) ? g->LastError : (DWORD)ParamIndexToInt64(0);
		SetOwnProp(_T("Number"), error);
		
		// Determine message based on error number.
		DWORD message_buf_size = _f_retval_buf_size;
		message = _f_retval_buf;
		DWORD size = (DWORD)_sntprintf(message, message_buf_size, (int)error < 0 ? _T("(0x%X) ") : _T("(%i) "), error);
		if (error) // Never show "Error: (0) The operation completed successfully."
			size += FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS, NULL, error, 0, message + size, message_buf_size - size, NULL);
		if (size)
		{
			if (message[size - 1] == '\n')
				message[--size] = '\0';
			if (message[size - 1] == '\r')
				message[--size] = '\0';
		}
	}
	else
		message = ParamIndexIsOmitted(0) ? Type() : ParamIndexToString(0, _f_number_buf);

#ifndef CONFIG_DEBUGGER
	if (ParamIndexIsOmitted(1) && g->CurrentFunc)
		what = g->CurrentFunc->mName;
	SetOwnProp(_T("Stack"), _T("")); // Avoid "unknown property" in compiled scripts.
#else
	DbgStack::Entry *stack_top = g_Debugger.mStack.mTop - 1;
	if (stack_top->type == DbgStack::SE_BIF && _tcsicmp(what, stack_top->func->mName))
		--stack_top;

	if (ParamIndexIsOmitted(1)) // "What"
	{
		if (g->CurrentFunc)
			what = g->CurrentFunc->mName;
	}
	else
	{
		int offset = ParamIndexIsNumeric(1) ? ParamIndexToInt(1) : 0;
		for (auto se = stack_top; se >= g_Debugger.mStack.mBottom; --se)
		{
			if (++offset == 0 || *what && !_tcsicmp(se->Name(), what))
			{
				line = se > g_Debugger.mStack.mBottom ? se[-1].line : se->line;
				// se->line contains the line at the given offset from the top of the stack.
				// Rather than returning the name of the function or sub which contains that
				// line, return the name of the function or sub which that line called.
				// In other words, an offset of -1 gives the name of the current function and
				// the file and number of the line which it was called from.
				what = se->Name();
				stack_top = se;
				break;
			}
			if (se->type == DbgStack::SE_Thread)
				break; // Look only within the current thread.
		}
	}

	TCHAR stack_buf[2048];
	GetScriptStack(stack_buf, _countof(stack_buf), stack_top);
	SetOwnProp(_T("Stack"), stack_buf);
#endif

	LPTSTR extra = ParamIndexToOptionalString(2, extra_buf);

	SetOwnProp(_T("Message"), message);
	SetOwnProp(_T("What"), const_cast<LPTSTR>(what));
	SetOwnProp(_T("File"), Line::sSourceFile[line->mFileIndex]);
	SetOwnProp(_T("Line"), line->mLineNumber);
	SetOwnProp(_T("Extra"), extra);
}



////////////////////////////////////////////////////////
// HELPER FUNCTIONS FOR TOKENS AND BUILT-IN FUNCTIONS //
////////////////////////////////////////////////////////

BOOL ResultToBOOL(LPTSTR aResult)
{
	UINT c1 = (UINT)*aResult; // UINT vs. UCHAR might squeeze a little more performance out of it.
	if (c1 > 48)     // Any UCHAR greater than '0' can't be a space(32), tab(9), '+'(43), or '-'(45), or '.'(46)...
		return TRUE; // ...so any string starting with c1>48 can't be anything that's false (e.g. " 0", "+0", "-0", ".0", "0").
	if (!c1 // Besides performance, must be checked anyway because otherwise IsNumeric() would consider "" to be non-numeric and thus TRUE.
		|| c1 == 48 && !aResult[1]) // The string "0" seems common enough to detect explicitly, for performance.
		return FALSE;
	// IsNumeric() is called below because there are many variants of a false string:
	// e.g. "0", "0.0", "0x0", ".0", "+0", "-0", and " 0" (leading spaces/tabs).
	switch (IsNumeric(aResult, true, false, true)) // It's purely numeric and not all whitespace (and due to earlier check, it's not blank).
	{
	case PURE_INTEGER: return ATOI64(aResult) != 0; // Could call ATOF() for both integers and floats; but ATOI64() probably performs better, and a float comparison to 0.0 might be a slower than an integer comparison to 0.
	case PURE_FLOAT:   return _tstof(aResult) != 0.0; // atof() vs. ATOF() because PURE_FLOAT is never hexadecimal.
	default: // PURE_NOT_NUMERIC.
		// Even a string containing all whitespace would be considered non-numeric since it's a non-blank string
		// that isn't equal to 0.
		return TRUE;
	}
}



BOOL VarToBOOL(Var &aVar)
{
	if (!aVar.HasContents()) // Must be checked first because otherwise IsNumeric() would consider "" to be non-numeric and thus TRUE.  For performance, it also exploits the binary number cache.
		return FALSE;
	switch (aVar.IsNumeric())
	{
	case PURE_INTEGER:
		return aVar.ToInt64() != 0;
	case PURE_FLOAT:
		return aVar.ToDouble() != 0.0;
	default:
		// Even a string containing all whitespace would be considered non-numeric since it's a non-blank string
		// that isn't equal to 0.
		return TRUE;
	}
}



BOOL TokenToBOOL(ExprTokenType &aToken)
{
	switch (aToken.symbol)
	{
	case SYM_INTEGER: // Probably the most common; e.g. both sides of "if (x>3 and x<6)" are the number 1/0.
		return aToken.value_int64 != 0; // Force it to be purely 1 or 0.
	case SYM_VAR:
		return VarToBOOL(*aToken.var);
	case SYM_FLOAT:
		return aToken.value_double != 0.0;
	case SYM_STRING:
		return ResultToBOOL(aToken.marker);
	default:
		// The only remaining valid symbol is SYM_OBJECT, which is always TRUE.
		return TRUE;
	}
}



ToggleValueType TokenToToggleValue(ExprTokenType &aToken)
{
	if (TokenIsNumeric(aToken))
	switch (TokenToInt64(aToken)) // Reserve other values for potential future use by requiring exact match.
	{
	case 1: return TOGGLED_ON;
	case 0: return TOGGLED_OFF;
	case -1: return TOGGLE;
	}
	return TOGGLE_INVALID;
}



SymbolType TokenIsNumeric(ExprTokenType &aToken)
{
	switch(aToken.symbol)
	{
	case SYM_INTEGER:
	case SYM_FLOAT:
		return aToken.symbol;
	case SYM_VAR: 
		return aToken.var->IsNumeric();
	default: // SYM_STRING: Callers of this function expect a "numeric" result for numeric strings.
		return IsNumeric(aToken.marker, true, false, true);
	}
}


SymbolType TokenIsPureNumeric(ExprTokenType &aToken)
{
	switch (aToken.symbol)
	{
	case SYM_INTEGER:
	case SYM_FLOAT:
		return aToken.symbol;
	case SYM_VAR:
		if (!aToken.var->IsUninitializedNormalVar()) // Caller doesn't want a warning, so avoid calling Contents().
			return aToken.var->IsPureNumeric();
		//else fall through:
	default:
		return PURE_NOT_NUMERIC;
	}
}


SymbolType TokenIsPureNumeric(ExprTokenType &aToken, SymbolType &aNumType)
// This function is called very frequently by ExpandExpression(), which needs to distinguish
// between numeric strings and pure numbers, but still needs to know if the string is numeric.
{
	switch(aToken.symbol)
	{
	case SYM_INTEGER:
	case SYM_FLOAT:
		return aNumType = aToken.symbol;
	case SYM_VAR:
		if (aToken.var->IsUninitializedNormalVar()) // Caller doesn't want a warning, so avoid calling Contents().
			return aNumType = PURE_NOT_NUMERIC; // i.e. empty string is non-numeric.
		if (aNumType = aToken.var->IsPureNumeric())
			return aNumType; // This var contains a pure binary number.
		// Otherwise, it might be a numeric string (i.e. impure).
		aNumType = aToken.var->IsNumeric(); // This also caches the PURE_NOT_NUMERIC result if applicable.
		return PURE_NOT_NUMERIC;
	case SYM_STRING:
		aNumType = IsNumeric(aToken.marker, true, false, true);
		return PURE_NOT_NUMERIC;
	default:
		// Only SYM_OBJECT and SYM_MISSING should be possible.
		return aNumType = PURE_NOT_NUMERIC;
	}
}


BOOL TokenIsEmptyString(ExprTokenType &aToken)
{
	switch (aToken.symbol)
	{
	case SYM_STRING:
		return !*aToken.marker;
	case SYM_VAR:
		return !aToken.var->HasContents();
	//case SYM_MISSING: // This case is omitted because it currently should be
		// impossible for all callers except for ParamIndexIsOmittedOrEmpty(),
		// which checks for it explicitly.
		//return TRUE;
	default:
		return FALSE;
	}
}


__int64 TokenToInt64(ExprTokenType &aToken)
// Caller has ensured that any SYM_VAR's Type() is VAR_NORMAL.
// Converts the contents of aToken to a 64-bit int.
{
	// Some callers, such as those that cast our return value to UINT, rely on the use of 64-bit
	// to preserve unsigned values and also wrap any signed values into the unsigned domain.
	switch (aToken.symbol)
	{
		case SYM_INTEGER:	return aToken.value_int64;
		case SYM_FLOAT:		return (__int64)aToken.value_double;
		case SYM_VAR:		return aToken.var->ToInt64();
		case SYM_STRING:	return ATOI64(aToken.marker);
	}
	// Since above didn't return, it can only be SYM_OBJECT or not an operand.
	return 0;
}



double TokenToDouble(ExprTokenType &aToken, BOOL aCheckForHex)
// Caller has ensured that any SYM_VAR's Type() is VAR_NORMAL.
// Converts the contents of aToken to a double.
{
	switch (aToken.symbol)
	{
		case SYM_FLOAT:		return aToken.value_double;
		case SYM_INTEGER:	return (double)aToken.value_int64;
		case SYM_VAR:		return aToken.var->ToDouble();
		case SYM_STRING:	return aCheckForHex ? ATOF(aToken.marker) : _tstof(aToken.marker); // atof() omits the check for hexadecimal.
	}
	// Since above didn't return, it can only be SYM_OBJECT or not an operand.
	return 0;
}



LPTSTR TokenToString(ExprTokenType &aToken, LPTSTR aBuf, size_t *aLength)
// Returns "" on failure to simplify logic in callers.  Otherwise, it returns either aBuf (if aBuf was needed
// for the conversion) or the token's own string.  aBuf may be NULL, in which case the caller presumably knows
// that this token is SYM_STRING or SYM_VAR (or the caller wants "" back for anything other than those).
// If aBuf is not NULL, caller has ensured that aBuf is at least MAX_NUMBER_SIZE in size.
{
	LPTSTR result;
	switch (aToken.symbol)
	{
	case SYM_VAR: // Caller has ensured that any SYM_VAR's Type() is VAR_NORMAL.
		result = aToken.var->Contents(); // Contents() vs. mCharContents in case mCharContents needs to be updated by Contents().
		if (aLength)
			*aLength = aToken.var->Length();
		return result;
	case SYM_STRING:
		result = aToken.marker;
		if (aLength)
		{
			if (aToken.marker_length == -1)
				break; // Call _tcslen(result) below.
			*aLength = aToken.marker_length;
		}
		return result;
	case SYM_INTEGER:
		result = aBuf ? ITOA64(aToken.value_int64, aBuf) : _T("");
		break;
	case SYM_FLOAT:
		if (aBuf)
		{
			int length = FTOA(aToken.value_double, aBuf, MAX_NUMBER_SIZE);
			if (aLength)
				*aLength = length;
			return aBuf;
		}
		//else continue on to return the default at the bottom.
	//case SYM_OBJECT: // Treat objects as empty strings (or TRUE where appropriate).
	//case SYM_MISSING:
	default:
		result = _T("");
	}
	if (aLength) // Caller wants to know the string's length as well.
		*aLength = _tcslen(result);
	return result;
}



ResultType TokenToDoubleOrInt64(const ExprTokenType &aInput, ExprTokenType &aOutput)
// Converts aToken's contents to a numeric value, either int64 or double (whichever is more appropriate).
// Returns FAIL when aToken isn't an operand or is but contains a string that isn't purely numeric.
{
	LPTSTR str;
	switch (aInput.symbol)
	{
		case SYM_INTEGER:
		case SYM_FLOAT:
			aOutput.symbol = aInput.symbol;
			aOutput.value_int64 = aInput.value_int64;
			return OK;
		case SYM_VAR:
			return aInput.var->ToDoubleOrInt64(aOutput);
		case SYM_STRING:   // v1.0.40.06: Fixed to be listed explicitly so that "default" case can return failure.
			str = aInput.marker;
			break;
		//case SYM_OBJECT: // L31: Treat objects as empty strings (or TRUE where appropriate).
		//case SYM_MISSING:
		default:
			return FAIL;
	}
	// Since above didn't return, interpret "str" as a number.
	switch (aOutput.symbol = IsNumeric(str, true, false, true))
	{
	case PURE_INTEGER:
		aOutput.value_int64 = ATOI64(str);
		break;
	case PURE_FLOAT:
		aOutput.value_double = ATOF(str);
		break;
	default: // Not a pure number.
		return FAIL;
	}
	return OK; // Since above didn't return, indicate success.
}



StringCaseSenseType TokenToStringCase(ExprTokenType& aToken)
{
	// Pure integers 1 and 0 corresponds to SCS_SENSITIVE and SCS_INSENSITIVE, respectively.
	// Pure floats returns SCS_INVALID.
	// For strings see Line::ConvertStringCaseSense.
	LPTSTR str = NULL;
	__int64 int_val = 0;
	switch (aToken.symbol)
	{
	case SYM_VAR:
		
		switch (aToken.var->IsPureNumeric())
		{
		case PURE_INTEGER: int_val = aToken.var->ToInt64(); break;
		case PURE_NOT_NUMERIC: str = aToken.var->Contents(); break;
		case PURE_FLOAT: 
		default:	
			return SCS_INVALID;
		}
		break;

	case SYM_INTEGER: int_val = TokenToInt64(aToken); break;
	case SYM_FLOAT: return SCS_INVALID;
	default: str = TokenToString(aToken); break;
	}
	if (str)
		return !_tcsicmp(str, _T("Logical"))	? SCS_INSENSITIVE_LOGICAL
												: Line::ConvertStringCaseSense(str);
	return int_val == 1 ? SCS_SENSITIVE						// 1	- Sensitive
						: (int_val == 0 ? SCS_INSENSITIVE	// 0	- Insensitive
										: SCS_INVALID);		// else - invalid.
}



Var *TokenToOutputVar(ExprTokenType &aToken)
{
	if (aToken.symbol == SYM_VAR && !VARREF_IS_READ(aToken.var_usage)) // VARREF_ISSET is tolerated for use by IsSet().
		return aToken.var;
	return dynamic_cast<VarRef *>(TokenToObject(aToken));
}



IObject *TokenToObject(ExprTokenType &aToken)
// L31: Returns IObject* from SYM_OBJECT or SYM_VAR (where var->HasObject()), NULL for other tokens.
// Caller is responsible for calling AddRef() if that is appropriate.
{
	if (aToken.symbol == SYM_OBJECT)
		return aToken.object;
	
	if (aToken.symbol == SYM_VAR)
		return aToken.var->ToObject();

	return NULL;
}



ResultType ValidateFunctor(IObject *aFunc, int aParamCount, ResultToken &aResultToken, int *aUseMinParams, bool aShowError)
{
	ASSERT(aFunc);
	__int64 min_params = 0, max_params = INT_MAX;
	auto min_result = aParamCount == -1 ? INVOKE_NOT_HANDLED
		: GetObjectIntProperty(aFunc, _T("MinParams"), min_params, aResultToken, true);
	if (!min_result)
		return FAIL;
	bool has_minparams = min_result != INVOKE_NOT_HANDLED; // For readability.
	if (aUseMinParams) // CallbackCreate's signal to default to MinParams.
	{
		if (!has_minparams)
			return aShowError ? aResultToken.UnknownMemberError(ExprTokenType(aFunc), IT_GET, _T("MinParams")) : CONDITION_FALSE;
		*aUseMinParams = aParamCount = (int)min_params;
	}
	else if (has_minparams && aParamCount < (int)min_params)
		return aShowError ? aResultToken.ValueError(ERR_INVALID_FUNCTOR) : CONDITION_FALSE;
	auto max_result = (aParamCount <= 0 || has_minparams && min_params == aParamCount)
		? INVOKE_NOT_HANDLED // No need to check MaxParams in the above cases.
		: GetObjectIntProperty(aFunc, _T("MaxParams"), max_params, aResultToken, true);
	if (!max_result)
		return FAIL;
	if (max_result != INVOKE_NOT_HANDLED && aParamCount > (int)max_params)
	{
		__int64 is_variadic = 0;
		auto result = GetObjectIntProperty(aFunc, _T("IsVariadic"), is_variadic, aResultToken, true);
		if (!result)
			return FAIL;
		if (!is_variadic) // or not defined.
			return aShowError ? aResultToken.ValueError(ERR_INVALID_FUNCTOR) : CONDITION_FALSE;
	}
	// If either MinParams or MaxParams was confirmed to exist, this is likely a valid
	// function object, so skip the following check for performance.  Otherwise, catch
	// likely errors by checking that the object is callable.
	if (min_result == INVOKE_NOT_HANDLED && max_result == INVOKE_NOT_HANDLED)
		if (Object *obj = dynamic_cast<Object *>(aFunc))
			if (!obj->HasMethod(_T("Call")))
				return aShowError ? aResultToken.UnknownMemberError(ExprTokenType(aFunc), IT_CALL, _T("Call")) : CONDITION_FALSE;
		// Otherwise: COM objects can be callable via DISPID_VALUE.  There's probably
		// no way to determine whether the object supports that without invoking it.
	return OK;
}



ResultType TokenSetResult(ResultToken &aResultToken, LPCTSTR aValue, size_t aLength)
// Utility function for handling memory allocation and return to callers of built-in functions; based on BIF_SubStr.
// Returns FAIL if malloc failed, in which case our caller is responsible for returning a sensible default value.
{
	if (aLength == -1)
		aLength = _tcslen(aValue); // Caller must not pass NULL for aResult in this case.
	if (aLength <= MAX_NUMBER_LENGTH) // v1.0.46.01: Avoid malloc() for small strings.  However, this improves speed by only 10% in a test where random 25-byte strings were extracted from a 700 KB string (probably because VC++'s malloc()/free() are very fast for small allocations).
		aResultToken.marker = aResultToken.buf; // Store the address of the result for the caller.
	else
	{
		// Caller has provided a mem_to_free (initially NULL) as a means of passing back memory we allocate here.
		// So if we change "result" to be non-NULL, the caller will take over responsibility for freeing that memory.
		if (   !(aResultToken.mem_to_free = tmalloc(aLength + 1))   ) // Out of memory.
			return aResultToken.MemoryError();
		aResultToken.marker = aResultToken.mem_to_free; // Store the address of the result for the caller.
	}
	if (aValue) // Caller may pass NULL to retrieve a buffer of sufficient size.
		tmemcpy(aResultToken.marker, aValue, aLength);
	aResultToken.marker[aLength] = '\0'; // Must be done separately from the memcpy() because the memcpy() might just be taking a substring (i.e. long before result's terminator).
	aResultToken.marker_length = aLength;
	return OK;
}



// TypeOfToken: Similar result to TokenIsPureNumeric, but may return SYM_OBJECT.
SymbolType TypeOfToken(ExprTokenType &aToken)
{
	switch (aToken.symbol)
	{
	case SYM_VAR:
		switch (aToken.var->IsPureNumericOrObject())
		{
		case VAR_ATTRIB_IS_INT64: return SYM_INTEGER;
		case VAR_ATTRIB_IS_DOUBLE: return SYM_FLOAT;
		case VAR_ATTRIB_IS_OBJECT: return SYM_OBJECT;
		}
		return SYM_STRING;
	default: // Providing a default case produces smaller code on release builds as the compiler can omit the other symbol checks.
#ifdef _DEBUG
		MsgBox(_T("DEBUG: Unhandled symbol type."));
#endif
	case SYM_STRING:
	case SYM_INTEGER:
	case SYM_FLOAT:
	case SYM_OBJECT:
	case SYM_MISSING:
		return aToken.symbol;
	}
}



ResultType ResultToken::Return(LPTSTR aValue, size_t aLength)
// Copy and return a string.
{
	ASSERT(aValue);
	symbol = SYM_STRING;
	return TokenSetResult(*this, aValue, aLength);
}



int ConvertJoy(LPTSTR aBuf, int *aJoystickID, bool aAllowOnlyButtons)
// The caller TextToKey() currently relies on the fact that when aAllowOnlyButtons==true, a value
// that can fit in a sc_type (USHORT) is returned, which is true since the joystick buttons
// are very small numbers (JOYCTRL_1==12).
{
	if (aJoystickID)
		*aJoystickID = 0;  // Set default output value for the caller.
	if (!aBuf || !*aBuf) return JOYCTRL_INVALID;
	LPTSTR aBuf_orig = aBuf;
	for (; *aBuf >= '0' && *aBuf <= '9'; ++aBuf); // self-contained loop to find the first non-digit.
	if (aBuf > aBuf_orig) // The string starts with a number.
	{
		int joystick_id = ATOI(aBuf_orig) - 1;
		if (joystick_id < 0 || joystick_id >= MAX_JOYSTICKS)
			return JOYCTRL_INVALID;
		if (aJoystickID)
			*aJoystickID = joystick_id;  // Use ATOI vs. atoi even though hex isn't supported yet.
	}

	if (!_tcsnicmp(aBuf, _T("Joy"), 3))
	{
		if (IsNumeric(aBuf + 3, false, false))
		{
			int offset = ATOI(aBuf + 3);
			if (offset < 1 || offset > MAX_JOY_BUTTONS)
				return JOYCTRL_INVALID;
			return JOYCTRL_1 + offset - 1;
		}
	}
	if (aAllowOnlyButtons)
		return JOYCTRL_INVALID;

	// Otherwise:
	if (!_tcsicmp(aBuf, _T("JoyX"))) return JOYCTRL_XPOS;
	if (!_tcsicmp(aBuf, _T("JoyY"))) return JOYCTRL_YPOS;
	if (!_tcsicmp(aBuf, _T("JoyZ"))) return JOYCTRL_ZPOS;
	if (!_tcsicmp(aBuf, _T("JoyR"))) return JOYCTRL_RPOS;
	if (!_tcsicmp(aBuf, _T("JoyU"))) return JOYCTRL_UPOS;
	if (!_tcsicmp(aBuf, _T("JoyV"))) return JOYCTRL_VPOS;
	if (!_tcsicmp(aBuf, _T("JoyPOV"))) return JOYCTRL_POV;
	if (!_tcsicmp(aBuf, _T("JoyName"))) return JOYCTRL_NAME;
	if (!_tcsicmp(aBuf, _T("JoyButtons"))) return JOYCTRL_BUTTONS;
	if (!_tcsicmp(aBuf, _T("JoyAxes"))) return JOYCTRL_AXES;
	if (!_tcsicmp(aBuf, _T("JoyInfo"))) return JOYCTRL_INFO;
	return JOYCTRL_INVALID;
}



bool ScriptGetKeyState(vk_type aVK, KeyStateTypes aKeyStateType)
// Returns true if "down", false if "up".
{
    if (!aVK) // Assume "up" if indeterminate.
		return false;

	switch (aKeyStateType)
	{
	case KEYSTATE_TOGGLE: // Whether a toggleable key such as CapsLock is currently turned on.
		return IsKeyToggledOn(aVK); // This also works for non-"lock" keys, but in that case the toggle state can be out of sync with other processes/threads.
	case KEYSTATE_PHYSICAL: // Physical state of key.
		if (IsMouseVK(aVK)) // mouse button
		{
			if (g_MouseHook) // mouse hook is installed, so use it's tracking of physical state.
				return g_PhysicalKeyState[aVK] & STATE_DOWN;
			else
				return IsKeyDownAsync(aVK);
		}
		else // keyboard
		{
			if (g_KeybdHook)
			{
				// Since the hook is installed, use its value rather than that from
				// GetAsyncKeyState(), which doesn't seem to return the physical state.
				// But first, correct the hook modifier state if it needs it.  See comments
				// in GetModifierLRState() for why this is needed:
				if (KeyToModifiersLR(aVK))    // It's a modifier.
					GetModifierLRState(true); // Correct hook's physical state if needed.
				return g_PhysicalKeyState[aVK] & STATE_DOWN;
			}
			else
				return IsKeyDownAsync(aVK);
		}
	} // switch()

	// Otherwise, use the default state-type: KEYSTATE_LOGICAL
	// On XP/2K at least, a key can be physically down even if it isn't logically down,
	// which is why the below specifically calls IsKeyDown() rather than some more
	// comprehensive method such as consulting the physical key state as tracked by the hook:
	// v1.0.42.01: For backward compatibility, the following hasn't been changed to IsKeyDownAsync().
	// One example is the journal playback hook: when a window owned by the script receives
	// such a keystroke, only GetKeyState() can detect the changed state of the key, not GetAsyncKeyState().
	// A new mode can be added to KeyWait & GetKeyState if Async is ever explicitly needed.
	return IsKeyDown(aVK);
}



bool ScriptGetJoyState(JoyControls aJoy, int aJoystickID, ExprTokenType &aToken, LPTSTR aBuf)
// Caller must ensure that aToken.marker is a buffer large enough to handle the longest thing put into
// it here, which is currently jc.szPname (size=32). Caller has set aToken.symbol to be SYM_STRING.
// aToken is used for the value being returned by GetKeyState() to the script, while this function's
// bool return value is used only by KeyWait, so is false for "up" and true for "down".
// If there was a problem determining the position/state, aToken is made blank and false is returned.
{
	// Set default in case of early return.
	aToken.symbol = SYM_STRING;
	aToken.marker = aBuf;
	*aBuf = '\0'; // Blank vs. string "0" serves as an indication of failure.

	if (!aJoy) // Currently never called this way.
		return false; // And leave aToken set to blank.

	bool aJoy_is_button = IS_JOYSTICK_BUTTON(aJoy);

	JOYCAPS jc;
	if (!aJoy_is_button && aJoy != JOYCTRL_POV)
	{
		// Get the joystick's range of motion so that we can report position as a percentage.
		if (joyGetDevCaps(aJoystickID, &jc, sizeof(JOYCAPS)) != JOYERR_NOERROR)
			ZeroMemory(&jc, sizeof(jc));  // Zero it on failure, for use of the zeroes later below.
	}

	// Fetch this struct's info only if needed:
	JOYINFOEX jie;
	if (aJoy != JOYCTRL_NAME && aJoy != JOYCTRL_BUTTONS && aJoy != JOYCTRL_AXES && aJoy != JOYCTRL_INFO)
	{
		jie.dwSize = sizeof(JOYINFOEX);
		jie.dwFlags = JOY_RETURNALL;
		if (joyGetPosEx(aJoystickID, &jie) != JOYERR_NOERROR)
			return false; // And leave aToken set to blank.
		if (aJoy_is_button)
		{
			bool is_down = ((jie.dwButtons >> (aJoy - JOYCTRL_1)) & (DWORD)0x01);
			aToken.symbol = SYM_INTEGER; // Override default type.
			aToken.value_int64 = is_down; // Always 1 or 0, since it's "bool" (and if not for that, bitwise-and).
			return is_down;
		}
	}

	// Otherwise:
	UINT range;
	LPTSTR buf_ptr;
	double result_double;  // Not initialized to help catch bugs.

	switch(aJoy)
	{
	case JOYCTRL_XPOS:
		range = (jc.wXmax > jc.wXmin) ? jc.wXmax - jc.wXmin : 0;
		result_double = range ? 100 * (double)jie.dwXpos / range : jie.dwXpos;
		break;
	case JOYCTRL_YPOS:
		range = (jc.wYmax > jc.wYmin) ? jc.wYmax - jc.wYmin : 0;
		result_double = range ? 100 * (double)jie.dwYpos / range : jie.dwYpos;
		break;
	case JOYCTRL_ZPOS:
		range = (jc.wZmax > jc.wZmin) ? jc.wZmax - jc.wZmin : 0;
		result_double = range ? 100 * (double)jie.dwZpos / range : jie.dwZpos;
		break;
	case JOYCTRL_RPOS:  // Rudder or 4th axis.
		range = (jc.wRmax > jc.wRmin) ? jc.wRmax - jc.wRmin : 0;
		result_double = range ? 100 * (double)jie.dwRpos / range : jie.dwRpos;
		break;
	case JOYCTRL_UPOS:  // 5th axis.
		range = (jc.wUmax > jc.wUmin) ? jc.wUmax - jc.wUmin : 0;
		result_double = range ? 100 * (double)jie.dwUpos / range : jie.dwUpos;
		break;
	case JOYCTRL_VPOS:  // 6th axis.
		range = (jc.wVmax > jc.wVmin) ? jc.wVmax - jc.wVmin : 0;
		result_double = range ? 100 * (double)jie.dwVpos / range : jie.dwVpos;
		break;

	case JOYCTRL_POV:
		aToken.symbol = SYM_INTEGER; // Override default type.
		if (jie.dwPOV == JOY_POVCENTERED) // Need to explicitly compare against JOY_POVCENTERED because it's a WORD not a DWORD.
		{
			aToken.value_int64 = -1;
			return false;
		}
		else
		{
			aToken.value_int64 = jie.dwPOV;
			return true;
		}
		// No break since above always returns.

	case JOYCTRL_NAME:
		_tcscpy(aToken.marker, jc.szPname);
		return false; // Return value not used.

	case JOYCTRL_BUTTONS:
		aToken.symbol = SYM_INTEGER; // Override default type.
		aToken.value_int64 = jc.wNumButtons; // wMaxButtons is the *driver's* max supported buttons.
		return false; // Return value not used.

	case JOYCTRL_AXES:
		aToken.symbol = SYM_INTEGER; // Override default type.
		aToken.value_int64 = jc.wNumAxes; // wMaxAxes is the *driver's* max supported axes.
		return false; // Return value not used.

	case JOYCTRL_INFO:
		buf_ptr = aToken.marker;
		if (jc.wCaps & JOYCAPS_HASZ)
			*buf_ptr++ = 'Z';
		if (jc.wCaps & JOYCAPS_HASR)
			*buf_ptr++ = 'R';
		if (jc.wCaps & JOYCAPS_HASU)
			*buf_ptr++ = 'U';
		if (jc.wCaps & JOYCAPS_HASV)
			*buf_ptr++ = 'V';
		if (jc.wCaps & JOYCAPS_HASPOV)
		{
			*buf_ptr++ = 'P';
			if (jc.wCaps & JOYCAPS_POV4DIR)
				*buf_ptr++ = 'D';
			if (jc.wCaps & JOYCAPS_POVCTS)
				*buf_ptr++ = 'C';
		}
		*buf_ptr = '\0'; // Final termination.
		return false; // Return value not used.
	} // switch()

	// If above didn't return, the result should now be in result_double.
	aToken.symbol = SYM_FLOAT; // Override default type.
	aToken.value_double = result_double;
	return result_double;
}



__int64 pow_ll(__int64 base, __int64 exp)
{
	/*
	Caller must ensure exp >= 0
	Below uses 'a^b' to denote 'raising a to the power of b'.
	Computes and returns base^exp. If the mathematical result doesn't fit in __int64, the result is undefined.
	By convention, x^0 returns 1, even when x == 0,	caller should ensure base is non-zero when exp is zero to handle 0^0.
	*/
	if (exp == 0)
		return 1ll;

	// based on: https://en.wikipedia.org/wiki/Exponentiation_by_squaring (2018-11-03)
	__int64 result = 1;
	while (exp > 1)
	{
		if (exp % 2) // exp is odd
			result *= base;
		base *= base;
		exp /= 2;
	}
	return result * base;
}
