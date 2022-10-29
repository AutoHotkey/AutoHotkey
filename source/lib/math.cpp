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

#include "stdafx.h"
#include "qmath.h"
#include "script.h"
#include <stdint.h>

#include "script_func_impl.h"



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



FResult DateAdd(StrArg aDateTime, double aTime, StrArg aTimeUnits, StrRet &aRetVal)
{
	FILETIME ft;
	if (!YYYYMMDDToFileTime(aDateTime, ft))
		return FR_E_ARG(0);

	// Use double to support a floating point value for days, hours, minutes, etc:
	double nUnits = aTime;

	// Convert to seconds:
	switch (ctoupper(*aTimeUnits))
	{
	case 'S': // Seconds
		break;
	case 'M': // Minutes
		nUnits *= ((double)60);
		break;
	case 'H': // Hours
		nUnits *= ((double)60 * 60);
		break;
	case 'D': // Days
		nUnits *= ((double)60 * 60 * 24);
		break;
	default: // Invalid
		return FR_E_ARG(2);
	}
	// Convert ft struct to a 64-bit variable (maybe there's some way to avoid these conversions):
	ULARGE_INTEGER ul;
	ul.LowPart = ft.dwLowDateTime;
	ul.HighPart = ft.dwHighDateTime;
	// Prior to adding nUnits to the result value, convert it from seconds to 10ths of a
	// microsecond (the units of the FILETIME struct).  Use int64 multiplication to avoid
	// floating-point rounding errors, such as with values of seconds > 115292150460.
	// Testing shows this keeps precision beyond year 9999.  Truncating any fractional part
	// is fine at this point because the resulting string only includes whole seconds.
	ul.QuadPart += (__int64)nUnits * 10000000;
	// Convert back into ft struct:
	ft.dwLowDateTime = ul.LowPart;
	ft.dwHighDateTime = ul.HighPart;
	aRetVal.SetTemp(FileTimeToYYYYMMDD(aRetVal.CallerBuf(), ft, false));
	return OK;
}



FResult DateDiff(StrArg aTime1, StrArg aTime2, StrArg aTimeUnits, __int64 &aRetVal)
{
	FResult fr;
	// If either parameter is blank, it will default to the current time:
	__int64 time_until = YYYYMMDDSecondsUntil(aTime2, aTime1, fr);
	if (fr != OK) // Usually caused by an invalid component in the date-time string.
		return fr;
	switch (ctoupper(*aTimeUnits))
	{
	case 'S': break;
	case 'M': time_until /= 60; break; // Minutes
	case 'H': time_until /= 60 * 60; break; // Hours
	case 'D': time_until /= 60 * 60 * 24; break; // Days
	default: // Invalid
		return FR_E_ARG(2);
	}
	aRetVal = time_until;
	return OK;
}
