
#define ParamIndexToString(index, ...)				TokenToString(*aParam[(index)], __VA_ARGS__)
#define ParamIndexToInt64(index)					TokenToInt64(*aParam[(index)])
#define ParamIndexToInt(index)						(int)ParamIndexToInt64(index)
#define ParamIndexToIntPtr(index)					(INT_PTR)ParamIndexToInt64(index)
#define ParamIndexToDouble(index)					TokenToDouble(*aParam[(index)])
#define ParamIndexToNumber(index, output)			TokenToDoubleOrInt64(*aParam[(index)], output)
#define ParamIndexToBOOL(index)						TokenToBOOL(*aParam[(index)])
#define ParamIndexToObject(index)					TokenToObject(*aParam[(index)])

// For functions that allow "" to mean parameter is omitted.
#define ParamIndexIsOmittedOrEmpty(index)  (ParamIndexIsOmitted(index) || TokenIsEmptyString(*aParam[(index)], TRUE))

// For functions that don't allow "" to mean parameter is omitted.
#define ParamIndexIsOmitted(index)  ((index) >= aParamCount || aParam[(index)]->symbol == SYM_MISSING)

#define PASTE(a, b) a##b

#define ParamIndexToOptionalType(type, index, def)	(ParamIndexIsOmitted(index) ? (def) : PASTE(ParamIndexTo,type)(index))
#define ParamIndexToOptionalInt(index, def)			ParamIndexToOptionalType(Int, index, def)
#define ParamIndexToOptionalIntPtr(index, def)		ParamIndexToOptionalType(IntPtr, index, def)
#define ParamIndexToOptionalDouble(index, def)		ParamIndexToOptionalType(Double, index, def)
#define ParamIndexToOptionalInt64(index, def)		ParamIndexToOptionalType(Int64, index, def)
#define ParamIndexToOptionalBOOL(index, def)		ParamIndexToOptionalType(BOOL, index, def)
#define ParamIndexToOptionalVar(index)				(((index) < aParamCount && aParam[index]->symbol == SYM_VAR) ? aParam[index]->var : NULL)


// Rather than adding a third value to the CaseSensitive parameter, it obeys StringCaseSense because:
// 1) It matches the behavior of the equal operator (=) in expressions.
// 2) It's more friendly for typical international uses because it avoids having to specify that special/third value
//    for every call of InStr.  It's nice to be able to omit the CaseSensitive parameter every time and know that
//    the behavior of both InStr and its counterpart the equals operator are always consistent with each other.
// If the parameter is (false or omitted) insensitive, resolve it to be Locale-mode if the StringCaseSense mode is
// either case-sensitive or Locale-insensitive.
// The parameter is assumed to be optional and always defaults to false.
#define ParamIndexToCaseSense(index)				(!ParamIndexIsOmitted(index) && ParamIndexToBOOL(index) ? SCS_SENSITIVE \
													: (g->StringCaseSense != SCS_INSENSITIVE ? SCS_INSENSITIVE_LOCALE : SCS_INSENSITIVE) )

#define ParamIndexToOptionalStringDef(index, def, ...)	(ParamIndexIsOmitted(index) ? (def) : ParamIndexToString(index, __VA_ARGS__))
// The macro below defaults to "", since that is by far the most common default.
// This allows it to skip the check for SYM_MISSING, which always has marker == _T("").
#define ParamIndexToOptionalString(index, ...)		(((index) < aParamCount) ? ParamIndexToString(index, __VA_ARGS__) : _T(""))

#define ParamIndexToOptionalObject(index)			((index) < aParamCount ? ParamIndexToObject(index) : NULL)

#define _f_param_string(name, index, ...) \
	TCHAR name##_buf[MAX_NUMBER_SIZE], *name = ParamIndexToString(index, name##_buf, __VA_ARGS__)
#define _f_param_string_opt(name, index, ...) \
	TCHAR name##_buf[MAX_NUMBER_SIZE], *name = ParamIndexToOptionalString(index, name##_buf, __VA_ARGS__)
#define _f_param_string_opt_def(name, index, def, ...) \
	TCHAR name##_buf[MAX_NUMBER_SIZE], *name = ParamIndexToOptionalStringDef(index, def, name##_buf, __VA_ARGS__)
