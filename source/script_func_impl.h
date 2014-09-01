
#define ParamIndexToString(index, ...)				TokenToString(*aParam[(index)], __VA_ARGS__)
#define ParamIndexToInt64(index)					TokenToInt64(*aParam[(index)])
#define ParamIndexToInt(index)						(int)ParamIndexToInt64(index)
#define ParamIndexToIntPtr(index)					(INT_PTR)ParamIndexToInt64(index)
#define ParamIndexToDouble(index)					TokenToDouble(*aParam[(index)])
#define ParamIndexToNumber(index, output)			TokenToDoubleOrInt64(*aParam[(index)], output)

#define ParamIndexLength(index, param_as_string)	EXPR_TOKEN_LENGTH(aParam[index], param_as_string)

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
#define ParamIndexToOptionalVar(index)				(((index) < aParamCount && aParam[index]->symbol == SYM_VAR) ? aParam[index]->var : NULL)

// The macro below defaults to "", since that is by far the most common default.
// This allows it to skip the check for SYM_MISSING, which always has marker == _T("").
#define ParamIndexToOptionalString(index, ...)		(((index) < aParamCount) ? ParamIndexToString(index, __VA_ARGS__) : _T(""))
