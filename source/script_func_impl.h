
#define ParamIndexToString(index, ...)				TokenToString(*aParam[(index)], __VA_ARGS__)
#define ParamIndexToInt64(index)					TokenToInt64(*aParam[(index)])
#define ParamIndexToInt(index)						(int)ParamIndexToInt64(index)
#define ParamIndexToIntPtr(index)					(INT_PTR)ParamIndexToInt64(index)
#define ParamIndexToDouble(index)					TokenToDouble(*aParam[(index)])
#define ParamIndexToNumber(index, output)			TokenToDoubleOrInt64(*aParam[(index)], output)
#define ParamIndexToBOOL(index)						TokenToBOOL(*aParam[(index)])
#define ParamIndexToObject(index)					TokenToObject(*aParam[(index)])
#define ParamIndexToToggleValue(index)				TokenToToggleValue(*aParam[(index)])

#define ParamIndexToOutputVar(index)				((index) < aParamCount ? TokenToOutputVar(*aParam[(index)]) : nullptr)

// Omitted parameter returns SCS_INSENSITIVE.
#define ParamIndexToCaseSense(index)				(ParamIndexIsOmitted((index)) ? SCS_INSENSITIVE : TokenToStringCase(*aParam[(index)]))

#define ParamIndexIsNumeric(index)  (TokenIsNumeric(*aParam[(index)]))

// For functions that allow "" to mean parameter is omitted.
#define ParamIndexIsOmittedOrEmpty(index)  (ParamIndexIsOmitted(index) || TokenIsEmptyString(*aParam[(index)]))

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

inline LPTSTR _OptionalStringDefaultHelper(LPTSTR aDef, LPTSTR aBuf = NULL, size_t *aLength = NULL)
{
	if (aLength)
		*aLength = _tcslen(aDef);
	return aDef;
}
#define ParamIndexToOptionalStringDef(index, def, ...) \
	(ParamIndexIsOmitted(index) ? _OptionalStringDefaultHelper(def, __VA_ARGS__) \
								: ParamIndexToString(index, __VA_ARGS__))
// The macro below defaults to "", since that is by far the most common default.
// This allows it to skip the check for SYM_MISSING, which always has marker == _T("").
#define ParamIndexToOptionalString(index, ...) \
	(((index) < aParamCount) ? ParamIndexToString(index, __VA_ARGS__) \
							: _OptionalStringDefaultHelper(_T(""), __VA_ARGS__))

#define ParamIndexToOptionalObject(index)			((index) < aParamCount ? ParamIndexToObject(index) : NULL)

#define _f_param_string(name, index, ...) \
	TCHAR name##_buf[MAX_NUMBER_SIZE], *name = ParamIndexToString(index, name##_buf, __VA_ARGS__)
#define _f_param_string_opt(name, index, ...) \
	TCHAR name##_buf[MAX_NUMBER_SIZE], *name = ParamIndexToOptionalString(index, name##_buf, __VA_ARGS__)
#define _f_param_string_opt_def(name, index, def, ...) \
	TCHAR name##_buf[MAX_NUMBER_SIZE], *name = ParamIndexToOptionalStringDef(index, def, name##_buf, __VA_ARGS__)

#define Throw_if_Param_NaN(ParamIndex) \
	if (!TokenIsNumeric(*aParam[(ParamIndex)])) \
		_f_throw_param((ParamIndex), _T("Number"))


#define Throw_if_RValue_NaN() \
	if (!TokenIsNumeric(aValue)) \
		_f_throw_type(_T("Number"), aValue)

#define BivRValueToString(...)  TokenToString(aValue, _f_number_buf, __VA_ARGS__)
#define BivRValueToInt64()  TokenToInt64(aValue)
#define BivRValueToBOOL()  TokenToBOOL(aValue)
#define BivRValueToObject()  TokenToObject(aValue)


template<class T>
BIF_DECL(NewObject)
{
	Object *obj = T::Create();
	if (!obj)
		_f_throw_oom;
	obj->New(aResultToken, aParam, aParamCount);
}
