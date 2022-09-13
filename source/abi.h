#pragma once

#include "StrRet.h"

// Marks functions which are called by script, for clarity and possible future use.
#define bif_impl

typedef HRESULT FResult;

constexpr int FR_OUR_FLAG = 0x20000000;
constexpr int FR_INT_FLAG = 0x40000000;
constexpr int FR_FACILITY_CONTROL = 0;
constexpr int FR_FACILITY_INT = 9;
constexpr int FR_FACILITY_ARG = 0xA;
constexpr int FR_FACILITY_ERR = 0xE;
constexpr FResult FR_FAIL = FR_OUR_FLAG | MAKE_HRESULT(SEVERITY_ERROR, FR_FACILITY_CONTROL, 0); // Error already shown/thrown.
constexpr FResult FR_ABORTED = FR_OUR_FLAG | MAKE_HRESULT(SEVERITY_SUCCESS, FR_FACILITY_CONTROL, 0); // Continuing after an error; return blank.

constexpr FResult FR_E_WIN32 = FR_OUR_FLAG | MAKE_HRESULT(SEVERITY_ERROR, FACILITY_WIN32, 0); // 0xA0070000
#define FR_E_WIN32(n) (FR_E_WIN32 | (n))
constexpr FResult FR_E_ARG_ZERO = FR_OUR_FLAG | MAKE_HRESULT(SEVERITY_ERROR, FR_FACILITY_ARG, 0); // 0xA00A0000
#define FR_E_ARG(n) (FR_E_ARG_ZERO | (n))
constexpr FResult FR_E_ARGS = FR_E_ARG(0xFFFF);

#define FR_THROW_INT(n) ((n) & 0xF0000000 ? FR_E_FAILED \
	: ((SEVERITY_ERROR << 31) | FR_OUR_FLAG | FR_INT_FLAG | (n))) // Throw an Error where Extra = an int in the range 0..0x0FFFFFFF.
#define FR_GET_THROWN_INT(fr) ((fr) & 0x0FFFFFFF)

constexpr int FR_ERR_BASE = FR_OUR_FLAG | MAKE_HRESULT(SEVERITY_ERROR, FR_FACILITY_ERR, 0); // 0xA00E0000
constexpr int FR_E_OUTOFMEM = FR_ERR_BASE | 1;
constexpr int FR_E_FAILED = FR_ERR_BASE | 2;


typedef LPCTSTR StrArg;

template<typename T> class optl
{
	const T *_value;
public:
	optl(T &v) : _value {&v} {}
	optl(nullptr_t) : _value {nullptr} {}
	bool has_value() { return _value != nullptr; }
	T operator* () { return *_value; }
	T value() { return *_value; }
	T value_or(T aDefault) { return _value ? *_value : aDefault; }
};

template<> class optl<StrArg>
{
	const StrArg _value;
public:
	optl(StrArg v) : _value {v} {}
	bool has_value() { return _value != nullptr; }
	bool has_nonempty_value() { return _value && *_value; }
	bool is_blank() { return _value && !*_value; }
	bool is_blank_or_omitted() { return !has_nonempty_value(); }
	StrArg value() { ASSERT(_value); return _value; }
	StrArg value_or(StrArg aDefault) { return _value ? _value : aDefault; }
	StrArg value_or_null() { return _value; }
	StrArg value_or_empty() { return _value ? _value : _T(""); }
};
