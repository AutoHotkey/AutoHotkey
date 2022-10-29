#pragma once

#include "abi.h"

struct ExprTokenType;
struct ResultToken;
struct IObject;

enum class MdType : UINT8
{
	Void		= 0,
	TypeInt		= 1,
	TypeFloat	= 2,
	Size32		= 0x10, // This is just so Int32 != TypeInt, for debugging purposes.
	Size64		= 0x20, // Flag for types that are 64-bit even on x86-32.
	Unsigned	= 0x40,
	Int32		= TypeInt | Size32,
	UInt32		= TypeInt | Size32 | Unsigned,
	Int64		= TypeInt | Size64,
	UInt64		= TypeInt | Size64 | Unsigned,
	Float64		= TypeFloat | Size64,
	String		= 3,
	Object,
	Variant, // Currently only for input (ExprTokenType) or retval (ResultToken).
	Bool32,
	ResultType,
	FResult,
	NzIntWin32, // BOOL result where FALSE means failure and GetLastError() is applicable.
	Params,
	TypeMask	= 0xF,
	BitsBase	= 99, // For encoding a small literal value to insert into the parameter list.
	Optional	= 0x80,
	RetVal,
	ThisCall, // Only valid at the beginning of the args.
	Out,
	// Only aliases from here on
	FirstModifier = Optional,
	BitsUpperBound = Optional,
	UIntPtr = Exp32or64(UInt32, UInt64),
	IntPtr = Exp32or64(Int32, Int64)
};

#define MdType_Type(t) ((MdType)((int)(t) & (int)MdType::TypeMask))
#define MdType_IsInt(t) (MdType_Type(t) == MdType::TypeInt)
#define MdType_IsFloat(t) (MdType_Type(t) == MdType::TypeFloat)
#define MdType_IsNum(t) (MdType_IsInt(t) || MdType_IsFloat(t))
#define MdType_Is64bit(t) (((int)(t) & (int)MdType::Size64) != 0)
#define MdType_IsMod(t) ((t) >= MdType::FirstModifier)

#define MdType_IsOut(t) ((t) == MdType::Out) // Macro supports the future addition of other Out modifiers.

#define MdType_IsBits(t) ((t) >= MdType::BitsBase && (t) < MdType::BitsUpperBound)
#define MdType_Bits(t) ((MdType)(static_cast<int>(t) + ((int)MdType::BitsBase + 1)))
#define MdType_BitsValue(t) (static_cast<int>(t) - ((int)MdType::BitsBase + 1)) // t100 = 0


template<MdType T> struct md_argtype;
template<> struct md_argtype<MdType::Int32> { typedef int t; };
template<> struct md_argtype<MdType::UInt32> { typedef UINT t; };
template<> struct md_argtype<MdType::Float64> { typedef double t; };
template<> struct md_argtype<MdType::Int64> { typedef __int64 t; };
template<> struct md_argtype<MdType::UInt64> { typedef UINT64 t; };
template<> struct md_argtype<MdType::String> { typedef StrArg t; };
template<> struct md_argtype<MdType::Object> { typedef IObject *t; };
template<> struct md_argtype<MdType::Void> { typedef void t; };
template<> struct md_argtype<MdType::Variant> { typedef ExprTokenType &t; };
template<> struct md_argtype<MdType::Bool32> { typedef BOOL t; };
template<> struct md_argtype<MdType::Params> { typedef VariantParams &t; };

template<MdType T = MdType::Int32> struct md_outtype { typedef typename md_argtype<T>::t &t; };
template<> struct md_outtype<MdType::String> { typedef StrRet &t; };
template<> struct md_outtype<MdType::Variant> { typedef ResultToken &t; };

template<MdType T> struct md_optout { typedef typename md_argtype<T>::t *t; };
template<> struct md_optout<MdType::String> { typedef StrRet *t; };
template<> struct md_optout<MdType::Variant> { typedef ResultToken *t; };

template<MdType T> struct md_optional { typedef typename optl<typename md_argtype<T>::t> t; };
template<> struct md_optional<MdType::Variant> { typedef ExprTokenType *t; };

template<MdType T> struct md_retval { typedef typename md_argtype<T>::t t; };
template<> struct md_retval<MdType::FResult> { typedef FResult t; };
template<> struct md_retval<MdType::ResultType> { typedef ResultType t; };
template<> struct md_retval<MdType::NzIntWin32> { typedef BOOL t; };

#include "map.h"

#define md_cat(a, ...) a ## __VA_ARGS__

#define md_arg_decl_type_In(type) md_argtype<MdType::type>::t
#define md_arg_decl_type_In_Opt(type) md_optional<MdType::type>::t
#define md_arg_decl_type_Out(type) md_outtype<MdType::type>::t
#define md_arg_decl_type_Out_Opt(type) md_optout<MdType::type>::t
#define md_arg_decl_type_Ret(type) md_arg_decl_type_Out(type)

#define md_arg_decl_(mod, type, name) md_arg_decl_type_##mod(type) name
#define md_arg_decl(p) md_arg_decl_ p

#define md_arg_decl_type(mod, type) md_arg_decl_type_##mod(type)

#define md_arg_none (In, Void, )


//#define md_opt_arg(type, name)		md_optional<MdType::type>::t name
//#define md_retval_StrRet StrRet aRetVal

#define md_arg_data_In
#define md_arg_data_In_Opt MdType::Optional,
#define md_arg_data_Out MdType::Out,
#define md_arg_data_Out_Opt MdType::Out, MdType::Optional,
#define md_arg_data_Ret MdType::Out, MdType::RetVal,
#define md_arg_data_(mod, type, name) md_arg_data_##mod MdType::type
#define md_arg_data(p) md_arg_data_ p

#define md_func_decl(script_name, native_name, retcode, ...) \
	md_retval<MdType::retcode>::t native_name( \
		MAP_LIST(md_arg_decl, __VA_ARGS__) );

// md_func_cast(...)(functionName) is used for overload resolution.
#define md_func_cast(retcode, ...) \
	static_cast<md_retval<MdType::retcode>::t (*)( MAP_LIST(md_arg_decl, __VA_ARGS__) )>

#define md_func_data(script_name, native_name, retcode, ...) \
	{ \
		_T(#script_name), \
		md_func_cast(retcode, __VA_ARGS__)(native_name), \
		MdType::retcode, \
		{ MAP_LIST(md_arg_data, __VA_ARGS__) } \
	},

#define md_func_x(script_name, native_name, retcode, ...) \
	md_cat(md_func_,md_mode)(script_name, native_name, retcode, __VA_ARGS__)

#define md_func(name, ...) \
	md_func_x(name, name, FResult, __VA_ARGS__)

#define md_func_v(name, ...) \
	md_func_x(name, name, Void, __VA_ARGS__)

#ifdef __INTELLISENSE__
#undef md_func_x
#define md_func_x(...) md_func_decl(__VA_ARGS__)
//#define md_func_x(...) md_func_data(__VA_ARGS__)
#endif


// For use reinterpreting member function pointers (not standard C++).
template<typename T> constexpr void* cast_into_voidp(T in)
{
	union { T in; void *out; } u { in };
	return u.out;
}

#define md_member_name_GET(base_name) get_##base_name
#define md_member_name_SET(base_name) set_##base_name
#define md_member_name_CALL(base_name) base_name

// The member function pointer type can be inferred by the template, but we don't want that;
// md_member_type() is used below to ensure the signature matches the metadata.
#define md_member_type(class_name, ...) FResult (class_name::*)( MAP_LIST(md_arg_decl, __VA_ARGS__) )

#define md_member(class_name, member_name, invoke_type, ...) \
	 md_member_x(class_name, member_name, member_name, invoke_type, __VA_ARGS__)
#define md_member_x(class_name, member_name, impl, invoke_type, ...) \
	{ _T(#member_name), \
		cast_into_voidp<md_member_type(class_name, __VA_ARGS__)>(&class_name::md_cat(md_member_name_,invoke_type)(impl)), \
		IT_##invoke_type, \
		{ MAP_LIST(md_arg_data, __VA_ARGS__) } }

#define md_property_get(class_name, member_name, arg_type) \
	md_member(class_name, member_name, GET, (Ret, arg_type, RetVal))
#define md_property_set(class_name, member_name, arg_type) \
	md_member(class_name, member_name, SET, (In, arg_type, Value))
#define md_property(class_name, member_name, arg_type) \
	md_property_get(class_name, member_name, arg_type), \
	md_property_set(class_name, member_name, arg_type)