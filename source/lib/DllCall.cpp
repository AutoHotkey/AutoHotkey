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
#include "script.h"
#include "globaldata.h"
#include "script_func_impl.h"



#ifdef ENABLE_DLLCALL

#ifdef WIN32_PLATFORM
// Interface for DynaCall():
#define  DC_MICROSOFT           0x0000      // Default
#define  DC_BORLAND             0x0001      // Borland compat
#define  DC_CALL_CDECL          0x0010      // __cdecl
#define  DC_CALL_STD            0x0020      // __stdcall
#define  DC_RETVAL_MATH4        0x0100      // Return value in ST
#define  DC_RETVAL_MATH8        0x0200      // Return value in ST

#define  DC_CALL_STD_BO         (DC_CALL_STD | DC_BORLAND)
#define  DC_CALL_STD_MS         (DC_CALL_STD | DC_MICROSOFT)
#define  DC_CALL_STD_M8         (DC_CALL_STD | DC_RETVAL_MATH8)
#endif

union DYNARESULT                // Various result types
{      
    int     Int;                // Generic four-byte type
    long    Long;               // Four-byte long
    void   *Pointer;            // 32-bit pointer
    float   Float;              // Four byte real
    double  Double;             // 8-byte real
    __int64 Int64;              // big int (64-bit)
	UINT_PTR UIntPtr;
};

struct DYNAPARM
{
    union
	{
		int value_int; // Args whose width is less than 32-bit are also put in here because they are right justified within a 32-bit block on the stack.
		float value_float;
		__int64 value_int64;
		UINT_PTR value_uintptr;
		double value_double;
		char *astr;
		wchar_t *wstr;
		void *ptr;
    };
	// Might help reduce struct size to keep other members last and adjacent to each other (due to
	// 8-byte alignment caused by the presence of double and __int64 members in the union above).
	DllArgTypes type;
	union
	{
		int struct_size;
		struct
		{
			bool passed_by_address;
			bool is_unsigned; // Allows return value and output parameters to be interpreted as unsigned vs. signed.
			bool is_hresult; // Only used for the return value.
		};
	};
};

#ifdef _WIN64
// This function was borrowed from http://dyncall.org/
extern "C" UINT_PTR PerformDynaCall(size_t stackArgsSize, DWORD_PTR* stackArgs, DWORD_PTR* regArgs, void* aFunction);

// Retrieve a float or double return value.  These don't actually do anything, since the value we
// want is already in the xmm0 register which is used to return float or double values.
// Many thanks to http://locklessinc.com/articles/c_abi_hacks/ for the original idea.
extern "C" float GetFloatRetval();
extern "C" double GetDoubleRetval();

static inline UINT_PTR DynaParamToElement(DYNAPARM& parm)
{
	if(parm.passed_by_address)
		return (UINT_PTR) &parm.value_uintptr;
	else
		return parm.value_uintptr;
}
#endif

#ifdef WIN32_PLATFORM
DYNARESULT DynaCall(int aFlags, void *aFunction, DYNAPARM aParam[], int aParamCount, DWORD &aException
	, void *&aRet, int aRetSize, DWORD aExtraStackSize = 0)
#elif defined(_WIN64)
DYNARESULT DynaCall(void *aFunction, DYNAPARM aParam[], int aParamCount, DWORD &aException, void*& aRet, int aRetSize)
#else
#error DllCall not supported on this platform
#endif
// Based on the code by Ton Plooy <tonp@xs4all.nl>.
// Call the specified function with the given parameters. Build a proper stack and take care of correct
// return value processing.
{
	aException = 0;  // Set default output parameter for caller.
	SetLastError(g->LastError); // v1.0.46.07: In case the function about to be called doesn't change last-error, this line serves to retain the script's previous last-error rather than some arbitrary one produced by AutoHotkey's own internal API calls.  This line has no measurable impact on performance.

    DYNARESULT Res = {0}; // This struct is to be returned to caller by value.

#ifdef WIN32_PLATFORM

	// Declaring all variables early should help minimize stack interference of C code with asm.
	DWORD *our_stack;
	// Used to read the structure
	DWORD *pdword;
	DWORD esp_start, esp_end; // , dwEAX, dwEDX;
	int i, esp_delta; // Declare this here rather than later to prevent C code from interfering with esp.

	// Reserve enough space on the stack to handle the worst case of our args (which is currently a
	// maximum of 8 bytes per arg). This avoids any chance that compiler-generated code will use
	// the stack in a way that disrupts our insertion of args onto the stack.
	DWORD reserved_stack_size = aParamCount * 8 + aExtraStackSize;
	_asm
	{
		mov our_stack, esp  // our_stack is the location where we will write our args (bypassing "push").
		mov esp_start, esp  // For detecting whether a DC_CALL_STD function was sent too many or too few args.
		sub esp, reserved_stack_size  // The stack grows downward, so this "allocates" space on the stack.
	}

	// "Push" args onto the portion of the stack reserved above. Every argument is aligned on a 4-byte boundary.
	// We start at the rightmost argument (i.e. reverse order).
	for (i = aParamCount - 1; i > -1; --i)
	{
		DYNAPARM &this_param = aParam[i]; // For performance and convenience.
		if (this_param.type == DLL_ARG_STRUCT)
		{
			int& size = this_param.struct_size;
			ASSERT(!(size % 4));
			for (pdword = (DWORD*)(this_param.value_uintptr + size); size; size -= 4)
				*--our_stack = *--pdword;
		}
		// Push the arg or its address onto the portion of the stack that was reserved for our use above.
		else if (this_param.passed_by_address)
			*--our_stack = (DWORD)&this_param.value_int; // Any union member would work.
		else // this_param's value is contained directly inside the union.
		{
			if (this_param.type == DLL_ARG_INT64 || this_param.type == DLL_ARG_DOUBLE)
			{
				our_stack -= 2;
				*(__int64*)our_stack = this_param.value_int64;
			}
			else
				*--our_stack = this_param.value_uintptr;
		}
    }

	if (aRet)
	{
		if (aRetSize == 1 || aRetSize == 2 || aRetSize == 4 || aRetSize == 8)
			aRet = NULL;
		else
		{
			// Return value isn't passed through registers, memory copy
			// is performed instead. Pass the pointer as hidden arg.
			*--our_stack = (DWORD)aRet;	// ESP = ESP - 4, SS:[ESP] = pMem
		}
	}

	// Call the function.
	__try // Each try/except section adds at most 240 bytes of uncompressed code, and typically doesn't measurably affect performance.
	{
		_asm
		{
			mov esp, our_stack	   // Adjust ESP to indicate that the args have already been pushed onto the stack.
			call[aFunction]        // Stack is now properly built, we can call the function
		}
	}
	__except(EXCEPTION_EXECUTE_HANDLER)
	{
		aException = GetExceptionCode(); // aException is an output parameter for our caller.
	}

	// Even if an exception occurred (perhaps due to the callee having been passed a bad pointer),
	// attempt to restore the stack to prevent making things even worse.
	_asm
	{
		mov esp_end, esp        // See below.
		mov esp, esp_start      //
		// For DC_CALL_STD functions (since they pop their own arguments off the stack):
		// Since the stack grows downward in memory, if the value of esp after the call is less than
		// that before the call's args were pushed onto the stack, there are still items left over on
		// the stack, meaning that too many args (or an arg too large) were passed to the callee.
		// Conversely, if esp is now greater that it should be, too many args were popped off the
		// stack by the callee, meaning that too few args were provided to it.  In either case,
		// and even for CDECL, the following line restores esp to what it was before we pushed the
		// function's args onto the stack, which in the case of DC_CALL_STD helps prevent crashes
		// due to too many or to few args having been passed.
	}

	// Possibly adjust stack and read return values.
	// The following is commented out because the stack (esp) is restored above, for both CDECL and STD.
	//if (aFlags & DC_CALL_CDECL)
	//	_asm add esp, our_stack_size    // CDECL requires us to restore the stack after the call.
	if (aFlags & DC_RETVAL_MATH4)
		_asm fstp dword ptr [Res]
	else if (aFlags & DC_RETVAL_MATH8)
		_asm fstp qword ptr [Res]
	else if (!aRet)
	{
		_asm
		{
			mov  DWORD PTR [Res], eax
			mov  DWORD PTR [Res + 4], edx
		}
	}

#endif // WIN32_PLATFORM
#ifdef _WIN64

	int params_left = aParamCount, i = 0;
	DWORD_PTR regArgs[4];
	DWORD_PTR* stackArgs = NULL;
	size_t stackArgsSize = 0;

	if (aRet)
	{
		if (aRetSize == 1 || aRetSize == 2 || aRetSize == 4 || aRetSize == 8)
			aRet = NULL;
		else
		{
			// Return value isn't passed through registers, memory copy
			// is performed instead. Pass the pointer as hidden arg.
			regArgs[i++] = (DWORD_PTR)aRet;
		}
	}
	// The first four parameters are passed in x64 through registers... like ARM :D
	for(; (i < 4) && params_left; i++, params_left--)
		regArgs[i] = DynaParamToElement(aParam[i]);

	// Copy the remaining parameters
	if(params_left)
	{
		stackArgsSize = params_left * 8;
		stackArgs = (DWORD_PTR*) _alloca(stackArgsSize);

		for(int i = 0; i < params_left; i ++)
			stackArgs[i] = DynaParamToElement(aParam[i+4]);
	}

	// Call the function.
	__try
	{
		Res.UIntPtr = PerformDynaCall(stackArgsSize, stackArgs, regArgs, aFunction);
	}
	__except(EXCEPTION_EXECUTE_HANDLER)
	{
		aException = GetExceptionCode(); // aException is an output parameter for our caller.
	}

#endif

	// v1.0.42.03: The following supports A_LastError. It's called even if an exception occurred because it
	// might add value in some such cases.  Benchmarks show that this has no measurable impact on performance.
	// A_LastError was implemented rather than trying to change things so that a script could use DllCall to
	// call GetLastError() because: Even if we could avoid calling any API function that resets LastError
	// (which seems unlikely) it would be difficult to maintain (and thus a source of bugs) as revisions are
	// made in the future.
	g->LastError = GetLastError();

	TCHAR buf[32];

#ifdef WIN32_PLATFORM
	esp_delta = esp_start - esp_end; // Positive number means too many args were passed, negative means too few.
	if (esp_delta && (aFlags & DC_CALL_STD))
	{
		_itot(esp_delta, buf, 10);
		if (esp_delta > 0)
			g_script.ThrowRuntimeException(_T("Parameter list too large, or call requires CDecl."), buf);
		else
			g_script.ThrowRuntimeException(_T("Parameter list too small."), buf);
	}
	else
#endif
	// Too many or too few args takes precedence over reporting the exception because it's more informative.
	// In other words, any exception was likely caused by the fact that there were too many or too few.
	if (aException)
	{
		// It's a little easier to recognize the common error codes when they're in hex format.
		buf[0] = '0';
		buf[1] = 'x';
		_ultot(aException, buf + 2, 16);
		g_script.ThrowRuntimeException(ERR_EXCEPTION, buf);
	}

	return Res;
}



void ConvertDllArgType(LPTSTR aBuf, DYNAPARM &aDynaParam)
// Helper function for DllCall().  Updates aDynaParam's type and other attributes.
{
	LPTSTR type_string = aBuf;
	TCHAR buf[32];
	
	if (ctoupper(*type_string) == 'U') // Unsigned
	{
		aDynaParam.is_unsigned = true;
		++type_string; // Omit the 'U' prefix from further consideration.
	}
	else
		aDynaParam.is_unsigned = false;
	
	// Check for empty string before checking for pointer suffix, so that we can skip the first character.  This is needed to simplify "Ptr" type-name support.
	if (!*type_string)
	{
		aDynaParam.type = DLL_ARG_INVALID; 
		return; 
	}

	tcslcpy(buf, type_string, _countof(buf)); // Make a modifiable copy for easier parsing below.

	// v1.0.30.02: The addition of 'P'.
	// However, the current detection below relies upon the fact that none of the types currently
	// contain the letter P anywhere in them, so it would have to be altered if that ever changes.
	LPTSTR cp = StrChrAny(buf + 1, _T("*pP")); // Asterisk or the letter P.  Relies on the check above to ensure type_string is not empty (and buf + 1 is valid).
	if (cp && !*omit_leading_whitespace(cp + 1)) // Additional validation: ensure nothing following the suffix.
	{
		aDynaParam.passed_by_address = true;
		// Remove trailing options so that stricmp() can be used below.
		// Allow optional space in front of asterisk (seems okay even for 'P').
		if (IS_SPACE_OR_TAB(cp[-1]))
		{
			cp = omit_trailing_whitespace(buf, cp - 1);
			cp[1] = '\0'; // Terminate at the leftmost whitespace to remove all whitespace and the suffix.
		}
		else
			*cp = '\0'; // Terminate at the suffix to remove it.
	}
	else
		aDynaParam.passed_by_address = false;
	// Using a switch benchmarks better than the old approach, which was an if-else-if ladder
	// of string comparisons.  The old approach appeared to penalize Int64 vs. Int, perhaps due
	// to position in the ladder.
	switch (ctolower(*buf))
	{
	case 'i':
		// This method currently benchmarks better than _tcsnicmp, especially for Int64.
		if (ctolower(buf[1]) == 'n' && ctolower(buf[2]) == 't')
		{
			if (!buf[3])
				aDynaParam.type = DLL_ARG_INT;
			else if (buf[3] == '6' && buf[4] == '4' && !buf[5])
				aDynaParam.type = DLL_ARG_INT64;
			else
				break;
			return;
		}
		break;
	case 'p': if (!_tcsicmp(buf, _T("Ptr")))	{ aDynaParam.type = Exp32or64(DLL_ARG_INT, DLL_ARG_INT64); return; } break;
	case 's': if (!_tcsicmp(buf, _T("Str")))	{ aDynaParam.type = DLL_ARG_STR; return; }
			  if (!_tcsicmp(buf, _T("Short")))	{ aDynaParam.type = DLL_ARG_SHORT; return; } break;
	case 'd': if (!_tcsicmp(buf, _T("Double")))	{ aDynaParam.type = DLL_ARG_DOUBLE; return; } break;
	case 'f': if (!_tcsicmp(buf, _T("Float")))	{ aDynaParam.type = DLL_ARG_FLOAT; return; } break;
	case 'a': if (!_tcsicmp(buf, _T("AStr")))	{ aDynaParam.type = DLL_ARG_ASTR; return; } break;
	case 'w': if (!_tcsicmp(buf, _T("WStr")))	{ aDynaParam.type = DLL_ARG_WSTR; return; } break;
	case 'c': if (!_tcsicmp(buf, _T("Char")))	{ aDynaParam.type = DLL_ARG_CHAR; return; } break;
	}
	int b = 0;
	if (!aDynaParam.passed_by_address && !aDynaParam.is_unsigned && (isdigit(buf[b]) || !_tcsnicmp(buf, _T("Struct"), b = 6) && isdigit(buf[b]))) {
		int i = b + 1, n = buf[b] - '0';
		while (isdigit(buf[i])) n = n * 10 + (buf[i++] - '0');
		aDynaParam.struct_size = n;
		aDynaParam.type = n && buf[i] == 0 ? DLL_ARG_STRUCT : DLL_ARG_INVALID;
	}
	else aDynaParam.type = DLL_ARG_INVALID;
}


void *GetDllProcAddress(LPCTSTR aDllFileFunc, HMODULE *hmodule_to_free) // L31: Contains code extracted from BIF_DllCall for reuse in ExpressionToPostfix.
{
	int i;
	void *function = NULL;
	TCHAR param1_buf[MAX_PATH*2], *_tfunction_name, *dll_name; // Must use MAX_PATH*2 because the function name is INSIDE the Dll file, and thus MAX_PATH can be exceeded.
#ifndef UNICODE
	char *function_name;
#endif

	// Define the standard libraries here. If they reside in %SYSTEMROOT%\system32 it is not
	// necessary to specify the full path (it wouldn't make sense anyway).
	static HMODULE sStdModule[] = {GetModuleHandle(_T("user32")), GetModuleHandle(_T("kernel32"))
		, GetModuleHandle(_T("comctl32")), GetModuleHandle(_T("gdi32"))}; // user32 is listed first for performance.
	static const int sStdModule_count = _countof(sStdModule);

	// Make a modifiable copy of param1 so that the DLL name and function name can be parsed out easily, and so that "A" or "W" can be appended if necessary (e.g. MessageBoxA):
	tcslcpy(param1_buf, aDllFileFunc, _countof(param1_buf) - 1); // -1 to reserve space for the "A" or "W" suffix later below.
	if (   !(_tfunction_name = _tcsrchr(param1_buf, '\\'))   ) // No DLL name specified, so a search among standard defaults will be done.
	{
		dll_name = NULL;
#ifdef UNICODE
		char function_name[MAX_PATH];
		WideCharToMultiByte(CP_ACP, 0, param1_buf, -1, function_name, _countof(function_name), NULL, NULL);
#else
		function_name = param1_buf;
#endif

		// Since no DLL was specified, search for the specified function among the standard modules.
		for (i = 0; i < sStdModule_count; ++i)
			if (   sStdModule[i] && (function = (void *)GetProcAddress(sStdModule[i], function_name))   )
				break;
		if (!function)
		{
			// Since the absence of the "A" suffix (e.g. MessageBoxA) is so common, try it that way
			// but only here with the standard libraries since the risk of ambiguity (calling the wrong
			// function) seems unacceptably high in a custom DLL.  For example, a custom DLL might have
			// function called "AA" but not one called "A".
			strcat(function_name, WINAPI_SUFFIX); // 1 byte of memory was already reserved above for the 'A'.
			for (i = 0; i < sStdModule_count; ++i)
				if (   sStdModule[i] && (function = (void *)GetProcAddress(sStdModule[i], function_name))   )
					break;
		}
	}
	else // DLL file name is explicitly present.
	{
		dll_name = param1_buf;
		*_tfunction_name = '\0';  // Terminate dll_name to split it off from function_name.
		++_tfunction_name; // Set it to the character after the last backslash.
#ifdef UNICODE
		char function_name[MAX_PATH];
		WideCharToMultiByte(CP_ACP, 0, _tfunction_name, -1, function_name, _countof(function_name), NULL, NULL);
#else
		function_name = _tfunction_name;
#endif

		// Get module handle. This will work when DLL is already loaded and might improve performance if
		// LoadLibrary is a high-overhead call even when the library already being loaded.  If
		// GetModuleHandle() fails, fall back to LoadLibrary().
		HMODULE hmodule;
		if (   !(hmodule = GetModuleHandle(dll_name))    )
			if (   !hmodule_to_free  ||  !(hmodule = *hmodule_to_free = LoadLibrary(dll_name))   )
			{
				if (hmodule_to_free) // L31: BIF_DllCall wants us to set throw.  ExpressionToPostfix passes NULL.
					g_script.ThrowRuntimeException(_T("Failed to load DLL."), dll_name);
				return NULL;
			}
		if (   !(function = (void *)GetProcAddress(hmodule, function_name))   )
		{
			// v1.0.34: If it's one of the standard libraries, try the "A" suffix.
			// jackieku: Try it anyway, there are many other DLLs that use this naming scheme, and it doesn't seem expensive.
			// If an user really cares he or she can always work around it by editing the script.
			//for (i = 0; i < sStdModule_count; ++i)
			//	if (hmodule == sStdModule[i]) // Match found.
			//	{
					strcat(function_name, WINAPI_SUFFIX); // 1 byte of memory was already reserved above for the 'A'.
					function = (void *)GetProcAddress(hmodule, function_name);
			//		break;
			//	}
		}
	}

	if (!function && hmodule_to_free) // Caller wants us to throw.
	{
		// This must be done here since only we know for certain that the dll
		// was loaded okay (if GetModuleHandle succeeded, nothing is passed
		// back to the caller).
		g_script.ThrowRuntimeException(ERR_NONEXISTENT_FUNCTION, _tfunction_name);
	}

	return function;
}


void GetDllArgObjectPtr(ResultToken &aResultToken, IObject *obj, size_t &aPtr)
{
	if (BufferObject::IsInstanceExact(obj))
	{
		aPtr = (size_t)((BufferObject *)obj)->Data();
		return;
	}
	auto result = GetObjectPtrProperty(obj, _T("Ptr"), aPtr, aResultToken, obj->IsOfType(Object::sPrototype));
	if (result != INVOKE_NOT_HANDLED)
		return;
	Object *b = (Object *)obj;
	if (b->GetDataPtr(aPtr) != OK)
		aResultToken.UnknownMemberError(ExprTokenType(obj), IT_GET, _T("Ptr"));
}


BIF_DECL(BIF_DllCall)
// Stores a number or a SYM_STRING result in aResultToken.
// Caller has set up aParam to be viewable as a left-to-right array of params rather than a stack.
// It has also ensured that the array has exactly aParamCount items in it.
// Author: Marcus Sonntag (Ultra)
{
	LPTSTR function_name = NULL;
	void *function = NULL; // Will hold the address of the function to be called.
	int vf_index = -1; // Set default: not ComCall.

	if (_f_callee_id == FID_ComCall)
	{
		function = NULL;
		if (!ParamIndexIsNumeric(0))
			_f_throw_param(0, _T("Integer"));
		vf_index = (int)ParamIndexToInt64(0);
		if (vf_index < 0) // But positive values aren't checked since there's no known upper bound.
			_f_throw_param(0);
		// Cheat a bit to make the second arg both the source of the virtual function
		// and the first parameter value (always an interface pointer):
		static ExprTokenType t_this_arg_type = _T("Ptr");
		aParam[0] = &t_this_arg_type;
	}
	else
	{
		// Check that the mandatory first parameter (DLL+Function) is valid.
		// (load-time validation has ensured at least one parameter is present).
		switch (TypeOfToken(*aParam[0]))
		{
		case SYM_INTEGER: // Might be the most common case, due to FinalizeExpression resolving function names at load time.
			// v1.0.46.08: Allow script to specify the address of a function, which might be useful for
			// calling functions that the script discovers through unusual means such as C++ member functions.
			function = (void *)ParamIndexToInt64(0);
			// A check like the following is not present due to rarity of need and because if the address
			// is zero or negative, the same result will occur as for any other invalid address:
			// an exception code of 0xc0000005.
			//if ((UINT64)temp64 < 0x10000 || (UINT64)temp64 > UINTPTR_MAX)
			//	_f_throw_param(0); // Stage 1 error: Invalid first param.
			//// Otherwise, assume it's a valid address:
			//	function = (void *)temp64;
			break;
		case SYM_STRING: // For performance, don't even consider the possibility that a string like "33" is a function-address.
			//function = NULL; // Already set: indicates that no function has been specified yet.
			break;
		case SYM_OBJECT:
			// Permit an object with Ptr property.  This enables DllCall or DllCall.Bind() to be used directly
			// as a method of an object, such as one used for wrapping a dll function.  It could also have other
			// uses, such as resolving and memoizing function addresses on first use.
			__int64 n;
			if (!GetObjectIntProperty(ParamIndexToObject(0), _T("Ptr"), n, aResultToken))
				return;
			function = (void *)n;
			break;
		default: // SYM_FLOAT, SYM_MISSING or (should be impossible) something else.
			_f_throw(ERR_PARAM1_INVALID, ErrorPrototype::Type);
		}
		if (!function)
			function_name = TokenToString(*aParam[0]);
		++aParam; // Normalize aParam to simplify ComCall vs. DllCall.
		--aParamCount;
	}

	// Determine the type of return value.
	DYNAPARM return_attrib = {0}; // Init all to default in case ConvertDllArgType() isn't called below. This struct holds the type and other attributes of the function's return value.
	void* return_struct_ptr = nullptr;
	int return_struct_size = 0;
#ifdef WIN32_PLATFORM
	int dll_call_mode = DC_CALL_STD; // Set default.  Can be overridden to DC_CALL_CDECL and flags can be OR'd into it.
	int struct_extra_size = 0;
#endif
	
#ifdef UNICODE
#define UorA_(u, a) u
#else
#define UorA_(u, a) a
#endif // UNICODE
	// for Unicode <-> ANSI charset conversion
	UorA_(CStringA**, CStringW**) pStr;
	struct _AutoFree {
		HMODULE hmodule_to_free;
		BufferObject* return_struct_obj;
		UorA_(CStringA**, CStringW**) ptr;
		int len;
		~_AutoFree() {
			for (int i = 0; i < len; i++)
				if (ptr[i])
					delete ptr[i];
			if (return_struct_obj)
				return_struct_obj->Release();
			if (hmodule_to_free)
				FreeLibrary(hmodule_to_free);
		}
	} free_after_exit{0};	// Avoid memory leaks when _f_throw_xxx
#undef UorA_

	if ( !(aParamCount % 2) ) // An even number of parameters indicates the return type has been omitted. aParamCount excludes DllCall's first parameter at this point.
	{
		return_attrib.type = DLL_ARG_INT;
		if (vf_index >= 0) // Default to HRESULT for ComCall.
			return_attrib.is_hresult = true;
		// Otherwise, assume normal INT (also covers BOOL).
	}
	else
	{
		// Check validity of this arg's return type:
		ExprTokenType &token = *aParam[aParamCount - 1];
		LPTSTR return_type_string = TokenToString(token); // If non-numeric it will return "", which is detected as invalid below.

		// 64-bit note: The calling convention detection code is preserved here for script compatibility.

		if (!_tcsnicmp(return_type_string, _T("CDecl"), 5)) // Alternate calling convention.
		{
#ifdef WIN32_PLATFORM
			dll_call_mode = DC_CALL_CDECL;
#endif
			return_type_string = omit_leading_whitespace(return_type_string + 5);
			if (!*return_type_string)
			{	// Take a shortcut since we know this empty string will be used as "Int":
				return_attrib.type = DLL_ARG_INT;
				goto has_valid_return_type;
			}
		}
		if (!_tcsicmp(return_type_string, _T("HRESULT")))
		{
			return_attrib.type = DLL_ARG_INT;
			return_attrib.is_hresult = true;
			//return_attrib.is_unsigned = true; // Not relevant since an exception is thrown for any negative value.
		}
		else if (TokenIsPureNumeric(token) == SYM_INTEGER)
		{
			__int64 size = TokenToInt64(token);
			// See the arg processing section below for comments.
			return_attrib.type = (size >= 1 && size < 65536) ? DLL_ARG_STRUCT : DLL_ARG_INVALID;
			return_struct_size = (int)size;
			return_type_string = nullptr; // To help detect errors.
		}
		else
		{
			ConvertDllArgType(return_type_string, return_attrib);
			if (return_attrib.type == DLL_ARG_STRUCT)
				return_struct_size = return_attrib.struct_size, return_attrib.struct_size = 0;
		}
		if (return_attrib.type == DLL_ARG_INVALID)
			_f_throw_value(ERR_INVALID_RETURN_TYPE);

has_valid_return_type:
		--aParamCount;  // Remove the last parameter from further consideration.
#ifdef WIN32_PLATFORM
		if (!return_attrib.passed_by_address) // i.e. the special return flags below are not needed when an address is being returned.
		{
			if (return_attrib.type == DLL_ARG_DOUBLE)
				dll_call_mode |= DC_RETVAL_MATH8;
			else if (return_attrib.type == DLL_ARG_FLOAT)
				dll_call_mode |= DC_RETVAL_MATH4;
		}
#endif
	}

	if (return_struct_size) {
		if (return_struct_ptr = malloc(return_struct_size))
			free_after_exit.return_struct_obj = BufferObject::Create(return_struct_ptr, return_struct_size);
		else
			_f_throw_oom;
	}

	// Using stack memory, create an array of dll args large enough to hold the actual number of args present.
	int arg_count = aParamCount/2;
	DYNAPARM *dyna_param = arg_count ? (DYNAPARM *)_alloca(arg_count * sizeof(DYNAPARM)) : NULL;
	// Above: _alloca() has been checked for code-bloat and it doesn't appear to be an issue.
	// Above: Fix for v1.0.36.07: According to MSDN, on failure, this implementation of _alloca() generates a
	// stack overflow exception rather than returning a NULL value.  Therefore, NULL is no longer checked,
	// nor is an exception block used since stack overflow in this case should be exceptionally rare (if it
	// does happen, it would probably mean the script or the program has a design flaw somewhere, such as
	// infinite recursion).

	LPTSTR arg_type_string;
	int i = arg_count * sizeof(void *);
	
	pStr = UorA(CStringA**, CStringW**)_alloca(i); // _alloca vs malloc can make a significant difference to performance in some cases.
	memset(pStr, 0, i);
	free_after_exit.ptr = pStr;
	free_after_exit.len = arg_count;

	// Above has already ensured that after the first parameter, there are either zero additional parameters
	// or an even number of them.  In other words, each arg type will have an arg value to go with it.
	// It has also verified that the dyna_param array is large enough to hold all of the args.
	for (arg_count = 0, i = 0; i < aParamCount; ++arg_count, i += 2)  // Same loop as used later below, so maintain them together.
	{
		// Store each arg into a dyna_param struct, using its arg type to determine how.
		DYNAPARM &this_dyna_param = dyna_param[arg_count];

		if (TokenIsPureNumeric(*aParam[i]) == SYM_INTEGER)
		{
			__int64 size = TokenToInt64(*aParam[i]);
			// Prohibiting size >= 64KB should detect any cases where a pointer value is unintentionally
			// being passed as an arg type, while not actually imposing a hard limit since larger numbers
			// can be passed as strings.  Note that the stack itself has a limit of 1-8MB.
			this_dyna_param.type = (size >= 1 && size < 65536) ? DLL_ARG_STRUCT : DLL_ARG_INVALID;
			this_dyna_param.struct_size = (int)size;
			arg_type_string = nullptr; // To help detect errors.
		}
		else
		{
			arg_type_string = TokenToString(*aParam[i]); // aBuf not needed since floating-point and "" are equally invalid.
			ConvertDllArgType(arg_type_string, this_dyna_param);
		}
		if (this_dyna_param.type == DLL_ARG_INVALID)
			_f_throw_value(ERR_INVALID_ARG_TYPE);

		IObject *this_param_obj = TokenToObject(*aParam[i + 1]);
		if (this_param_obj && this_dyna_param.type != DLL_ARG_STRUCT)
		{
			if ((this_dyna_param.passed_by_address || this_dyna_param.type == DLL_ARG_STR)
				&& dynamic_cast<VarRef*>(this_param_obj))
			{
				aParam[i + 1] = (ExprTokenType *)_alloca(sizeof(ExprTokenType));
				aParam[i + 1]->SetVarRef(static_cast<VarRef*>(this_param_obj));
				this_param_obj = nullptr;
			}
			else if (ctoupper(*arg_type_string) == 'P')
			{
				// Support Buffer.Ptr, but only for "Ptr" type.  All other types are reserved for possible
				// future use, which might be general like obj.ToValue(), or might be specific to DllCall
				// or the particular type of this arg (Int, Float, etc.).
				GetDllArgObjectPtr(aResultToken, this_param_obj, this_dyna_param.value_uintptr);
				if (aResultToken.Exited())
					return;
				continue;
			}
		}
		ExprTokenType &this_param = *aParam[i + 1];
		if (this_param.symbol == SYM_MISSING)
			_f_throw(ERR_PARAM_REQUIRED);

		switch (this_dyna_param.type)
		{
		case DLL_ARG_STR:
			if (IS_NUMERIC(this_param.symbol) || this_param_obj)
			{
				// For now, string args must be real strings rather than floats or ints.  An alternative
				// to this would be to convert it to number using persistent memory from the caller (which
				// is necessary because our own stack memory should not be passed to any function since
				// that might cause it to return a pointer to stack memory, or update an output-parameter
				// to be stack memory, which would be invalid memory upon return to the caller).
				// The complexity of this doesn't seem worth the rarity of the need, so this will be
				// documented in the help file.
				_f_throw_type(_T("String"), this_param);
			}
			// Otherwise, it's a supported type of string.
			this_dyna_param.ptr = TokenToString(this_param); // SYM_VAR's Type() is always VAR_NORMAL (except lvalues in expressions).
			// NOTES ABOUT THE ABOVE:
			// UPDATE: The v1.0.44.14 item below doesn't work in release mode, only debug mode (turning off
			// "string pooling" doesn't help either).  So it's commented out until a way is found
			// to pass the address of a read-only empty string (if such a thing is possible in
			// release mode).  Such a string should have the following properties:
			// 1) The first byte at its address should be '\0' so that functions can read it
			//    and recognize it as a valid empty string.
			// 2) The memory address should be readable but not writable: it should throw an
			//    access violation if the function tries to write to it (like "" does in debug mode).
			// SO INSTEAD of the following, DllCall() now checks further below for whether sEmptyString
			// has been overwritten/trashed by the call, and if so displays a warning dialog.
			// See note above about this: v1.0.44.14: If a variable is being passed that has no capacity, pass a
			// read-only memory area instead of a writable empty string. There are two big benefits to this:
			// 1) It forces an immediate exception (catchable by DllCall's exception handler) so
			//    that the program doesn't crash from memory corruption later on.
			// 2) It avoids corrupting the program's static memory area (because sEmptyString
			//    resides there), which can save many hours of debugging for users when the program
			//    crashes on some seemingly unrelated line.
			// Of course, it's not a complete solution because it doesn't stop a script from
			// passing a variable whose capacity is non-zero yet too small to handle what the
			// function will write to it.  But it's a far cry better than nothing because it's
			// common for a script to (unintentionally) pass an empty/uninitialized variable to
			// some function that writes a string to it.
			//if (this_dyna_param.str == Var::sEmptyString) // To improve performance, compare directly to Var::sEmptyString rather than calling Capacity().
			//	this_dyna_param.str = _T(""); // Make it read-only to force an exception.  See comments above.
			break;
		case DLL_ARG_xSTR:
			// See the section above for comments.
			if (IS_NUMERIC(this_param.symbol) || this_param_obj)
				_f_throw_type(_T("String"), this_param);
			// String needing translation: ASTR on Unicode build, WSTR on ANSI build.
			pStr[arg_count] = new UorA(CStringCharFromWChar,CStringWCharFromChar)(TokenToString(this_param));
			this_dyna_param.ptr = pStr[arg_count]->GetBuffer();
			break;

		case DLL_ARG_DOUBLE:
		case DLL_ARG_FLOAT:
			// This currently doesn't validate that this_dyna_param.is_unsigned==false, since it seems
			// too rare and mostly harmless to worry about something like "Ufloat" having been specified.
			if (!TokenIsNumeric(this_param))
				_f_throw_type(_T("Number"), this_param);
			this_dyna_param.value_double = TokenToDouble(this_param);
			if (this_dyna_param.type == DLL_ARG_FLOAT)
				this_dyna_param.value_float = (float)this_dyna_param.value_double;
			break;
			
		case DLL_ARG_STRUCT: {
			if (this_param_obj) {
				GetDllArgObjectPtr(aResultToken, this_param_obj, this_dyna_param.value_uintptr);
				if (aResultToken.Exited())
					return;
			}
			else if (TokenIsPureNumeric(this_param) == SYM_INTEGER)
				this_dyna_param.value_int64 = TokenToInt64(this_param);
			else
				_o_throw_type(_T("Integer"), this_param);
			if (this_dyna_param.value_uintptr < 65536)
				_o_throw_value(ERR_INVALID_VALUE);

			int& size = this_dyna_param.struct_size;
#ifdef _WIN64
			this_dyna_param.type = DLL_ARG_INT64;
			if (size <= 8)
				this_dyna_param.value_int64 = *(__int64*)this_dyna_param.ptr;
			size = 0;
#else
			size = (size + sizeof(void*) - 1) & -(int)sizeof(void*);
			if (size > 8)
				struct_extra_size += size - 8;
			else if (size > 4) {
				this_dyna_param.type = DLL_ARG_INT64;
				this_dyna_param.value_int64 = *(__int64*)this_dyna_param.ptr;
				size = 0;
			}
			else {
				this_dyna_param.type = DLL_ARG_INT;
				this_dyna_param.value_int = *(int*)this_dyna_param.ptr;
				size = 0;
			}
#endif // _WIN64
			break;
		}

		default: // Namely:
		//case DLL_ARG_INT:
		//case DLL_ARG_SHORT:
		//case DLL_ARG_CHAR:
		//case DLL_ARG_INT64:
			if (!TokenIsNumeric(this_param))
				_f_throw_type(_T("Number"), this_param);
			// Note that since v2.0-a083-97803aeb, TokenToInt64 supports conversion of large unsigned 64-bit
			// numbers from strings (producing a negative value, but with the right bit representation).
			// This allows large unsigned literals and numeric strings to be passed to DllCall (regardless
			// of whether Int64 or UInt64 is used), but the script itself will interpret the value as signed
			// if greater than _I64_MAX.  Any UInt64 values returned by DllCall can be safely passed back
			// without loss, and can be operated upon by the bitwise operators, although arithmetic and
			// string conversion will treat the value as Int64.
			this_dyna_param.value_int64 = TokenToInt64(this_param);
		} // switch (this_dyna_param.type)
	} // for() each arg.
    
	if (vf_index >= 0) // ComCall
	{
		if ((UINT_PTR)dyna_param[0].ptr < 65536) // Basic sanity check to catch null pointers and small numbers.  On Win32, the first 64KB of address space is always invalid.
			_f_throw_param(1);
		LPVOID *vftbl = *(LPVOID **)dyna_param[0].ptr;
		function = vftbl[vf_index];
	}
	else if (!function) // The function's address hasn't yet been determined.
	{
		function = GetDllProcAddress(function_name, &free_after_exit.hmodule_to_free);
		if (!function)
		{
			// GetDllProcAddress has thrown the appropriate exception.
			aResultToken.SetExitResult(FAIL);
			return;
		}
	}

	////////////////////////
	// Call the DLL function
	////////////////////////
	DWORD exception_occurred; // Must not be named "exception_code" to avoid interfering with MSVC macros.
	DYNARESULT return_value;  // Doing assignment (below) as separate step avoids compiler warning about "goto end" skipping it.
#ifdef WIN32_PLATFORM
	return_value = DynaCall(dll_call_mode, function, dyna_param, arg_count, exception_occurred
		, return_struct_ptr, return_struct_size, struct_extra_size + (return_struct_ptr ? 4 : 0));
#endif
#ifdef _WIN64
	return_value = DynaCall(function, dyna_param, arg_count, exception_occurred, return_struct_ptr, return_struct_size);
#endif

	if (*Var::sEmptyString)
	{
		// v1.0.45.01 Above has detected that a variable of zero capacity was passed to the called function
		// and the function wrote to it (assuming sEmptyString wasn't already trashed some other way even
		// before the call).  So patch up the empty string to stabilize a little; but it's too late to
		// salvage this instance of the program because there's no knowing how much static data adjacent to
		// sEmptyString has been overwritten and corrupted.
		*Var::sEmptyString = '\0';
		// Don't bother with freeing hmodule_to_free since a critical error like this calls for minimal cleanup.
		// The OS almost certainly frees it upon termination anyway.
		// Call CriticalError() so that the user knows *which* DllCall is at fault:
		g_script.CriticalError(_T("An invalid write to an empty variable was detected."));
		// CriticalError always terminates the process.
	}

	if (g->ThrownToken || return_attrib.is_hresult && FAILED((HRESULT)return_value.Int))
	{
		if (!g->ThrownToken)
			// "Error values (as defined by the FAILED macro) are never returned"; so FAIL, not FAIL_OR_OK.
			g_script.Win32Error((DWORD)return_value.Int, FAIL);
		// If a script exception was thrown by DynaCall(), it was either because the called function threw
		// a SEH exception or because the stdcall parameter list was the wrong size.  In any of these cases,
		// set FAIL result to ensure control is transferred as expected (exiting the thread or TRY block).
		aResultToken.SetExitResult(FAIL);
		// But continue on to write out any output parameters because the called function might have
		// had a chance to update them before aborting.  They might be of some use in debugging the
		// issue, though the script would have to catch the exception to be able to inspect them.
	}
	else // The call was successful.  Interpret and store the return value.
	{
		// If the return value is passed by address, dereference it here.
		if (return_attrib.passed_by_address)
		{
			return_attrib.passed_by_address = false; // Because the address is about to be dereferenced/resolved.

			switch(return_attrib.type)
			{
			case DLL_ARG_INT64:
			case DLL_ARG_DOUBLE:
#ifdef _WIN64 // fincs: pointers are 64-bit on x64.
			case DLL_ARG_WSTR:
			case DLL_ARG_ASTR:
#endif
				// Same as next section but for eight bytes:
				return_value.Int64 = *(__int64 *)return_value.Pointer;
				break;
			default: // Namely:
			//case DLL_ARG_STR:  // Even strings can be passed by address, which is equivalent to "char **".
			//case DLL_ARG_INT:
			//case DLL_ARG_SHORT:
			//case DLL_ARG_CHAR:
			//case DLL_ARG_FLOAT:
				// All the above are stored in four bytes, so a straight dereference will copy the value
				// over unchanged, even if it's a float.
				return_value.Int = *(int *)return_value.Pointer;
			}
		}
#ifdef _WIN64
		else
		{
			switch(return_attrib.type)
			{
			// Floating-point values are returned via the xmm0 register. Copy it for use in the next section:
			case DLL_ARG_FLOAT:
				return_value.Float = GetFloatRetval();
				break;
			case DLL_ARG_DOUBLE:
				return_value.Double = GetDoubleRetval();
				break;
			}
		}
#endif

		switch(return_attrib.type)
		{
		case DLL_ARG_INT: // Listed first for performance. If the function has a void return value (formerly DLL_ARG_NONE), the value assigned here is undefined and inconsequential since the script should be designed to ignore it.
			ASSERT(aResultToken.symbol == SYM_INTEGER);
			if (return_attrib.is_unsigned)
				aResultToken.value_int64 = (UINT)return_value.Int; // Preserve unsigned nature upon promotion to signed 64-bit.
			else // Signed.
				aResultToken.value_int64 = return_value.Int;
			break;
		case DLL_ARG_STR:
			// The contents of the string returned from the function must not reside in our stack memory since
			// that will vanish when we return to our caller.  As long as every string that went into the
			// function isn't on our stack (which is the case), there should be no way for what comes out to be
			// on the stack either.
			aResultToken.symbol = SYM_STRING;
			aResultToken.marker = (LPTSTR)(return_value.Pointer ? return_value.Pointer : _T(""));
			// Above: Fix for v1.0.33.01: Don't allow marker to be set to NULL, which prevents crash
			// with something like the following, which in this case probably happens because the inner
			// call produces a non-numeric string, which "int" then sees as zero, which CharLower() then
			// sees as NULL, which causes CharLower to return NULL rather than a real string:
			//result := DllCall("CharLower", "int", DllCall("CharUpper", "str", MyVar, "str"), "str")
			break;
		case DLL_ARG_xSTR:
			{	// String needing translation: ASTR on Unicode build, WSTR on ANSI build.
#ifdef UNICODE
				LPCSTR result = (LPCSTR)return_value.Pointer;
#else
				LPCWSTR result = (LPCWSTR)return_value.Pointer;
#endif
				if (result && *result)
				{
#ifdef UNICODE		// Perform the translation:
					CStringWCharFromChar result_buf(result);
#else
					CStringCharFromWChar result_buf(result);
#endif
					// Store the length of the translated string first since DetachBuffer() clears it.
					aResultToken.marker_length = result_buf.GetLength();
					// Now attempt to take ownership of the malloc'd memory, to return to our caller.
					if (aResultToken.mem_to_free = result_buf.DetachBuffer())
						aResultToken.marker = aResultToken.mem_to_free;
					else
						aResultToken.marker = _T("");
				}
				else
					aResultToken.marker = _T("");
				aResultToken.symbol = SYM_STRING;
			}
			break;
		case DLL_ARG_SHORT:
			ASSERT(aResultToken.symbol == SYM_INTEGER);
			if (return_attrib.is_unsigned)
				aResultToken.value_int64 = return_value.Int & 0x0000FFFF; // This also forces the value into the unsigned domain of a signed int.
			else // Signed.
				aResultToken.value_int64 = (SHORT)(WORD)return_value.Int; // These casts properly preserve negatives.
			break;
		case DLL_ARG_CHAR:
			ASSERT(aResultToken.symbol == SYM_INTEGER);
			if (return_attrib.is_unsigned)
				aResultToken.value_int64 = return_value.Int & 0x000000FF; // This also forces the value into the unsigned domain of a signed int.
			else // Signed.
				aResultToken.value_int64 = (char)(BYTE)return_value.Int; // These casts properly preserve negatives.
			break;
		case DLL_ARG_INT64:
			// Even for unsigned 64-bit values, it seems best both for simplicity and consistency to write
			// them back out to the script as signed values because script internals are not currently
			// equipped to handle unsigned 64-bit values.  This has been documented.
			ASSERT(aResultToken.symbol == SYM_INTEGER);
			aResultToken.value_int64 = return_value.Int64;
			break;
		case DLL_ARG_FLOAT:
			aResultToken.symbol = SYM_FLOAT;
			aResultToken.value_double = return_value.Float;
			break;
		case DLL_ARG_DOUBLE:
			aResultToken.symbol = SYM_FLOAT; // There is no SYM_DOUBLE since all floats are stored as doubles.
			aResultToken.value_double = return_value.Double;
			break;
		case DLL_ARG_STRUCT:
			if (!return_struct_ptr)
				memcpy(free_after_exit.return_struct_obj->Data(), &return_value.Int64, return_struct_size);
			aResultToken.SetValue(free_after_exit.return_struct_obj);
			break;
		//default: // Should never be reached unless there's a bug.
		//	aResultToken.symbol = SYM_STRING;
		//	aResultToken.marker = "";
		} // switch(return_attrib.type)
	} // Storing the return value when no exception occurred.

	// Store any output parameters back into the input variables.  This allows a function to change the
	// contents of a variable for the following arg types: String and Pointer to <various number types>.
	for (arg_count = 0, i = 0; i < aParamCount; ++arg_count, i += 2) // Same loop as used above, so maintain them together.
	{
		ExprTokenType &this_param = *aParam[i + 1];  // Resolved for performance and convenience.
		DYNAPARM &this_dyna_param = dyna_param[arg_count];

		if (IObject * obj = TokenToObject(this_param)) // Implies the type is "Ptr" or "Ptr*".
		{
			if (this_dyna_param.passed_by_address)
				SetObjectIntProperty(obj, _T("Ptr"), this_dyna_param.value_int64, aResultToken);
			continue;
		}

		if (this_param.symbol != SYM_VAR)
			continue;
		Var &output_var = *this_param.var;

		if (!this_dyna_param.passed_by_address)
		{
			if (this_dyna_param.type == DLL_ARG_STR) // Native string type for current build config.
			{
				// Update the variable's length and check for null termination.  This could be skipped
				// when a naked variable (not VarRef) is passed since that's supposed to be input-only,
				// but seems better to do this unconditionally since the function can in fact modify
				// the variable's contents, and detecting buffer overrun errors seems more important
				// than any performance gain from skipping this.
				output_var.SetLengthFromContents();
				output_var.Close(); // Clear the attributes of the variable to reflect the fact that the contents may have changed.
			}
			// Nothing is done for xSTR since 1) we didn't pass the variable's contents to the function
			// so its length doesn't need updating, and 2) the buffer that was passed was only as large
			// as the input string, so has very little practical use for output.
			// No other types can be output parameters when !passed_by_address.
			continue;
		}
		if (VARREF_IS_READ(this_param.var_usage))
			continue; // Output parameters are copied back only if provided with a VarRef (&variable).

		switch (this_dyna_param.type)
		{
		case DLL_ARG_INT:
			if (this_dyna_param.is_unsigned)
				output_var.Assign((DWORD)this_dyna_param.value_int);
			else // Signed.
				output_var.Assign(this_dyna_param.value_int);
			break;
		case DLL_ARG_SHORT:
			if (this_dyna_param.is_unsigned) // Force omission of the high-order word in case it is non-zero from a parameter that was originally and erroneously larger than a short.
				output_var.Assign(this_dyna_param.value_int & 0x0000FFFF); // This also forces the value into the unsigned domain of a signed int.
			else // Signed.
				output_var.Assign((int)(SHORT)(WORD)this_dyna_param.value_int); // These casts properly preserve negatives.
			break;
		case DLL_ARG_CHAR:
			if (this_dyna_param.is_unsigned) // Force omission of the high-order bits in case it is non-zero from a parameter that was originally and erroneously larger than a char.
				output_var.Assign(this_dyna_param.value_int & 0x000000FF); // This also forces the value into the unsigned domain of a signed int.
			else // Signed.
				output_var.Assign((int)(char)(BYTE)this_dyna_param.value_int); // These casts properly preserve negatives.
			break;
		case DLL_ARG_INT64: // Unsigned and signed are both written as signed for the reasons described elsewhere above.
			output_var.Assign(this_dyna_param.value_int64);
			break;
		case DLL_ARG_FLOAT:
			output_var.Assign(this_dyna_param.value_float);
			break;
		case DLL_ARG_DOUBLE:
			output_var.Assign(this_dyna_param.value_double);
			break;
		case DLL_ARG_STR: // Str*
			// The use of LPWSTR* vs. LPWSTR typically means the function will pass back the
			// address of a string, not modify the string itself.  This is also consistent with
			// passed_by_address for all other types.  However, it must be used carefully since
			// there's no way for Str* to know how or whether the function requires the string
			// to be freed (e.g. by calling CoTaskMemFree()).
			if (this_dyna_param.ptr != output_var.Contents(FALSE)
				&& !output_var.AssignString((LPTSTR)this_dyna_param.ptr))
				aResultToken.SetExitResult(FAIL);
			break;
		case DLL_ARG_xSTR: // AStr* on Unicode builds and WStr* on ANSI builds.
			if (this_dyna_param.ptr != output_var.Contents(FALSE)
				&& !output_var.AssignStringFromCodePage(UorA(LPSTR,LPWSTR)this_dyna_param.ptr))
				aResultToken.SetExitResult(FAIL);
		}
	}

	if (aResultToken.symbol == SYM_OBJECT)
		free_after_exit.return_struct_obj = nullptr;	// Avoid buffer obj be released
}

#endif
