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
#include "application.h"
#include "script_func_impl.h"
#include "MdFunc.h"



#ifdef ENABLE_REGISTERCALLBACK

struct RCArgType
{
	MdType type;
	Object *proto;
};


struct RCCallbackFunc // Used by BIF_CallbackCreate() and related.
{
#ifdef WIN32_PLATFORM
	ULONG data1;	//E8 00 00 00
	ULONG data2;	//00 8D 44 24
	ULONG data3;	//08 50 FF 15
	UINT64 (CALLBACK **callfuncptr)(UINT_PTR*, char*);
	ULONG data4;	//59 84 C4 nn
	USHORT data5;	//FF E1
#endif
#ifdef _WIN64
	UINT64 data1; // 0xfffffffff9058d48
	UINT64 data2; // 0x9090900000000325
	void (*stub)();
	UINT64 (CALLBACK *callfuncptr)(UINT_PTR*, char*);
#endif
	//code ends
	UCHAR actual_param_count; // This is the actual (not formal) number of parameters passed from the caller to the callback. Kept adjacent to the USHORT above to conserve memory due to 4-byte struct alignment.
#define CBF_CREATE_NEW_THREAD	1
#define CBF_PASS_PARAMS_POINTER	2
#define CBF_HAS_ARG_TYPES		4
	UCHAR flags; // Kept adjacent to above to conserve memory due to 4-byte struct alignment in 32-bit builds.
	IObject *func; // The function object to be called whenever the callback's caller calls callfuncptr.
};


static float ReturnFloat_(float f) { return f; }
static double ReturnDouble_(double f) { return f; }
static void ReturnFloat(float f) { reinterpret_cast<void (*)(float)>(&ReturnFloat_)(f); }
static void ReturnDouble(double f) { reinterpret_cast<void (*)(double)>(&ReturnDouble_)(f); }


#ifdef _WIN64
extern "C" void RegisterCallbackAsmStub();
#endif


UINT64 CALLBACK RegisterCallbackCStub(UINT_PTR *params, char *address) // Used by BIF_RegisterCallback().
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
	auto arg_type = (cb.flags & CBF_HAS_ARG_TYPES) ? (RCArgType*)(&cb + 1) : nullptr;

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

	UINT64 number_to_return = 0;
	ExprTokenType **param, one_param, *one_param_ptr;
	int param_count = 0;
	int ret_size = 0;
	void *ret_ptr = nullptr;
	MdType ret_type = MdType::Void;
	bool aborted = false;

	if (cb.flags & CBF_PASS_PARAMS_POINTER)
	{
		param_count = 1;
		param = &one_param_ptr;
		one_param_ptr = &one_param;
		one_param.SetValue((UINT_PTR)params);
	}
	else
	{
		param = (ExprTokenType **)_alloca(cb.actual_param_count * sizeof(ExprTokenType *));
		if (arg_type) // v2.1: Convert incoming parameters to previously specified types.
		{
			UINT_PTR *next_param = params;
			
			auto &ret = arg_type[cb.actual_param_count];
			if (ret.type == MdType::Struct)
				ret_size = (int)ret.proto->StructSize();
			if (ret_size < 3 || ret_size == 4 || ret_size == 8)
				ret_ptr = &number_to_return;
			else
			{
				ret_ptr = *(void**)next_param; // First real arg is a pointer to the return struct.
				number_to_return = (UINT_PTR)ret_ptr;
				next_param++;
			}

			for (; param_count < cb.actual_param_count; ++param_count, ++next_param)
			{
				if (arg_type[param_count].type == MdType::Struct)
				{
					FuncResult &result_token = *(new (_alloca(sizeof(FuncResult))) FuncResult);
					param[param_count] = &result_token;
					// Create a struct Object as a by-value copy of the parameter, to avoid aliasing
					// stack memory if the script retains a reference after the function returns.
#ifdef WIN32_PLATFORM
					auto p = (UINT_PTR)next_param;
#else
					auto ss = arg_type[param_count].proto->StructSize();
					// For struct parameters of sizes other than those below, x64 passes them by address.
					auto p = (ss < 3 || ss == 4 || ss == 8) ? (UINT_PTR)next_param : *next_param;
#endif
					auto obj = Object::CreateStructCopyNoDelete(arg_type[param_count].proto, p);
					if (!obj)
					{
						aborted = true;
						break;
					}
					auto result = obj->Invoke(result_token, IT_GET | IF_BYPASS_METAFUNC, _T("__value"), ExprTokenType(obj), nullptr, 0);
					if (result == FAIL || result == EARLY_EXIT)
					{
						aborted = true;
						break;
					}
					if (result != INVOKE_NOT_HANDLED)
						obj->Release();
					else
						result_token.SetValue(obj);
#ifdef WIN32_PLATFORM
					// Adjust for structs larger than 4 bytes passed by value.
					next_param += (((int)arg_type[param_count].proto->StructSize() + 3) >> 2) - 1;
#endif
				}
				else
				{
					param[param_count] = (ExprTokenType *)_alloca(sizeof(ExprTokenType));
					TypedPtrToToken(arg_type[param_count].type, next_param, *param[param_count]);
#ifdef WIN32_PLATFORM
					switch (arg_type[param_count].type)
					{
					case MdType::Float64:
					case MdType::Int64:
					case MdType::UInt64:
						++next_param; // An additional 4 bytes.
					}
#endif
				}
			}
		}
		else
		{
			param_count = cb.actual_param_count;
			for (int i = 0; i < param_count; ++i)
			{
				param[i] = (ExprTokenType *)_alloca(sizeof(ExprTokenType));
				param[i]->SetValue((UINT_PTR)params[i]);
			}
		}
	}
	
	FuncResult result_token;
	if (!aborted)
	{
		ExprTokenType this_token(cb.func);
		auto result = cb.func->Invoke(result_token, IT_CALL, nullptr, this_token, param, param_count);
		if (result == INVOKE_NOT_HANDLED)
			result = result_token.UnknownMemberError(this_token, IT_CALL, nullptr);
		aborted = result == FAIL || result == EARLY_EXIT;
		if (!arg_type)
		{
			// UINT_PTR and not UINT64, to preserve documented behaviour on 32-bit builds.
			number_to_return = (UINT_PTR)TokenToInt64(result_token);
		}
		else if (!(result == FAIL || result == EARLY_EXIT))
		{
			auto &ret = arg_type[cb.actual_param_count];
			ret_type = ret.type;
			if (ret_type == MdType::Struct)
			{
				Object *obj;
				auto result_obj = TokenToObject(result_token);
				if (result_obj && result_obj->IsOfType(ret.proto))
					obj = (Object*)result_obj;
				else
				{
					FuncResult fr;
					auto result = Object::CreateStruct(fr, ret.proto);
					if (result == OK)
					{
						ASSERT(fr.symbol == SYM_OBJECT && dynamic_cast<Object*>(fr.object));
						// New has set fr.object=obj but has not called AddRef, so don't call Free.
						obj = (Object*)fr.object;
						fr.InitInvokeRetVal();
						ExprTokenType *prm = &result_token;
						// Invoke __Value as a "conversion operator".
						result = obj->Invoke(fr, IT_SET | IF_BYPASS_METAFUNC | IF_NO_NEW_PROPS, _T("__Value"), ExprTokenType(obj), &prm, 1);
						fr.Free();
						if (result == INVOKE_NOT_HANDLED) // Conversion not handled; report the error.
						{
							auto classname = ret.proto->GetOwnPropString(_T("__Class"));
							TypeError(classname ? classname : _T("Struct"), fr);
						}
					}
					//else an error was already raised.  Copy obj's data anyway (possibly all zeroes),
					// rather than leaving the return value undefined.
				}
				// Copy returned struct by value into the caller-provided space.
				// obj can be a derived struct class, in which case it must be truncated.
				memcpy(ret_ptr, (void*)obj->DataPtr(), ret.proto->StructSize());
				if (obj != result_obj)
					obj->Release();
			}
			else if (ret_type != MdType::Void)
			{
				SetValueOfTypeAtPtr(ret_type, ret_ptr, result_token, result_token);
			}
		}
		result_token.Free();
	}

	if (arg_type)
		while (--param_count >= 0)
			if (arg_type[param_count].type == MdType::Struct)
				((FuncResult *)param[param_count])->Free();
	
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

	// Use some hackery to set FPU register immediately prior to return.
	// It may have already been set as a side-effect of SetValueOfTypeAtPtr,
	// but there are cases where it can be overwritten before this point.
	if (ret_type == MdType::Float32)
		ReturnFloat(*(float*)ret_ptr);
	else if (ret_type == MdType::Float64)
		ReturnDouble(*(double*)ret_ptr);
	return number_to_return;
}



bif_impl FResult CallbackCreate(IObject *func, optl<StrArg> aOptions, ExprTokenType *aParams, UINT_PTR &aRetVal)
// Returns: Address of callback procedure.
// Parameters:
// 1: Name of the function to be called when the callback routine is executed.
// 2: Options.
// 3: Number of parameters of callback.
//
// Author: Original x86 RegisterCallback() was created by Jonathan Rennison (JGR).
//   x64 support by fincs.  Various changes by Lexikos.
{
	auto options = aOptions.value_or_empty();
	bool pass_params_pointer = _tcschr(options, '&'); // Callback wants the address of the parameter list instead of their values.
#ifdef WIN32_PLATFORM
	bool use_cdecl = StrChrAny(options, _T("Cc")); // Recognize "C" as the "CDecl" option.
	bool require_param_count = !use_cdecl; // Param count must be specified for x86 stdcall.
#else
	bool require_param_count = false;
#endif

	bool params_specified = aParams != nullptr;
	Array *param_types = nullptr; // v2.1: Pass an array of parameter types.
	IObjectRef obj_ref;
	if (params_specified && !TokenIsNumeric(*aParams))
	{
		param_types = Array::FromEnumerable(*aParams);
		if (!param_types)
			return FR_FAIL;
		obj_ref = param_types;
		param_types->Release();
	}
	int actual_param_count = param_types ? (int)param_types->Length() - 1 : aParams ? (int)TokenToInt64(*aParams) : 0;
	if (pass_params_pointer && (param_types || require_param_count && !params_specified)
		|| (actual_param_count < 0 || actual_param_count > 255))
		return FR_E_ARG(2);

	ResultToken result_token; // Just used for .result.
	auto fr = ValidateFunctor(func
		, pass_params_pointer ? 1 : actual_param_count // Count of script parameters being passed.
		// Use MinParams as actual_param_count if unspecified and no & option.
		, params_specified || pass_params_pointer ? nullptr : &actual_param_count);
	if (fr != OK)
		return fr;
	
#ifdef WIN32_PLATFORM
	if (!use_cdecl && actual_param_count > 31) // The ASM instruction currently used limits parameters to 31 (which should be plenty for any realistic use).
		return FR_E_ARG(2);
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
	SIZE_T rc_size = sizeof(RCCallbackFunc) + (param_types ? param_types->Length() * sizeof(RCArgType) : 0);
	RCCallbackFunc *callbackfunc = (RCCallbackFunc*) GlobalAlloc(GMEM_FIXED, rc_size);	//allocate structure off process heap, automatically RWE and fixed.
	if (!callbackfunc)
		return FR_E_OUTOFMEM;
	RCCallbackFunc &cb = *callbackfunc; // For convenience and possible code-size reduction.
	RCArgType *at = nullptr;

	if (param_types) // v2.1: Pass an array of parameter types.
	{
		at = (RCArgType*)(callbackfunc + 1);
		ExprTokenType v;
		for (UINT i = 0; i < param_types->Length(); ++i)
		{
			param_types->ItemToToken(i, v);
			at[i].type = MdType::Void;
			at[i].proto = nullptr;
			if (v.symbol == SYM_OBJECT && v.object->IsOfType(Object::sPrototype))
			{
				auto cls = (Object*)v.object;
				auto proto = cls->ClassGetPrototype();
				if (proto && proto->IsDerivedFrom(Object::sStructPrototype))
				{
					at[i].type = proto->GetStructMdType();
					if (at[i].type == MdType::Void)
					{
						at[i].type = MdType::Struct;
						at[i].proto = (Object*)proto;
					}
					continue;
				}
			}
			else if (v.symbol == SYM_STRING && !_tcsicmp(v.marker, _T("Void")))
			{
				if (i + 1 == param_types->Length())
					continue;
			}
			GlobalFree((HGLOBAL)callbackfunc);
			return FR_E_ARG(2);
		}
		for (UINT i = 0; i < param_types->Length(); ++i)
			if (at[i].proto)
				at[i].proto->AddRef();
	}

#ifdef WIN32_PLATFORM
	int param_slot_count = actual_param_count;
	if (at && !use_cdecl) // v2.1: Adjust for structs larger than 4 bytes passed by value.
	{
		for (int i = 0; i < actual_param_count; ++i)
			if (at[i].proto)
				param_slot_count += (((int)at[i].proto->LockStructSize() + 3) >> 2) - 1;
		if (at[actual_param_count].proto) // Return type is a struct.
		{
			int ret_size = (int)at[actual_param_count].proto->LockStructSize();
			if (!(ret_size < 3 || ret_size == 4 || ret_size == 8)) // Not returned by register.
				++param_slot_count;
		}
	}

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

	static auto funcaddrptr = &RegisterCallbackCStub; // Use fixed absolute address of pointer to function, instead of varying relative offset to function.
	cb.callfuncptr = &funcaddrptr; // xxxx: Address of C stub.

	cb.data4=0xC48359 // pop ecx -- 59 ;return address... add esp, xx -- 83 C4 xx ;stack correct (add argument to add esp, nn for stack correction).
		+ (use_cdecl ? 0 : param_slot_count << 26);

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
	if (param_types)
		cb.flags |= CBF_HAS_ARG_TYPES;

	// If DEP is enabled (and sometimes when DEP is apparently "disabled"), we must change the
	// protection of the page of memory in which the callback resides to allow it to execute:
	DWORD dwOldProtect;
	VirtualProtect(callbackfunc, sizeof(RCCallbackFunc), PAGE_EXECUTE_READWRITE, &dwOldProtect);

	aRetVal = (UINT_PTR)callbackfunc; // Yield the callable address as the result.
	return OK;
}

bif_impl FResult CallbackFree(UINT_PTR aCallback)
{
	if (aCallback < 65536) // Basic sanity check to catch incoming raw addresses that are zero or blank.  On Win32, the first 64KB of address space is always invalid.
		return FR_E_ARG(0);
	RCCallbackFunc *callbackfunc = (RCCallbackFunc *)aCallback;
	callbackfunc->func->Release();
	callbackfunc->func = NULL; // To help detect bugs.
	if (callbackfunc->flags & CBF_HAS_ARG_TYPES)
	{
		auto at = (RCArgType*)(callbackfunc + 1);
		for (int i = 0; i <= callbackfunc->actual_param_count; ++i)
			if (at[i].proto)
				at[i].proto->Release();
	}
	GlobalFree(callbackfunc);
	return OK;
}

#endif
