;///////////////////////////////////////////////////////////////////////
;
;   AutoHotkey v2 / Windows ARM64 dynamic call bridge
;   Based on dyncall ARM64 MASM call core and older AutoHotkey ARM64 port ideas.
;
;   Provides the symbols expected by AutoHotkey v2.0.x:
;       PerformDynaCall
;       DynaCall
;       GetFloatRetval
;       GetDoubleRetval
;       RegisterCallbackAsmStub
;
;   Assemble with:
;       armasm64 /c /Foarm64call.obj arm64call_ahk_v2_dyncall_based.asm
;
;   Notes:
;   - PerformDynaCall matches AHK v2 DllCall.cpp declaration:
;       extern "C" UINT_PTR PerformDynaCall(size_t stackArgsSize,
;           DWORD_PTR* stackArgs, DWORD_PTR* regArgs, void* aFunction);
;
;     Incoming ARM64 registers:
;       x0 = stackArgsSize in bytes
;       x1 = stackArgs pointer
;       x2 = regArgs pointer, AHK currently prepares 4 qwords in x64 path
;       x3 = target function pointer
;
;   - DynaCall matches MdFunc.cpp declaration:
;       extern "C" UINT64 DynaCall(size_t aArgCount, UINT_PTR *aArg,
;           void *aFunction, DWORD aFlag);
;
;     Incoming ARM64 registers:
;       x0 = argument count
;       x1 = argument array, one UINT_PTR slot per arg
;       x2 = target function pointer
;       x3 = flag/thiscall indicator currently ignored here
;
;   - CallbackCreate still needs a C++ side ARM64 patch in CCallback.cpp,
;     because the current _WIN64 block writes x64 machine-code bytes into
;     RCCallbackFunc. This stub exports the expected symbol and mirrors the
;     old ARM64-port idea, but the entry thunk layout must match CCallback.cpp.
;
;///////////////////////////////////////////////////////////////////////

        AREA .text, CODE, ARM64

        EXPORT PerformDynaCall
        EXPORT DynaCall
        EXPORT GetFloatRetval
        EXPORT GetDoubleRetval
        EXPORT RegisterCallbackAsmStub

; ----------------------------------------------------------------------
; UINT_PTR PerformDynaCall(size_t stackArgsSize,
;                          DWORD_PTR* stackArgs,
;                          DWORD_PTR* regArgs,
;                          void* aFunction)
;
; AHK v2 x64-oriented C++ path supplies 4 register slots and remaining
; args in stackArgs. ARM64 has x0-x7 and d0-d7. To avoid changing C++ for
; this first step, we load x0-x3/d0-d3 from regArgs and, if present, lift
; the first 4 stack argument slots into x4-x7/d4-d7. The remaining stack
; args begin after those lifted slots.
;
; This makes DllCall with up to 8 integer/pointer args behave more like
; native Windows ARM64 without immediately rewriting DllCall.cpp.
; ----------------------------------------------------------------------

PerformDynaCall PROC
        ; Prolog.
        stp     x29, x30, [sp, #-16]!
        mov     x29, sp

        ; Preserve callee-saved registers we use.
        stp     x19, x20, [sp, #-16]!
        stp     x21, x22, [sp, #-16]!
        stp     x23, x24, [sp, #-16]!
        stp     x25, x26, [sp, #-16]!

        ; Incoming AHK v2 signature:
        ;   x0 = stackArgsSize in bytes
        ;   x1 = stackArgs pointer
        ;   x2 = regArgs pointer
        ;   x3 = target function pointer
        ;
        ; For ARM64, DllCall.cpp must provide regArgs in dyncall layout:
        ;   regArgs[0..7]  = payload for d0-d7
        ;   regArgs[8..15] = payload for x0-x7
        mov     x19, x0              ; stackArgsSize
        mov     x20, x1              ; stackArgs
        mov     x21, x2              ; regArgs
        mov     x22, x3              ; target function

        ; Allocate callee stack area, 16-byte aligned.
        add     x23, x19, #15
        and     x23, x23, #-16       ; aligned stack bytes
        sub     sp, sp, x23

        ; Copy stack arguments in 8-byte slots.
        mov     x24, xzr             ; offset
        cbz     x19, pf_stack_done
pf_stack_loop
        cmp     x24, x19
        b.ge    pf_stack_done
        ldr     x25, [x20, x24]
        str     x25, [sp, x24]
        add     x24, x24, #8
        b       pf_stack_loop
pf_stack_done

        ; Load floating-point argument registers d0-d7.
        ldr     d0, [x21, #0]
        ldr     d1, [x21, #8]
        ldr     d2, [x21, #16]
        ldr     d3, [x21, #24]
        ldr     d4, [x21, #32]
        ldr     d5, [x21, #40]
        ldr     d6, [x21, #48]
        ldr     d7, [x21, #56]

        ; Load integer/pointer argument registers x0-x7.
        ldr     x0, [x21, #64]
        ldr     x1, [x21, #72]
        ldr     x2, [x21, #80]
        ldr     x3, [x21, #88]
        ldr     x4, [x21, #96]
        ldr     x5, [x21, #104]
        ldr     x6, [x21, #112]
        ldr     x7, [x21, #120]

        ; Call target.
        blr     x22

        ; Save floating-point return value immediately.
        ldr     x10, =AhkArm64FloatRetval
        str     s0, [x10]
        ldr     x10, =AhkArm64DoubleRetval
        str     d0, [x10]

        ; Restore stack and callee-saved regs.
        mov     sp, x29
        sub     sp, sp, #64
        ldp     x25, x26, [sp], #16
        ldp     x23, x24, [sp], #16
        ldp     x21, x22, [sp], #16
        ldp     x19, x20, [sp], #16
        ldp     x29, x30, [sp], #16
        ret
        ENDP

; ----------------------------------------------------------------------
; UINT64 DynaCall(size_t aArgCount, UINT_PTR *aArg, void *aFunction, DWORD aFlag)
;
; Used by MdFunc.cpp for calls to known internal native functions.
; This maps arg[0..7] to x0..x7 and d0..d7 and places arg[8+] on stack.
; ----------------------------------------------------------------------

DynaCall PROC
        stp     x29, x30, [sp, #-16]!
        mov     x29, sp

        stp     x19, x20, [sp, #-16]!
        stp     x21, x22, [sp, #-16]!
        stp     x23, x24, [sp, #-16]!
        stp     x25, x26, [sp, #-16]!

        mov     x19, x0              ; arg count
        mov     x20, x1              ; arg array
        mov     x21, x2              ; target function

        ; stackCount = max(argCount - 8, 0)
        cmp     x19, #8
        sub     x22, x19, #8
        csel    x22, x22, xzr, hi
        lsl     x23, x22, #3         ; stack bytes
        add     x24, x23, #15
        and     x24, x24, #-16       ; aligned stack bytes
        sub     sp, sp, x24

        ; Copy arg[8+] to callee stack area.
        mov     x25, xzr             ; offset bytes
        cbz     x23, md_stack_done
md_stack_loop
        cmp     x25, x23
        b.ge    md_stack_done
        add     x26, x20, #64        ; &arg[8]
        ldr     x4, [x26, x25]
        str     x4, [sp, x25]
        add     x25, x25, #8
        b       md_stack_loop
md_stack_done

        ; Load x0-x7/d0-d7, absent args become zero.
        mov     x0, xzr
        mov     x1, xzr
        mov     x2, xzr
        mov     x3, xzr
        mov     x4, xzr
        mov     x5, xzr
        mov     x6, xzr
        mov     x7, xzr
        fmov    d0, xzr
        fmov    d1, xzr
        fmov    d2, xzr
        fmov    d3, xzr
        fmov    d4, xzr
        fmov    d5, xzr
        fmov    d6, xzr
        fmov    d7, xzr

        cmp     x19, #1
        b.lo    md_regs_done
        ldr     x0, [x20, #0]
        ldr     d0, [x20, #0]
        cmp     x19, #2
        b.lo    md_regs_done
        ldr     x1, [x20, #8]
        ldr     d1, [x20, #8]
        cmp     x19, #3
        b.lo    md_regs_done
        ldr     x2, [x20, #16]
        ldr     d2, [x20, #16]
        cmp     x19, #4
        b.lo    md_regs_done
        ldr     x3, [x20, #24]
        ldr     d3, [x20, #24]
        cmp     x19, #5
        b.lo    md_regs_done
        ldr     x4, [x20, #32]
        ldr     d4, [x20, #32]
        cmp     x19, #6
        b.lo    md_regs_done
        ldr     x5, [x20, #40]
        ldr     d5, [x20, #40]
        cmp     x19, #7
        b.lo    md_regs_done
        ldr     x6, [x20, #48]
        ldr     d6, [x20, #48]
        cmp     x19, #8
        b.lo    md_regs_done
        ldr     x7, [x20, #56]
        ldr     d7, [x20, #56]
md_regs_done

        blr     x21

        ; Save floating-point return value immediately for MdFunc callers too.
        ldr     x10, =AhkArm64FloatRetval
        str     s0, [x10]
        ldr     x10, =AhkArm64DoubleRetval
        str     d0, [x10]

        mov     sp, x29
        sub     sp, sp, #64
        ldp     x25, x26, [sp], #16
        ldp     x23, x24, [sp], #16
        ldp     x21, x22, [sp], #16
        ldp     x19, x20, [sp], #16
        ldp     x29, x30, [sp], #16
        ret
        ENDP

; ----------------------------------------------------------------------
; Floating-point return helpers.
; Like x64 version, these intentionally do nothing. C++ declares their
; return type so the compiler reads the current floating return register.
; ----------------------------------------------------------------------

GetFloatRetval PROC
        ldr     x10, =AhkArm64FloatRetval
        ldr     s0, [x10]
        ret
        ENDP

GetDoubleRetval PROC
        ldr     x10, =AhkArm64DoubleRetval
        ldr     d0, [x10]
        ret
        ENDP

; ----------------------------------------------------------------------
; RegisterCallbackAsmStub
;
; Exported so the project links. This mirrors the old AHK ARM64-port idea:
; save x0-x7, pass a parameter buffer and callback structure address to the
; C stub. However CCallback.cpp must generate an ARM64-compatible entry thunk
; which puts the RCCallbackFunc* in x9 before jumping here.
;
; Expected on entry after C++ patch:
;       x9 = RCCallbackFunc* cb
;       cb.callfuncptr/stub layout must be verified in CCallback.cpp.
; ----------------------------------------------------------------------

RegisterCallbackAsmStub PROC
        stp     x29, x30, [sp, #-16]!
        mov     x29, sp

        ; Reserve 64 bytes for x0-x7 and 64 bytes scratch/shadow, aligned.
        sub     sp, sp, #128

        stp     x0, x1, [sp, #0]
        stp     x2, x3, [sp, #16]
        stp     x4, x5, [sp, #32]
        stp     x6, x7, [sp, #48]

        ; x0 = UINT_PTR* params
        ; x1 = char* address / RCCallbackFunc* cb
        mov     x0, sp
        mov     x1, x9

        ; In AHK v2 x64 layout, callfuncptr follows stub pointer within RCCallbackFunc.
        ; For ARM64 C++ patch, keep callfuncptr at offset 24 if using:
        ;   UINT64 data1;
        ;   UINT64 data2;
        ;   void (*stub)();
        ;   UINT_PTR (CALLBACK *callfuncptr)(UINT_PTR*, char*);
        ldr     x11, [x9, #24]
        blr     x11

        mov     sp, x29
        ldp     x29, x30, [sp], #16
        ret
        ENDP


; ----------------------------------------------------------------------
; Storage for floating-point return values.
; Must be writable data, not code, because PerformDynaCall/DynaCall store d0
; immediately after the target call. Thread-local would be nicer long-term,
; but this is sufficient to verify ARM64 Double/Float return handling and
; matches AHK's mostly single-call synchronous use here.
; ----------------------------------------------------------------------

        AREA .data, DATA, ALIGN=3
AhkArm64FloatRetval
        DCD     0
        DCD     0
AhkArm64DoubleRetval
        DCQ     0

        END
