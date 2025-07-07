;///////////////////////////////////////////////////////////////////////
;
;	Windows ARM64 dynamic function call
;	Based on dyncall-1.4 implementation
;
;///////////////////////////////////////////////////////////////////////

; ARM64 calling convention:
; x0-x7: integer arguments
; v0-v7: floating point arguments
; x0: return value (integer)
; v0: return value (floating point)

	AREA |.text|, CODE, READONLY, ARM64

; PerformDynaCall function
; Arguments:
;   x0: size of arguments to be passed via stack
;   x1: pointer to arguments to be passed via stack
;   x2: pointer to arguments to be passed by registers
;   x3: target function pointer
	EXPORT PerformDynaCall
PerformDynaCall PROC
	stp x29, x30, [sp, #-16]!
	mov x29, sp
	ldr d0, [x2, #0]
	ldr d1, [x2, #8]
	ldr d2, [x2, #16]
	ldr d3, [x2, #24]
	ldr d4, [x2, #32]
	ldr d5, [x2, #40]
	ldr d6, [x2, #48]
	ldr d7, [x2, #56]
	sub sp, sp, x0
	eor x4, x4, x4
	mov x5, x1
	mov x6, sp
PerformDynaCall_next
	cmp x4, x0
	b.ge PerformDynaCall_done
	ldp x7, x9, [x5], #16
	stp x7, x9, [x6], #16
	add x4, x4, 16
	b PerformDynaCall_next
PerformDynaCall_done
	mov x9, x3
	add x10, x2, 64
	ldr x0, [x10, #0]
	ldr x1, [x10, #8]
	ldr x2, [x10, #16]
	ldr x3, [x10, #24]
	ldr x4, [x10, #32]
	ldr x5, [x10, #40]
	ldr x6, [x10, #48]
	ldr x7, [x10, #56]
	blr x9
	mov sp, x29
	ldp x29, x30, [sp], 16
	ret
	ENDP

; DynaCall function (simplified version)
; Arguments:
;   x0: number of QWORDs to be passed as arguments
;   x1: pointer to arguments
;   x2: target function pointer
;   x3: not used
	EXPORT DynaCall
DynaCall PROC
	stp x29, x30, [sp, #-16]!
	mov x29, sp
	mov x4, 8
	cmp x4, x0
	csel x4, x4, x0, gt
	lsl x4, x4, 3
	sub sp, sp, x4
	mov x8, sp
	and x8, x8, #-16
	mov sp, x8
	mov x9, x2
	mov x4, 0
	mov x5, x1
	mov x6, sp
	cmp x0, 0
	beq DynaCall_skip_copy
DynaCall_copy_args
	ldr x7, [x5], 8
	str x7, [x6], 8
	add x4, x4, 1
	cmp x4, x0
	b.lt DynaCall_copy_args
DynaCall_skip_copy
	ldp x0, x1, [sp, #0]
	ldp x2, x3, [sp, #16]
	ldp x4, x5, [sp, #32]
	ldp x6, x7, [sp, #48]
	ldp d0, d1, [sp, #64]
	ldp d2, d3, [sp, #80]
	ldp d4, d5, [sp, #96]
	ldp d6, d7, [sp, #112]
	blr x9
	mov sp, x29
	ldp x29, x30, [sp], 16
	ret
	ENDP

; GetFloatRetval function
; Return value is already in s0
	EXPORT GetFloatRetval
GetFloatRetval PROC
	ret
	ENDP

; GetDoubleRetval function
; Return value is already in d0
	EXPORT GetDoubleRetval
GetDoubleRetval PROC
	ret
	ENDP

	END 