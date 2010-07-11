;///////////////////////////////////////////////////////////////////////
;
;	Windows x64 dynamic function call
;	Code adapted from http://www.dyncall.org/
;
;///////////////////////////////////////////////////////////////////////

;//////////////////////////////////////////////////////////////////////////////
;
; Copyright (c) 2007-2009 Daniel Adler <dadler@uni-goettingen.de>, 
;                         Tassilo Philipp <tphilipp@potion-studios.com>
;
; Permission to use, copy, modify, and distribute this software for any
; purpose with or without fee is hereby granted, provided that the above
; copyright notice and this permission notice appear in all copies.
;
; THE SOFTWARE IS PROVIDED "AS IS" AND THE AUTHOR DISCLAIMS ALL WARRANTIES
; WITH REGARD TO THIS SOFTWARE INCLUDING ALL IMPLIED WARRANTIES OF
; MERCHANTABILITY AND FITNESS. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR
; ANY SPECIAL, DIRECT, INDIRECT, OR CONSEQUENTIAL DAMAGES OR ANY DAMAGES
; WHATSOEVER RESULTING FROM LOSS OF USE, DATA OR PROFITS, WHETHER IN AN
; ACTION OF CONTRACT, NEGLIGENCE OR OTHER TORTIOUS ACTION, ARISING OUT OF
; OR IN CONNECTION WITH THE USE OR PERFORMANCE OF THIS SOFTWARE.
;
;//////////////////////////////////////////////////////////////////////////////

.code

PerformDynaCall proc frame
	option prologue:none, epilogue:none

	; Arguments:
	; rcx: size of arguments to be passed via stack
	; rdx: pointer to arguments to be passed via stack
	; r8:  pointer to arguments to be passed by registers
	; r9:  target function pointer

	; Save some registers
	push rbp
.pushreg rbp
	push rsi
.pushreg rsi
	push rdi
.pushreg rdi

	; Save the stack pointer
	mov rbp, rsp
.setframe rbp, 0
.endprolog

	; Setup stack frame by subtracting the size of the arguments
	sub rsp, rcx

	; Ensure the stack is 16-byte aligned
	mov rax, rcx
	and rax, 15
	mov rsi, 16
	sub rsi, rax
	sub rsp, rsi

	; Save function address
	mov rax, r9

	; Copy the stack arguments
	mov rsi, rdx ; let rsi point to the arguments.
	mov rdi, rsp ; store pointer to beginning of stack arguments in rdi (for rep movsb).
	rep movsb    ; @@@ should be optimized (e.g. movq)

	; Copy the register arguments
	mov rcx,  qword ptr[r8   ] ; copy first four arguments to rcx, rdx, r8, r9 and xmm0-xmm3.
	mov rdx,  qword ptr[r8+ 8]
	mov r9,   qword ptr[r8+24] ; set r9 first to not overwrite r8 too soon.
	mov r8,   qword ptr[r8+16]
	movd xmm0, rcx
	movd xmm1, rdx
	movd xmm2, r8
	movd xmm3, r9

	; Call function
	sub rsp, 8*4 ; spill area
	call rax

	; Restore stack pointer
	;mov rsp, rbp	; Although mov works too, MSDN says "add RSP,constant or lea RSP,constant[FPReg]" must be used for exception-handling to work reliably. http://msdn.microsoft.com/en-us/library/tawsa7cb.aspx
	lea rsp, [rbp]

	; Restore some registers and return
	pop rdi
	pop rsi
	pop rbp
	ret
PerformDynaCall endp

read_xmm0_float proc
	option prologue:none, epilogue:none
	; Nothing is actually done here - we just declare the appropriate return type in C++.
	ret
read_xmm0_float endp

read_xmm0_double proc
	option prologue:none, epilogue:none
	; See above.
	ret
read_xmm0_double endp

end