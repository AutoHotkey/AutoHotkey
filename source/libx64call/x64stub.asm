;///////////////////////////////////////////////////////////////////////
;
;	Windows x64 RegisterCallback stub
;	written by fincs
;
;///////////////////////////////////////////////////////////////////////

.code

CallbackFunctionOffset = 8*3

RegisterCallbackAsmStub proc frame
	option prologue:none, epilogue:none
.allocstack 8*5
.endprolog

	; For the 'push' further below
	add rsp, 8*5

	; Save the parameters in the spill area for consistency
	push r9
	push r8
	push rdx
	push rcx

	; Set parameters for the upcoming function call
	mov rcx, rsp ; UINT_PTR* aParams
	mov rdx, rax ; RCCallbackFunc* cbAddress

	; Call callback stub function
	sub rsp, 8*6 ; retaddr+padding+spill area
	call qword ptr[rax+CallbackFunctionOffset]
	add rsp, 8*5

	; Return
	ret
RegisterCallbackAsmStub endp

end
