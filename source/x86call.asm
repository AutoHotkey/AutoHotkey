.model flat, C
.code

DynaCall proc cargs:dword, pargs:dword, pfn:dword, opt:dword
    mov ecx, cargs
    mov edx, ecx        ; keep ecx = cargs for rep movsd
    shl edx, 2
    sub esp, edx        ; allocate cargs*4 bytes on stack
    
    ; MASM inserts a prolog/epilog which restores esp before return, but if we relied
    ; on it to also save/restore esi and edi, it would do so in a way that requires our
    ; call below to account for cdecl vs stdcall.  So we just do this manually.
    push esi
    push edi
    mov esi, pargs      ; set source
    lea edi, [esp+8]    ; set dest (+8 to adjust for pushing esi/edi)
    rep movsd           ; copy ecx (cargs) dwords from esi (pargs) to edi (stack)
    pop edi
    pop esi
    
    cmp opt, 0
    je over
    pop ecx             ; for thiscall

over:
    call pfn
    
    ret
DynaCall endp

GetFloatRetval proc
    ; Nothing is actually done here - we just declare the appropriate return type in C++.
    ret
GetFloatRetval endp

GetDoubleRetval proc
    ; See above.
    ret
GetDoubleRetval endp

end