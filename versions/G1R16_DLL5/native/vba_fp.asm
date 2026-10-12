; Match VBA's x87 expression evaluation and store-to-Double boundaries.
; Uses the caller's floating-point control word without changing it.
; SPDX-License-Identifier: LGPL-2.1-or-later
.code
; Fused sparse loops retain row order and round/check every updated Double.
rpx_fp_scatter PROC FRAME
    sub rsp,40
    .allocstack 40
    .endprolog
    mov r10,qword ptr [rsp+80]
    mov r11,qword ptr [rsp+88]
    movsd xmm0,qword ptr [rsp+96]
    movsd qword ptr [rsp],xmm0
scatter_next:
    cmp edx,r8d
    jge scatter_ok
    mov eax,dword ptr [r9+rdx*4]
    cmp eax,0
    jl scatter_bad
    cmp eax,ecx
    jge scatter_bad
    fld qword ptr [r10+rdx*8]
    fmul qword ptr [rsp]
    fsubr qword ptr [r11+rax*8]
    fstp qword ptr [r11+rax*8]
    mov rax,qword ptr [r11+rax*8]
    shr rax,52
    and eax,07FFh
    cmp eax,07FFh
    je scatter_overflow
    inc edx
    jmp scatter_next
scatter_bad:
    mov eax,100
    jmp scatter_exit
scatter_overflow:
    mov eax,206
    jmp scatter_exit
scatter_ok:
    xor eax,eax
scatter_exit:
    add rsp,40
    ret
rpx_fp_scatter ENDP
rpx_fp_gather PROC FRAME
    sub rsp,40
    .allocstack 40
    .endprolog
    mov r10,qword ptr [rsp+80]
    mov r11,qword ptr [rsp+88]
    mov rax,qword ptr [rsp+96]
    mov qword ptr [rsp+16],rax
    movsd xmm0,qword ptr [rax]
    movsd qword ptr [rsp],xmm0
gather_next:
    cmp edx,r8d
    jge gather_ok
    mov eax,dword ptr [r9+rdx*4]
    cmp eax,0
    jl gather_bad
    cmp eax,ecx
    jge gather_bad
    fld qword ptr [r10+rdx*8]
    fmul qword ptr [r11+rax*8]
    fsubr qword ptr [rsp]
    fstp qword ptr [rsp]
    mov rax,qword ptr [rsp]
    shr rax,52
    and eax,07FFh
    cmp eax,07FFh
    je gather_overflow
    inc edx
    jmp gather_next
gather_bad:
    mov eax,100
    jmp gather_exit
gather_overflow:
    mov eax,206
    jmp gather_exit
gather_ok:
    mov rax,qword ptr [rsp+16]
    movsd xmm0,qword ptr [rsp]
    movsd qword ptr [rax],xmm0
    xor eax,eax
gather_exit:
    add rsp,40
    ret
rpx_fp_gather ENDP
; Every fstp below is the same stored-Double boundary as the original helper.
DD_SUM MACRO lhs, rhs, out_hi, out_lo, scratch
    fld lhs
    fadd rhs
    fstp out_hi
    fld out_hi
    fsub lhs
    fstp scratch
    fld out_hi
    fsub scratch
    fsubr lhs
    fld rhs
    fsub scratch
    faddp st(1),st(0)
    fstp out_lo
ENDM
DD_TERM MACRO lhs, rhs
    fld lhs
    fmul rhs
    fstp qword ptr [rsp+48]
    fld qword ptr [rsp+40]
    fadd qword ptr [rsp+48]
    fstp qword ptr [rsp+40]
ENDM
rpx_fp_dd_add PROC FRAME
    sub rsp,120
    .allocstack 120
    .endprolog
    DD_SUM qword ptr [rcx],qword ptr [rdx],qword ptr [rsp],qword ptr [rsp+8],qword ptr [rsp+96]
    DD_SUM qword ptr [rcx+8],qword ptr [rdx+8],qword ptr [rsp+16],qword ptr [rsp+24],qword ptr [rsp+96]
    fld qword ptr [rsp+8]
    fadd qword ptr [rsp+16]
    fstp qword ptr [rsp+32]
    DD_SUM qword ptr [rsp],qword ptr [rsp+32],qword ptr [rsp+40],qword ptr [rsp+48],qword ptr [rsp+96]
    fld qword ptr [rsp+48]
    fadd qword ptr [rsp+24]
    fstp qword ptr [rsp+56]
    DD_SUM qword ptr [rsp+40],qword ptr [rsp+56],qword ptr [r8],qword ptr [r8+8],qword ptr [rsp+96]
    add rsp,120
    ret
rpx_fp_dd_add ENDP
rpx_fp_dd_product_parts PROC FRAME
    sub rsp,120
    .allocstack 120
    .endprolog
    mov r10,rcx
    mov r11,rdx
    movsd qword ptr [rsp],xmm2
    movsd qword ptr [rsp+8],xmm3
    movsd xmm0,qword ptr [rsp+160]
    movsd qword ptr [rsp+16],xmm0
    movsd xmm0,qword ptr [rsp+168]
    movsd qword ptr [rsp+24],xmm0
    movsd xmm0,qword ptr [rsp+176]
    movsd qword ptr [rsp+32],xmm0
    mov r9,qword ptr [rsp+184]
    fld qword ptr [rsp]
    fmul qword ptr [rsp+16]
    fstp qword ptr [rsp+48]
    fld qword ptr [rsp+48]
    fsub qword ptr [rsp+32]
    fstp qword ptr [rsp+40]
    DD_TERM qword ptr [rsp],qword ptr [rsp+24]
    DD_TERM qword ptr [rsp+8],qword ptr [rsp+16]
    DD_TERM qword ptr [rsp+8],qword ptr [rsp+24]
    DD_TERM qword ptr [r10],qword ptr [r11+8]
    DD_TERM qword ptr [r10+8],qword ptr [r11]
    DD_TERM qword ptr [r10+8],qword ptr [r11+8]
    DD_SUM qword ptr [rsp+32],qword ptr [rsp+40],qword ptr [r9],qword ptr [r9+8],qword ptr [rsp+56]
    add rsp,120
    ret
rpx_fp_dd_product_parts ENDP
rpx_fp_mulminus PROC FRAME
    sub rsp,40
    .allocstack 40
    .endprolog
    movsd qword ptr [rsp],xmm0
    movsd qword ptr [rsp+8],xmm1
    movsd qword ptr [rsp+16],xmm2
    fld qword ptr [rsp]
    fmul qword ptr [rsp+8]
    fsub qword ptr [rsp+16]
    fstp qword ptr [rsp]
    movsd xmm0,qword ptr [rsp]
    add rsp,40
    ret
rpx_fp_mulminus ENDP
rpx_fp_subsub PROC FRAME
    sub rsp,40
    .allocstack 40
    .endprolog
    movsd qword ptr [rsp],xmm0
    movsd qword ptr [rsp+8],xmm1
    movsd qword ptr [rsp+16],xmm2
    fld qword ptr [rsp]
    fsub qword ptr [rsp+8]
    fsub qword ptr [rsp+16]
    fstp qword ptr [rsp]
    movsd xmm0,qword ptr [rsp]
    add rsp,40
    ret
rpx_fp_subsub ENDP
rpx_fp_dot2 PROC FRAME
    sub rsp,40
    .allocstack 40
    .endprolog
    movsd qword ptr [rsp],xmm0
    movsd qword ptr [rsp+8],xmm1
    movsd qword ptr [rsp+16],xmm2
    movsd qword ptr [rsp+24],xmm3
    fld qword ptr [rsp]
    fmul qword ptr [rsp+8]
    fld qword ptr [rsp+16]
    fmul qword ptr [rsp+24]
    faddp st(1),st(0)
    fstp qword ptr [rsp]
    movsd xmm0,qword ptr [rsp]
    add rsp,40
    ret
rpx_fp_dot2 ENDP
rpx_fp_addmuldiv PROC FRAME
    sub rsp,40
    .allocstack 40
    .endprolog
    movsd qword ptr [rsp],xmm0
    movsd qword ptr [rsp+8],xmm1
    movsd qword ptr [rsp+16],xmm2
    movsd qword ptr [rsp+24],xmm3
    fld qword ptr [rsp+8]
    fmul qword ptr [rsp+16]
    fdiv qword ptr [rsp+24]
    fadd qword ptr [rsp]
    fstp qword ptr [rsp]
    movsd xmm0,qword ptr [rsp]
    add rsp,40
    ret
rpx_fp_addmuldiv ENDP
rpx_fp_submuldiv PROC FRAME
    sub rsp,40
    .allocstack 40
    .endprolog
    movsd qword ptr [rsp],xmm0
    movsd qword ptr [rsp+8],xmm1
    movsd qword ptr [rsp+16],xmm2
    movsd qword ptr [rsp+24],xmm3
    fld qword ptr [rsp+8]
    fmul qword ptr [rsp+16]
    fsubr qword ptr [rsp]
    fdiv qword ptr [rsp+24]
    fstp qword ptr [rsp]
    movsd xmm0,qword ptr [rsp]
    add rsp,40
    ret
rpx_fp_submuldiv ENDP
rpx_fp_jordan_div PROC FRAME
    sub rsp,40
    .allocstack 40
    .endprolog
    movsd qword ptr [rsp],xmm0
    movsd qword ptr [rsp+8],xmm1
    movsd qword ptr [rsp+16],xmm2
    movsd qword ptr [rsp+24],xmm3
    fld qword ptr [rsp]
    fmul qword ptr [rsp+8]
    fld qword ptr [rsp+16]
    fmul qword ptr [rsp+24]
    fsubp st(1),st(0)
    fld qword ptr [rsp+80]
    fmul qword ptr [rsp+88]
    fsubp st(1),st(0)
    fdiv qword ptr [rsp+96]
    fstp qword ptr [rsp]
    movsd xmm0,qword ptr [rsp]
    add rsp,40
    ret
rpx_fp_jordan_div ENDP
rpx_fp_near_cone PROC FRAME
    sub rsp,40
    .allocstack 40
    .endprolog
    movsd qword ptr [rsp],xmm0
    movsd qword ptr [rsp+8],xmm1
    mov rax,03eb0c6f7a0b5ed8dh
    mov qword ptr [rsp+16],rax
    fld qword ptr [rsp+16]
    fld qword ptr [rsp]
    fsub qword ptr [rsp+8]
    fdiv qword ptr [rsp]
    fcomip st(0),st(1)
    setb al
    movzx eax,al
    fstp st(0)
    add rsp,40
    ret
rpx_fp_near_cone ENDP
rpx_fp_mul PROC FRAME
    sub rsp,24
    .allocstack 24
    .endprolog
    movsd qword ptr [rsp],xmm0
    movsd qword ptr [rsp+8],xmm1
    fld qword ptr [rsp]
    fmul qword ptr [rsp+8]
    fstp qword ptr [rsp]
    movsd xmm0,qword ptr [rsp]
    add rsp,24
    ret
rpx_fp_mul ENDP
rpx_fp_add PROC FRAME
    sub rsp,24
    .allocstack 24
    .endprolog
    movsd qword ptr [rsp],xmm0
    movsd qword ptr [rsp+8],xmm1
    fld qword ptr [rsp]
    fadd qword ptr [rsp+8]
    fstp qword ptr [rsp]
    movsd xmm0,qword ptr [rsp]
    add rsp,24
    ret
rpx_fp_add ENDP
rpx_fp_sub PROC FRAME
    sub rsp,24
    .allocstack 24
    .endprolog
    movsd qword ptr [rsp],xmm0
    movsd qword ptr [rsp+8],xmm1
    fld qword ptr [rsp]
    fsub qword ptr [rsp+8]
    fstp qword ptr [rsp]
    movsd xmm0,qword ptr [rsp]
    add rsp,24
    ret
rpx_fp_sub ENDP
rpx_fp_mulsub PROC FRAME
    sub rsp,40
    .allocstack 40
    .endprolog
    movsd qword ptr [rsp],xmm0
    movsd qword ptr [rsp+8],xmm1
    movsd qword ptr [rsp+16],xmm2
    fld qword ptr [rsp+8]
    fmul qword ptr [rsp+16]
    fsubr qword ptr [rsp]
    fstp qword ptr [rsp]
    movsd xmm0,qword ptr [rsp]
    add rsp,40
    ret
rpx_fp_mulsub ENDP
rpx_fp_div PROC FRAME
    sub rsp,40
    .allocstack 40
    .endprolog
    movsd qword ptr [rsp],xmm0
    movsd qword ptr [rsp+8],xmm1
    fld qword ptr [rsp]
    fdiv qword ptr [rsp+8]
    fstp qword ptr [rsp]
    movsd xmm0,qword ptr [rsp]
    add rsp,40
    ret
rpx_fp_div ENDP
rpx_fp_adddiv PROC FRAME
    sub rsp,40
    .allocstack 40
    .endprolog
    movsd qword ptr [rsp],xmm0
    movsd qword ptr [rsp+8],xmm1
    movsd qword ptr [rsp+16],xmm2
    fld qword ptr [rsp]
    fadd qword ptr [rsp+8]
    fdiv qword ptr [rsp+16]
    fstp qword ptr [rsp]
    movsd xmm0,qword ptr [rsp]
    add rsp,40
    ret
rpx_fp_adddiv ENDP
rpx_fp_twosum PROC FRAME
    sub rsp,40
    .allocstack 40
    .endprolog
    movsd qword ptr [rsp],xmm0
    movsd qword ptr [rsp+8],xmm1
    fld qword ptr [rsp]
    fadd qword ptr [rsp+8]
    fstp qword ptr [r8]
    fld qword ptr [r8]
    fsub qword ptr [rsp]
    fstp qword ptr [rsp+16]
    fld qword ptr [r8]
    fsub qword ptr [rsp+16]
    fsubr qword ptr [rsp]
    fld qword ptr [rsp+8]
    fsub qword ptr [rsp+16]
    faddp st(1),st(0)
    fstp qword ptr [r8+8]
    add rsp,40
    ret
rpx_fp_twosum ENDP
rpx_fp_sqrt PROC FRAME
    sub rsp,24
    .allocstack 24
    .endprolog
    sqrtsd xmm0,xmm0
    add rsp,24
    ret
rpx_fp_sqrt ENDP
rpx_fp_norm2 PROC FRAME
    sub rsp,24
    .allocstack 24
    .endprolog
    movsd qword ptr [rsp],xmm0
    movsd qword ptr [rsp+8],xmm1
    fld qword ptr [rsp]
    fmul st(0),st(0)
    fld qword ptr [rsp+8]
    fmul st(0),st(0)
    faddp st(1),st(0)
    ; VBA Sqr's Double argument is rounded before the square-root operation.
    fstp qword ptr [rsp]
    movsd xmm0,qword ptr [rsp]
    sqrtsd xmm0,xmm0
    add rsp,24
    ret
rpx_fp_norm2 ENDP
rpx_fp_addmul PROC FRAME
    sub rsp,40
    .allocstack 40
    .endprolog
    movsd qword ptr [rsp],xmm0
    movsd qword ptr [rsp+8],xmm1
    movsd qword ptr [rsp+16],xmm2
    fld qword ptr [rsp+8]
    fmul qword ptr [rsp+16]
    fadd qword ptr [rsp]
    fstp qword ptr [rsp]
    movsd xmm0,qword ptr [rsp]
    add rsp,40
    ret
rpx_fp_addmul ENDP
rpx_fp_mul3 PROC FRAME
    sub rsp,40
    .allocstack 40
    .endprolog
    movsd qword ptr [rsp],xmm0
    movsd qword ptr [rsp+8],xmm1
    movsd qword ptr [rsp+16],xmm2
    fld qword ptr [rsp]
    fmul qword ptr [rsp+8]
    fmul qword ptr [rsp+16]
    fstp qword ptr [rsp]
    movsd xmm0,qword ptr [rsp]
    add rsp,40
    ret
rpx_fp_mul3 ENDP
rpx_fp_addmul3 PROC FRAME
    sub rsp,40
    .allocstack 40
    .endprolog
    movsd qword ptr [rsp],xmm0
    movsd qword ptr [rsp+8],xmm1
    movsd qword ptr [rsp+16],xmm2
    movsd qword ptr [rsp+24],xmm3
    fld qword ptr [rsp+8]
    fmul qword ptr [rsp+16]
    fmul qword ptr [rsp+24]
    fadd qword ptr [rsp]
    fstp qword ptr [rsp]
    movsd xmm0,qword ptr [rsp]
    add rsp,40
    ret
rpx_fp_addmul3 ENDP
rpx_fp_dot3 PROC FRAME
    sub rsp,40
    .allocstack 40
    .endprolog
    movsd qword ptr [rsp],xmm0
    movsd qword ptr [rsp+8],xmm1
    movsd qword ptr [rsp+16],xmm2
    movsd qword ptr [rsp+24],xmm3
    fld qword ptr [rsp]
    fmul qword ptr [rsp+8]
    fld qword ptr [rsp+16]
    fmul qword ptr [rsp+24]
    faddp st(1),st(0)
    fld qword ptr [rsp+80]
    fmul qword ptr [rsp+88]
    faddp st(1),st(0)
    fstp qword ptr [rsp]
    movsd xmm0,qword ptr [rsp]
    add rsp,40
    ret
rpx_fp_dot3 ENDP
rpx_fp_det3 PROC FRAME
    sub rsp,40
    .allocstack 40
    .endprolog
    movsd qword ptr [rsp],xmm0
    movsd qword ptr [rsp+8],xmm1
    movsd qword ptr [rsp+16],xmm2
    fld qword ptr [rsp]
    fmul st(0),st(0)
    fld qword ptr [rsp+8]
    fmul st(0),st(0)
    fsubp st(1),st(0)
    fld qword ptr [rsp+16]
    fmul st(0),st(0)
    fsubp st(1),st(0)
    fstp qword ptr [rsp]
    movsd xmm0,qword ptr [rsp]
    add rsp,40
    ret
rpx_fp_det3 ENDP
rpx_fp_ratio_delta PROC FRAME
    sub rsp,24
    .allocstack 24
    .endprolog
    movsd qword ptr [rsp],xmm0
    movsd qword ptr [rsp+8],xmm1
    fld qword ptr [rsp]
    fsub qword ptr [rsp+8]
    fdiv qword ptr [rsp]
    fstp qword ptr [rsp]
    movsd xmm0,qword ptr [rsp]
    add rsp,24
    ret
rpx_fp_ratio_delta ENDP
rpx_fp_lorentz_det PROC FRAME
    sub rsp,24
    .allocstack 24
    .endprolog
    movsd qword ptr [rsp],xmm0
    movsd qword ptr [rsp+8],xmm1
    fld qword ptr [rsp]
    fsub qword ptr [rsp+8]
    fld qword ptr [rsp]
    fadd qword ptr [rsp+8]
    fmulp st(1),st(0)
    fstp qword ptr [rsp]
    movsd xmm0,qword ptr [rsp]
    add rsp,24
    ret
rpx_fp_lorentz_det ENDP
rpx_fp_delta_div2 PROC FRAME
    sub rsp,40
    .allocstack 40
    .endprolog
    movsd qword ptr [rsp],xmm0
    movsd qword ptr [rsp+8],xmm1
    movsd qword ptr [rsp+16],xmm2
    fld qword ptr [rsp]
    fsub qword ptr [rsp+8]
    fld qword ptr [rsp+16]
    fadd st(0),st(0)
    fdivp st(1),st(0)
    fstp qword ptr [rsp]
    movsd xmm0,qword ptr [rsp]
    add rsp,40
    ret
rpx_fp_delta_div2 ENDP
rpx_fp_controlword PROC FRAME
    sub rsp,24
    .allocstack 24
    .endprolog
    fnstcw word ptr [rsp]
    movzx eax,word ptr [rsp]
    test ecx,ecx
    jz cw_done
    mov word ptr [rsp+8],dx
    fldcw word ptr [rsp+8]
cw_done:
    add rsp,24
    ret
rpx_fp_controlword ENDP
rpx_fp_bilinear3 PROC FRAME
    sub rsp,40
    .allocstack 40
    .endprolog
    movsd qword ptr [rsp],xmm0
    movsd qword ptr [rsp+8],xmm1
    movsd qword ptr [rsp+16],xmm2
    movsd qword ptr [rsp+24],xmm3
    fld qword ptr [rsp]
    fmul qword ptr [rsp+8]
    fld qword ptr [rsp+16]
    fmul qword ptr [rsp+24]
    fsubp st(1),st(0)
    fld qword ptr [rsp+80]
    fmul qword ptr [rsp+88]
    fsubp st(1),st(0)
    fadd st(0),st(0)
    fstp qword ptr [rsp]
    movsd xmm0,qword ptr [rsp]
    add rsp,40
    ret
rpx_fp_bilinear3 ENDP
rpx_fp_abs6 PROC FRAME
    sub rsp,40
    .allocstack 40
    .endprolog
    movsd qword ptr [rsp],xmm0
    movsd qword ptr [rsp+8],xmm1
    movsd qword ptr [rsp+16],xmm2
    movsd qword ptr [rsp+24],xmm3
    fld qword ptr [rsp]
    fabs
    fld qword ptr [rsp+8]
    fabs
    faddp st(1),st(0)
    fld qword ptr [rsp+16]
    fabs
    faddp st(1),st(0)
    fld qword ptr [rsp+24]
    fabs
    faddp st(1),st(0)
    fld qword ptr [rsp+80]
    fabs
    faddp st(1),st(0)
    fld qword ptr [rsp+88]
    fabs
    faddp st(1),st(0)
    fstp qword ptr [rsp]
    movsd xmm0,qword ptr [rsp]
    add rsp,40
    ret
rpx_fp_abs6 ENDP
rpx_fp_discriminant PROC FRAME
    sub rsp,40
    .allocstack 40
    .endprolog
    movsd qword ptr [rsp],xmm0
    movsd qword ptr [rsp+8],xmm1
    movsd qword ptr [rsp+16],xmm2
    movsd qword ptr [rsp+24],xmm3
    fld qword ptr [rsp]
    fmul st(0),st(0)
    fld qword ptr [rsp+8]
    fadd st(0),st(0)
    fadd st(0),st(0)
    fmul qword ptr [rsp+16]
    fsubp st(1),st(0)
    fstp qword ptr [rsp]
    movsd xmm0,qword ptr [rsp]
    add rsp,40
    ret
rpx_fp_discriminant ENDP
rpx_fp_boundary_tiny PROC FRAME
    sub rsp,40
    .allocstack 40
    .endprolog
    movsd qword ptr [rsp],xmm0
    movsd qword ptr [rsp+8],xmm1
    movsd qword ptr [rsp+16],xmm2
    movsd qword ptr [rsp+24],xmm3
    mov rax,03c32725dd1d243ach
    mov qword ptr [rsp+16],rax
    fld qword ptr [rsp+16]
    fmul qword ptr [rsp+8]
    fmul qword ptr [rsp+8]
    fld qword ptr [rsp]
    fabs
    fcomip st(0),st(1)
    setbe al
    movzx eax,al
    fstp st(0)
    add rsp,40
    ret
rpx_fp_boundary_tiny ENDP
rpx_fp_offset_tau PROC FRAME
    sub rsp,40
    .allocstack 40
    .endprolog
    movsd qword ptr [rsp],xmm0
    movsd qword ptr [rsp+8],xmm1
    movsd qword ptr [rsp+16],xmm2
    movsd qword ptr [rsp+24],xmm3
    mov rax,03ff0000000000000h
    mov qword ptr [rsp+24],rax
    fld qword ptr [rsp+8]
    fsub qword ptr [rsp+24]
    fmul qword ptr [rsp+16]
    fadd qword ptr [rsp]
    fstp qword ptr [rsp]
    movsd xmm0,qword ptr [rsp]
    add rsp,40
    ret
rpx_fp_offset_tau ENDP
rpx_fp_greater_product PROC FRAME
    sub rsp,40
    .allocstack 40
    .endprolog
    movsd qword ptr [rsp],xmm0
    movsd qword ptr [rsp+8],xmm1
    movsd qword ptr [rsp+16],xmm2
    fld qword ptr [rsp+8]
    fmul qword ptr [rsp+16]
    fld qword ptr [rsp]
    fcomip st(0),st(1)
    seta al
    movzx eax,al
    fstp st(0)
    add rsp,40
    ret
rpx_fp_greater_product ENDP
END
