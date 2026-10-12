; Match VBA's x87 expression evaluation and store-to-Double boundaries.
; Uses the caller's floating-point control word without changing it.
; SPDX-License-Identifier: LGPL-2.1-or-later
.code
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
