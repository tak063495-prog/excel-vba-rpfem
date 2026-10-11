; Match VBA's x87 expression evaluation and store-to-Double boundaries.
; Uses the caller's floating-point control word without changing it.
; SPDX-License-Identifier: LGPL-2.1-or-later
.code
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
END
