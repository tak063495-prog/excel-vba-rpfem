/* G1R15 numeric kernels. Sparse LDL is adapted from the existing VBA
 * adaptation of Timothy A. Davis' LDL (LGPL-2.1-or-later).
 * SPDX-License-Identifier: LGPL-2.1-or-later
 * No allocations, global factor state, reordered reductions or FMA.
 * All arrays and the symbolic pattern remain owned by VBA.
 */
#include <stdint.h>
#include <stddef.h>
#include <string.h>
#include <math.h>
#include <float.h>
#include <windows.h>
#pragma fp_contract(off)
#define API __declspec(dllexport) int __stdcall
typedef struct { double hi, lo; } dd;
enum { OK=0, PIVOT=1, BAD_ARGUMENT=100, NONFINITE_PRODUCT=201,
       PRODUCT_UNDERFLOW=202, ZERO_DIVISOR=203, NONFINITE_VALUE=204,
       DOUBLE_ZERO_DIVISOR=205, DOUBLE_OVERFLOW=206 };
_Static_assert(sizeof(dd)==16, "VBA DD layout");
API rpx_abi(void) { return 0x00160003; }
extern double rpx_fp_mulsub(double accum,double a,double b);
extern double rpx_fp_mul(double a,double b);
extern double rpx_fp_add(double a,double b);
extern double rpx_fp_sub(double a,double b);
extern double rpx_fp_div(double a,double b);
extern double rpx_fp_adddiv(double a,double b,double divisor);
extern void rpx_fp_twosum(double a,double b,dd *out);
extern int32_t rpx_fp_controlword(int32_t mode,int32_t word);
API rpx_test_controlword(int32_t mode,int32_t word) { return rpx_fp_controlword(mode,word); }

static dd twosum(double a, double b) {
    dd x; rpx_fp_twosum(a,b,&x); return x;
}
static void split(double a, double *hi, double *lo) {
    uint64_t u; memcpy(&u,&a,8); u &= UINT64_C(0xfffffffff8000000);
    memcpy(hi,&u,8); *lo=rpx_fp_sub(a,*hi);
}
static dd add(dd a, dd b) {
    dd s=twosum(a.hi,b.hi), t=twosum(a.lo,b.lo), u=twosum(s.hi,rpx_fp_add(s.lo,t.hi));
    return twosum(u.hi,rpx_fp_add(u.lo,t.lo));
}
static dd sub(dd a, dd b) { b.hi=-b.hi; b.lo=-b.lo; return add(a,b); }
static dd from_double(double a) { dd x={a,0.0}; return x; }
static dd mul_known(dd a, dd b, double bh, double bl, int *error) {
    double p=rpx_fp_mul(a.hi,b.hi), e, ah,al,t;
    if(!isfinite(p)) { *error=DOUBLE_OVERFLOW; return from_double(0); }
    if(fabs(p)>1e290) { *error=NONFINITE_PRODUCT; return from_double(0); }
    if(p==0.0 && a.hi!=0.0 && b.hi!=0.0) { *error=PRODUCT_UNDERFLOW; return from_double(0); }
    split(a.hi,&ah,&al);
    t=rpx_fp_mul(ah,bh); e=rpx_fp_sub(t,p);
    t=rpx_fp_mul(ah,bl); e=rpx_fp_add(e,t); t=rpx_fp_mul(al,bh); e=rpx_fp_add(e,t); t=rpx_fp_mul(al,bl); e=rpx_fp_add(e,t);
    t=rpx_fp_mul(a.hi,b.lo); e=rpx_fp_add(e,t); t=rpx_fp_mul(a.lo,b.hi); e=rpx_fp_add(e,t); t=rpx_fp_mul(a.lo,b.lo); e=rpx_fp_add(e,t);
    return twosum(p,e);
}
static dd mul(dd a,dd b,int *error) { double bh,bl; split(b.hi,&bh,&bl); return mul_known(a,b,bh,bl,error); }
static dd divide(dd a,dd b,int *error) {
    dd q,r,t; int k;
    if(b.hi==0.0) { *error=ZERO_DIVISOR; return from_double(0); }
    q=from_double(rpx_fp_div(a.hi,b.hi));
    if(!isfinite(q.hi)) { *error=DOUBLE_OVERFLOW; return q; }
    for(k=0;k<2;k++) { t=mul(b,q,error); if(*error) return q; r=sub(a,t);
        t=from_double(rpx_fp_adddiv(r.hi,r.lo,b.hi));
        if(!isfinite(t.hi)) { *error=DOUBLE_OVERFLOW; return q; }
        q=add(q,t); }
    return q;
}
static double value(dd a,int *error) {
    double v=rpx_fp_add(a.hi,a.lo);
    if(!isfinite(v)) { *error=DOUBLE_OVERFLOW; return v; }
    if(!isfinite(a.hi)||!isfinite(a.lo)||!isfinite(v)||fabs(a.hi)>1e290||fabs(a.lo)>1e290||fabs(v)>1e290) *error=NONFINITE_VALUE;
    return v;
}
static int bounds(int32_t n,int32_t first,int32_t last) { return n>0 && first>=0 && last>=first && last<=n; }

API rpx_numeric(int32_t n,int32_t first,int32_t last,
    const int32_t *ap,const int32_t *ai,const double *av,
    const int32_t *cp,const int32_t *ri,const int32_t *rp,const int32_t *rn,const int32_t *wa,
    double *l,double *d,double *y,double *minimum,int32_t *failure) {
    int32_t k,p,q,j,r,stop,nnz,m;
    double v,dk,lij;
    if(!bounds(n,first,last)||!ap||!ai||!av||!cp||!ri||!rp||!rn||!wa||!l||!d||!y||!minimum||!failure) return BAD_ARGUMENT;
    nnz=cp[n]; m=ap[n];
    if(nnz<0 || m<0 || ap[0]!=0 || rp[0]!=0 || rp[n]!=nnz) return BAD_ARGUMENT;
    for(k=first;k<last;k++) {
        if(ap[k]<0||ap[k+1]<ap[k]||ap[k+1]>m||rp[k]<0||rp[k+1]<rp[k]||rp[k+1]>nnz) return BAD_ARGUMENT;
        for(p=ap[k];p<ap[k+1];p++) { j=ai[p]; if(j<0||j>k) return BAD_ARGUMENT; y[j]=rpx_fp_add(y[j],av[p]); if(!isfinite(y[j])) return DOUBLE_OVERFLOW; }
        dk=y[k]; y[k]=0.0;
        for(q=rp[k];q<rp[k+1];q++) {
            j=rn[q]; if(j<0||j>=k) return BAD_ARGUMENT;
            v=y[j]; y[j]=0.0; stop=wa[q];
            if(cp[j]<0||stop<cp[j]||stop>=cp[j+1]||stop>=nnz) return BAD_ARGUMENT;
            for(p=cp[j];p<stop;p++) { r=ri[p]; if(r<0||r>=n) return BAD_ARGUMENT; y[r]=rpx_fp_mulsub(y[r],l[p],v); if(!isfinite(y[r])) return DOUBLE_OVERFLOW; }
            if(fabs(d[j])<1e-290) { *failure=j; return PIVOT; }
            lij=rpx_fp_div(v,d[j]); if(!isfinite(lij)) return DOUBLE_OVERFLOW;
            dk=rpx_fp_mulsub(dk,v,lij); if(!isfinite(dk)) return DOUBLE_OVERFLOW; l[stop]=lij;
        }
        if(!isfinite(dk)) return DOUBLE_OVERFLOW;
        if(fabs(dk)<1e-290) { *failure=k; return PIVOT; }
        d[k]=dk; if(fabs(dk)<*minimum) *minimum=fabs(dk);
    }
    return OK;
}
API rpx_dd_numeric(int32_t n,int32_t first,int32_t last,
    const int32_t *ap,const int32_t *ai,const dd *av,
    const int32_t *cp,const int32_t *ri,const int32_t *rp,const int32_t *rn,const int32_t *wa,
    dd *l,dd *d,dd *y,double *minimum,int32_t *failure) {
    int32_t k,p,q,j,r,stop,nnz,m; int error=0;
    double bh,bl,piv; dd v,dk,lij,term;
    if(!bounds(n,first,last)||!ap||!ai||!av||!cp||!ri||!rp||!rn||!wa||!l||!d||!y||!minimum||!failure) return BAD_ARGUMENT;
    nnz=cp[n]; m=ap[n];
    if(nnz<0||m<0||ap[0]!=0||rp[0]!=0||rp[n]!=nnz) return BAD_ARGUMENT;
    for(k=first;k<last;k++) {
        if(ap[k]<0||ap[k+1]<ap[k]||ap[k+1]>m||rp[k]<0||rp[k+1]<rp[k]||rp[k+1]>nnz) return BAD_ARGUMENT;
        for(p=ap[k];p<ap[k+1];p++) { j=ai[p]; if(j<0||j>k) return BAD_ARGUMENT; y[j]=add(y[j],av[p]); }
        dk=y[k]; y[k]=from_double(0);
        for(q=rp[k];q<rp[k+1];q++) {
            j=rn[q]; if(j<0||j>=k) return BAD_ARGUMENT;
            v=y[j]; y[j]=from_double(0); stop=wa[q]; split(v.hi,&bh,&bl);
            if(cp[j]<0||stop<cp[j]||stop>=cp[j+1]||stop>=nnz) return BAD_ARGUMENT;
            for(p=cp[j];p<stop;p++) {
                r=ri[p]; if(r<0||r>=n) return BAD_ARGUMENT;
                term=mul_known(l[p],v,bh,bl,&error); if(error) return error; y[r]=sub(y[r],term);
            }
            piv=value(d[j],&error); if(error) return error;
            if(fabs(piv)<1e-290) { *failure=j; return PIVOT; }
            lij=divide(v,d[j],&error); if(error) return error;
            term=mul(v,lij,&error); if(error) return error; dk=sub(dk,term); l[stop]=lij;
        }
        piv=value(dk,&error); if(error) return error;
        if(fabs(piv)<1e-290) { *failure=k; return PIVOT; }
        d[k]=dk; if(fabs(piv)<*minimum) *minimum=fabs(piv);
    }
    return OK;
}
API rpx_backsolve(int32_t n,int32_t phase,int32_t first,int32_t last,
    const int32_t *cp,const int32_t *ri,const double *l,const double *d,const double *rhs,double *x) {
    int32_t j,p,r,nnz; double v;
    if(!bounds(n,first,last)||!cp||!ri||!l||!d||!rhs||!x||phase<0||phase>3) return BAD_ARGUMENT;
    nnz=cp[n];
    if(phase==0) { for(j=first;j<last;j++) x[j]=rhs[j]; return OK; }
    if(phase==2) { for(j=first;j<last;j++) { if(d[j]==0.0) return DOUBLE_ZERO_DIVISOR; x[j]=rpx_fp_div(x[j],d[j]); if(!isfinite(x[j])) return DOUBLE_OVERFLOW; } return OK; }
    if(phase==1) {
        for(j=first;j<last;j++) {
            if(cp[j]<0||cp[j+1]<cp[j]||cp[j+1]>nnz) return BAD_ARGUMENT;
            v=x[j]; for(p=cp[j];p<cp[j+1];p++) { r=ri[p]; if(r<0||r>=n) return BAD_ARGUMENT; x[r]=rpx_fp_mulsub(x[r],l[p],v); if(!isfinite(x[r])) return DOUBLE_OVERFLOW; }
        }
    } else {
        for(j=last-1;j>=first;j--) {
            if(cp[j]<0||cp[j+1]<cp[j]||cp[j+1]>nnz) return BAD_ARGUMENT;
            v=x[j]; for(p=cp[j];p<cp[j+1];p++) { r=ri[p]; if(r<0||r>=n) return BAD_ARGUMENT; v=rpx_fp_mulsub(v,l[p],x[r]); if(!isfinite(v)) return DOUBLE_OVERFLOW; } x[j]=v;
        }
    }
    return OK;
}
API rpx_dd_backsolve(int32_t n,int32_t phase,int32_t first,int32_t last,
    const int32_t *cp,const int32_t *ri,const dd *l,const dd *d,const double *rhs,dd *x) {
    int32_t j,p,r,nnz; int error=0; double bh,bl; dd v,t;
    if(!bounds(n,first,last)||!cp||!ri||!l||!d||!rhs||!x||phase<0||phase>3) return BAD_ARGUMENT;
    nnz=cp[n];
    if(phase==0) { for(j=first;j<last;j++) x[j]=from_double(rhs[j]); return OK; }
    if(phase==2) { for(j=first;j<last;j++) { x[j]=divide(x[j],d[j],&error); if(error) return error; } return OK; }
    if(phase==1) {
        for(j=first;j<last;j++) {
            if(cp[j]<0||cp[j+1]<cp[j]||cp[j+1]>nnz) return BAD_ARGUMENT;
            v=x[j]; split(v.hi,&bh,&bl);
            for(p=cp[j];p<cp[j+1];p++) {
                r=ri[p]; if(r<0||r>=n) return BAD_ARGUMENT;
                t=mul_known(l[p],v,bh,bl,&error); if(error) return error; x[r]=sub(x[r],t);
            }
        }
    } else {
        for(j=last-1;j>=first;j--) {
            if(cp[j]<0||cp[j+1]<cp[j]||cp[j+1]>nnz) return BAD_ARGUMENT;
            v=x[j]; for(p=cp[j];p<cp[j+1];p++) {
                r=ri[p]; if(r<0||r>=n) return BAD_ARGUMENT;
                t=mul(l[p],x[r],&error); if(error) return error; v=sub(v,t);
            } x[j]=v;
        }
    }
    return OK;
}
/* Test interface also verifies DD error codes without involving Excel. */
API rpx_dd_test(int32_t operation,const dd *a,const dd *b,dd *out) {
    int error=0; if(!a||!b||!out) return BAD_ARGUMENT;
    switch(operation) { case 0:*out=add(*a,*b);break;case 1:*out=sub(*a,*b);break;
        case 2:*out=mul(*a,*b,&error);break;case 3:*out=divide(*a,*b,&error);break;default:return BAD_ARGUMENT; }
    return error;
}
#include "rpfem_bulk.h"
