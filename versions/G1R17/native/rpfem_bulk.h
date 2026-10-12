/* G1R16 DLL_3 bulk kernels. Included after the checked DD primitives.
 * SPDX-License-Identifier: LGPL-2.1-or-later
 * VBA owns the input/output arrays; first/last bound cancellation intervals.
 * Matrix layout is VBA column-major, with cone as the fastest dimension.
 */
extern double rpx_fp_sqrt(double);
extern double rpx_fp_dot3(double,double,double,double,double,double);
extern double rpx_fp_addmul(double,double,double);
extern double rpx_fp_addmul3(double,double,double,double);
extern double rpx_fp_mul3(double,double,double);
extern double rpx_fp_quad(double,double,double,double,int);
extern double rpx_fp_norm2(double,double);
extern double rpx_fp_det3(double,double,double);
extern double rpx_fp_ratio_delta(double,double);
extern double rpx_fp_lorentz_det(double,double);
extern double rpx_fp_delta_div2(double,double,double);
extern int32_t rpx_fp_near_cone(double,double);
extern double rpx_fp_addmuldiv(double,double,double,double);
extern double rpx_fp_submuldiv(double,double,double,double);
extern double rpx_fp_jordan_div(double,double,double,double,double,double,double);
extern double rpx_fp_dot2(double,double,double,double);
enum { NT_INTERIOR=301,NT_CORRECTOR=302,NT_POWER=303 };
/* VBA's ^ operator uses legacy MSVCRT pow. UCRT pow differs by an ulp
 * on some inputs; matching it is necessary near the cone boundary. */
typedef double (__cdecl *legacy_power)(double,double);
static legacy_power nt_power;
API rpx_bulk_ready(void) {
    HMODULE module;
    if(nt_power)return OK;
    module=LoadLibraryExW(L"msvcrt.dll",NULL,LOAD_LIBRARY_SEARCH_SYSTEM32);
    if(!module)return 306;
    nt_power=(legacy_power)(void*)GetProcAddress(module,"pow");
    if(!nt_power){FreeLibrary(module);return 306;}
    return OK;
}
static int valid_double(double v) { return isfinite(v); }
static int normal_power(const double *a,double power,double *out) {
    double rad=rpx_fp_norm2(a[1],a[2]),plus,minus,factor;
    if(!valid_double(rad))return DOUBLE_OVERFLOW;
    if(a[0]<=rad)return NT_INTERIOR;
    if(rpx_bulk_ready()!=OK)return 306;
    plus=nt_power(rpx_fp_add(a[0],rad),power);minus=nt_power(rpx_fp_sub(a[0],rad),power);
    if(!valid_double(plus)||!valid_double(minus))return DOUBLE_OVERFLOW;
    out[0]=rpx_fp_adddiv(plus,minus,2.0);out[1]=0;out[2]=0;
    if(rad>1e-30){factor=rpx_fp_delta_div2(plus,minus,rad);out[1]=rpx_fp_mul(a[1],factor);out[2]=rpx_fp_mul(a[2],factor);}
    return OK;
}
static int normal_quad(const double *a,double *out) {
    int i,j;double rad=rpx_fp_norm2(a[1],a[2]),det=rpx_fp_lorentz_det(a[0],rad);
    if(!isfinite(rad)||!isfinite(det))return DOUBLE_OVERFLOW;
    for(i=0;i<3;i++)for(j=0;j<3;j++){out[i+3*j]=rpx_fp_mul3(2,a[i],a[j]);if(!isfinite(out[i+3*j]))return DOUBLE_OVERFLOW;}
    for(i=0;i<3;i++){out[i+3*i]=(i==0)?rpx_fp_sub(out[i+3*i],det):rpx_fp_add(out[i+3*i],det);if(!isfinite(out[i+3*i]))return DOUBLE_OVERFLOW;}return OK;
}
static int normal_mv(const double *m,const double *a,double *out) {
    int i,j;double v;
    for(i=0;i<3;i++){v=0;for(j=0;j<3;j++){v=rpx_fp_addmul(v,m[i+3*j],a[j]);if(!isfinite(v))return DOUBLE_OVERFLOW;}out[i]=v;}return OK;
}
static int normal_nt(const double *s,const double *z,double *h,double *wmat,double *winv,double *si,double *v) {
    double a[3],b[3],w[3],qs[9];int e;
    e=normal_power(s,.5,a);if(e)return e;e=normal_quad(a,qs);if(e)return e;e=normal_mv(qs,z,b);if(e)return e;
    e=normal_power(b,-.5,a);if(e)return e;e=normal_mv(qs,a,w);if(e)return e;
    e=normal_power(w,-1,a);if(e)return e;e=normal_quad(a,h);if(e)return e;
    e=normal_power(w,.5,a);if(e)return e;e=normal_quad(a,wmat);if(e)return e;
    e=normal_power(w,-.5,a);if(e)return e;e=normal_quad(a,winv);if(e)return e;
    e=normal_power(s,-1,si);if(e)return e;return normal_mv(winv,s,v);
}
static dd dd_sqrt(dd a,int *e) {
    double v=value(a,e);dd q,t;int i;
    if(*e)return from_double(0);
    if(v<0){*e=NT_INTERIOR;return from_double(0);}if(v==0)return from_double(0);
    q=from_double(rpx_fp_sqrt(v));
    for(i=0;i<2;i++){t=divide(a,q,e);if(*e)return q;t=add(q,t);q=mul(from_double(.5),t,e);if(*e)return q;}
    return q;
}
static dd dd_det(const dd *a,int *e) {
    dd d=mul(a[0],a[0],e),t;if(*e)return d;
    t=mul(a[1],a[1],e);if(*e)return d;d=sub(d,t);
    t=mul(a[2],a[2],e);return sub(d,t);
}
static int dd_power(const dd *a,double power,dd *out) {
    int e=0;dd determinant=dd_det(a,&e),rad,plus,minus,factor,t;
    if(e)return e;
    if(value(a[0],&e)<=0||value(determinant,&e)<=0)return e?e:NT_INTERIOR;
    rad=mul(a[1],a[1],&e);if(e)return e;t=mul(a[2],a[2],&e);if(e)return e;rad=dd_sqrt(add(rad,t),&e);if(e)return e;
    plus=add(a[0],rad);minus=divide(determinant,plus,&e);if(e)return e;
    if(power==.5){plus=dd_sqrt(plus,&e);if(e)return e;minus=dd_sqrt(minus,&e);}
    else if(power==-.5){t=dd_sqrt(plus,&e);if(e)return e;plus=divide(from_double(1),t,&e);if(e)return e;t=dd_sqrt(minus,&e);if(e)return e;minus=divide(from_double(1),t,&e);}
    else if(power==-1){plus=divide(from_double(1),plus,&e);if(e)return e;minus=divide(from_double(1),minus,&e);}
    else return NT_POWER;
    if(e)return e;
    out[0]=mul(from_double(.5),add(plus,minus),&e);if(e)return e;out[1]=from_double(0);out[2]=from_double(0);
    if(value(rad,&e)>1e-30){t=mul(from_double(2),rad,&e);if(e)return e;factor=divide(sub(plus,minus),t,&e);if(e)return e;out[1]=mul(a[1],factor,&e);if(e)return e;out[2]=mul(a[2],factor,&e);}
    return e;
}
static int dd_quad(const dd *a,dd *m) {
    int i,j,e=0;dd d=dd_det(a,&e),t;if(e)return e;
    for(i=0;i<3;i++)for(j=0;j<3;j++){t=mul(a[i],a[j],&e);if(e)return e;m[i+3*j]=mul(from_double(2),t,&e);if(e)return e;}
    for(i=0;i<3;i++)m[i+3*i]=(i==0)?sub(m[i+3*i],d):add(m[i+3*i],d);
    return OK;
}
static int dd_mv(const dd *m,const dd *a,dd *out) {
    int i,j,e=0;dd v,t;for(i=0;i<3;i++){v=from_double(0);for(j=0;j<3;j++){t=mul(m[i+3*j],a[j],&e);if(e)return e;v=add(v,t);}out[i]=v;}return e;
}
static int precise_nt(const double *s,const double *z,double *h,double *wm,double *wi,double *si,double *v,dd *hd) {
    dd sd[3],zd[3],a[3],b[3],w[3],qs[9],m[9];int i,e=0;
    for(i=0;i<3;i++){sd[i]=from_double(s[i]);zd[i]=from_double(z[i]);}
    e=dd_power(sd,.5,a);if(e)return e;e=dd_quad(a,qs);if(e)return e;e=dd_mv(qs,zd,b);if(e)return e;
    e=dd_power(b,-.5,a);if(e)return e;e=dd_mv(qs,a,w);if(e)return e;
    e=dd_power(w,-1,a);if(e)return e;e=dd_quad(a,m);if(e)return e;for(i=0;i<9;i++){hd[i]=m[i];h[i]=value(m[i],&e);}if(e)return e;
    e=dd_power(w,.5,a);if(e)return e;e=dd_quad(a,m);if(e)return e;for(i=0;i<9;i++)wm[i]=value(m[i],&e);if(e)return e;
    e=dd_power(w,-.5,a);if(e)return e;e=dd_quad(a,m);if(e)return e;for(i=0;i<9;i++)wi[i]=value(m[i],&e);if(e)return e;
    e=dd_mv(m,sd,b);if(e)return e;e=dd_power(sd,-1,a);if(e)return e;for(i=0;i<3;i++){si[i]=value(a[i],&e);v[i]=value(b[i],&e);}return e;
}
API rpx_nt_test(int32_t precise,const double *s,const double *z,double *h,double *wm,double *wi,double *si,double *v,dd *hd) {
    if(!s||!z||!h||!wm||!wi||!si||!v||!hd)return BAD_ARGUMENT;
    return precise?precise_nt(s,z,h,wm,wi,si,v,hd):normal_nt(s,z,h,wm,wi,si,v);
}
API rpx_nt_scalar_test(int32_t operation,const double *a,double power,double *out) {
    if(!a||!out)return BAD_ARGUMENT;
    if(operation==0)return normal_power(a,power,out);
    if(operation==1)return normal_quad(a,out);
    if(operation==2){out[0]=rpx_fp_sqrt(a[0]);return OK;}
    if(operation==3){if(rpx_bulk_ready()!=OK)return 306;out[0]=rpx_fp_norm2(a[1],a[2]);out[1]=nt_power(rpx_fp_add(a[0],out[0]),power);out[2]=nt_power(rpx_fp_sub(a[0],out[0]),power);out[3]=rpx_fp_delta_div2(out[1],out[2],out[0]);return OK;}
    return BAD_ARGUMENT;
}
API rpx_nt_trace(const double *s,const double *z,double *out) {
    double a[3],b[3],w[3],qs[9];int e;
    e=normal_power(s,.5,a);if(e)return e;memcpy(out,a,24);out+=3;
    normal_quad(a,qs);memcpy(out,qs,72);out+=9;
    normal_mv(qs,z,b);memcpy(out,b,24);out+=3;
    e=normal_power(b,-.5,a);if(e)return e;memcpy(out,a,24);out+=3;
    normal_mv(qs,a,w);memcpy(out,w,24);out+=3;
    e=normal_power(w,-1,a);if(e)return e;memcpy(out,a,24);out+=3;
    normal_quad(a,qs);memcpy(out,qs,72);out+=9;
    e=normal_power(w,.5,a);if(e)return e;memcpy(out,a,24);out+=3;
    normal_quad(a,qs);memcpy(out,qs,72);return OK;
}
API rpx_scaling(int32_t nb,int32_t ns,int32_t first,int32_t last,int32_t precise,
    const int32_t *dims,const int32_t *starts,const double *slack,const double *dual,
    double *hh,double *wm,double *wi,double *si,double *v,int16_t *flags,dd *hd,int32_t *precise_count) {
    int32_t k,p,i,j;int e;double s[3],z[3],h[9],w[9],inv[9],iv[3],vv[3],rs,rz;dd dh[9];
    if(!bounds(nb,first,last)||ns<1||!dims||!starts||!slack||!dual||!hh||!wm||!wi||!si||!v||!flags||!hd||!precise_count)return BAD_ARGUMENT;
    for(k=first;k<last;k++){
        p=starts[k];if(p<0||(dims[k]!=1&&dims[k]!=3)||p>ns-dims[k])return BAD_ARGUMENT;
        if(dims[k]==1){if(slack[p]==0)return DOUBLE_ZERO_DIVISOR;hh[k]=rpx_fp_div(dual[p],slack[p]);si[p]=rpx_fp_div(1,slack[p]);if(!isfinite(hh[k])||!isfinite(si[p]))return DOUBLE_OVERFLOW;continue;}
        for(i=0;i<3;i++){s[i]=slack[p+i];z[i]=dual[p+i];}
        if(precise){rs=rpx_fp_norm2(s[1],s[2]);rz=rpx_fp_norm2(z[1],z[2]);if(!isfinite(rs)||!isfinite(rz))return DOUBLE_OVERFLOW;if(s[0]<=rs||z[0]<=rz)return 304;
            flags[k]=(rpx_fp_near_cone(s[0],rs)||rpx_fp_near_cone(z[0],rz))?-1:0;}
        if(precise&&flags[k]){e=precise_nt(s,z,h,w,inv,iv,vv,dh);if(e)return e==NT_INTERIOR?305:e;(*precise_count)++;}
        else{e=normal_nt(s,z,h,w,inv,iv,vv);if(e)return e;}
        for(i=0;i<3;i++){si[p+i]=iv[i];v[p+i]=vv[i];for(j=0;j<3;j++){
            hh[k+nb*(i+3*j)]=h[i+3*j];wm[k+nb*(i+3*j)]=w[i+3*j];wi[k+nb*(i+3*j)]=inv[i+3*j];
            if(precise&&flags[k])hd[k+nb*(i+3*j)]=dh[i+3*j];}}
    }
    return OK;
}
/* Stable packed cones: coefficients are grouped by cone, row, original entry.
 * ids/coefficients/slots remain unchanged throughout one compiled solve.
 */
API rpx_gproduct(int32_t nb,int32_t ns,int32_t nx,int32_t first,int32_t last,
 const int32_t *dims,const int32_t *starts,const int32_t *ip,const int32_t *cp0,const int32_t *ids,const double *coef,
 const double *offset,const double *x,double *out,int32_t with_offset) {
    int32_t k,r,j,count,p,id;double value0;
    if(!bounds(nb,first,last)||ns<1||nx<1||!dims||!starts||!ip||!cp0||!ids||!coef||!offset||!x||!out)return BAD_ARGUMENT;
    for(k=first;k<last;k++){count=ip[k+1]-ip[k];if(count<0||starts[k]<0||starts[k]>ns-dims[k])return BAD_ARGUMENT;
        for(r=0;r<dims[k];r++){p=starts[k]+r;value0=with_offset?offset[p]:0;
            for(j=0;j<count;j++){id=ids[ip[k]+j];if(id<0||id>=nx)return BAD_ARGUMENT;value0=rpx_fp_addmul(value0,coef[cp0[k]+r*count+j],x[id]);if(!isfinite(value0))return DOUBLE_OVERFLOW;}out[p]=value0;}}
    return OK;
}
API rpx_gtadd(int32_t nb,int32_t ns,int32_t nx,int32_t first,int32_t last,
 const int32_t *dims,const int32_t *starts,const int32_t *ip,const int32_t *cp0,const int32_t *ids,const double *coef,
 const double *x,double *out,double factor) {
    int32_t k,r,j,count,id;double v0;
    if(!bounds(nb,first,last)||ns<1||nx<1||!dims||!starts||!ip||!cp0||!ids||!coef||!x||!out)return BAD_ARGUMENT;
    for(k=first;k<last;k++){count=ip[k+1]-ip[k];for(j=0;j<count;j++){v0=0;id=ids[ip[k]+j];if(id<0||id>=nx)return BAD_ARGUMENT;
        for(r=0;r<dims[k];r++)v0=rpx_fp_addmul(v0,coef[cp0[k]+r*count+j],x[starts[k]+r]);out[id]=rpx_fp_addmul(out[id],factor,v0);if(!isfinite(out[id]))return DOUBLE_OVERFLOW;}}
    return OK;
}
API rpx_kkt_cones(int32_t nb,int32_t nnz,int32_t first,int32_t last,
 const int32_t *dims,const int32_t *ip,const int32_t *cp0,const int32_t *sp,const int32_t *slot,
 const double *coef,const double *hh,double *values) {
    int32_t i,j,k,cnt,at,p;double c0,c1,c2,w0,w1,w2,v0;
    if(!bounds(nb,first,last)||nnz<1||!dims||!ip||!cp0||!sp||!slot||!coef||!hh||!values)return BAD_ARGUMENT;
    for(i=first;i<last;i++){cnt=ip[i+1]-ip[i];at=sp[i];for(j=0;j<cnt;j++){
        c0=coef[cp0[i]+j];if(dims[i]==3){c1=coef[cp0[i]+cnt+j];c2=coef[cp0[i]+2*cnt+j];
            w0=rpx_fp_dot3(hh[i],c0,hh[i+3*nb],c1,hh[i+6*nb],c2);
            w1=rpx_fp_dot3(hh[i+nb],c0,hh[i+4*nb],c1,hh[i+7*nb],c2);
            w2=rpx_fp_dot3(hh[i+2*nb],c0,hh[i+5*nb],c1,hh[i+8*nb],c2);
        }else{w0=0;w1=0;w2=0;}
        for(k=0;k<=j;k++){p=slot[at++];if(p<0||p>=nnz)return BAD_ARGUMENT;
            if(dims[i]==1)values[p]=rpx_fp_addmul3(values[p],c0,hh[i],coef[cp0[i]+k]);
            else{v0=rpx_fp_dot3(coef[cp0[i]+k],w0,coef[cp0[i]+cnt+k],w1,coef[cp0[i]+2*cnt+k],w2);values[p]=rpx_fp_add(values[p],v0);}
            if(!isfinite(values[p]))return DOUBLE_OVERFLOW;
        }
    }}return OK;
}
typedef struct { double hi,lo,sh,sl; } ddcoef;
API rpx_corrector(int32_t nb,int32_t ns,int32_t phase,int32_t first,int32_t last,const int32_t *dims,const int32_t *starts,double target,const uintptr_t *args) {
    const double *dual,*sinv,*v,*ww,*wi,*ds,*dz,*slack;double *rc;
    int32_t k,p,i,j;int e;double wm[9],inv[9],vv[3],a[3],b[3],c[3],d[3],out[3],det;
    if(nb<1||ns<1||!dims||!starts||!args)return BAD_ARGUMENT;
    for(i=0;i<9;i++)if(!args[i])return BAD_ARGUMENT;
    dual=(const double*)args[0];sinv=(const double*)args[1];v=(const double*)args[2];ww=(const double*)args[3];wi=(const double*)args[4];ds=(const double*)args[5];dz=(const double*)args[6];slack=(const double*)args[7];rc=(double*)args[8];
    if(phase==0){if(!bounds(ns,first,last))return BAD_ARGUMENT;for(i=first;i<last;i++){rc[i]=rpx_fp_mulsub(dual[i],target,sinv[i]);if(!isfinite(rc[i]))return DOUBLE_OVERFLOW;}return OK;}
    if(phase!=1||!bounds(nb,first,last))return BAD_ARGUMENT;
    for(k=first;k<last;k++){p=starts[k];if(p<0||(dims[k]!=1&&dims[k]!=3)||p>ns-dims[k])return BAD_ARGUMENT;
        if(dims[k]==1){if(slack[p]==0)return DOUBLE_ZERO_DIVISOR;rc[p]=rpx_fp_addmuldiv(rc[p],ds[p],dz[p],slack[p]);if(!isfinite(rc[p]))return DOUBLE_OVERFLOW;}
        else{for(i=0;i<3;i++){vv[i]=v[p+i];for(j=0;j<3;j++){wm[i+3*j]=ww[k+nb*(i+3*j)];inv[i+3*j]=wi[k+nb*(i+3*j)];}}
            e=normal_mv(inv,ds+p,a);if(e)return e;e=normal_mv(wm,dz+p,b);if(e)return e;
            c[0]=rpx_fp_dot3(a[0],b[0],a[1],b[1],a[2],b[2]);c[1]=rpx_fp_dot2(a[0],b[1],b[0],a[1]);c[2]=rpx_fp_dot2(a[0],b[2],b[0],a[2]);
            for(i=0;i<3;i++)if(!isfinite(c[i]))return DOUBLE_OVERFLOW;
            det=rpx_fp_det3(vv[0],vv[1],vv[2]);if(!isfinite(det))return DOUBLE_OVERFLOW;if(det<=0)return NT_CORRECTOR;
            d[0]=rpx_fp_jordan_div(vv[0],c[0],vv[1],c[1],vv[2],c[2],det);
            if(vv[0]==0)return DOUBLE_ZERO_DIVISOR;
            d[1]=rpx_fp_submuldiv(c[1],vv[1],d[0],vv[0]);d[2]=rpx_fp_submuldiv(c[2],vv[2],d[0],vv[0]);for(i=0;i<3;i++)if(!isfinite(d[i]))return DOUBLE_OVERFLOW;e=normal_mv(inv,d,out);if(e)return e;
            for(i=0;i<3;i++){rc[p+i]=rpx_fp_add(rc[p+i],out[i]);if(!isfinite(rc[p+i]))return DOUBLE_OVERFLOW;}
        }
    }return OK;
}
API rpx_kkt_init(int32_t count,int32_t first,int32_t last,const double *base,double *out) {
    if(!bounds(count,first,last)||!base||!out)return BAD_ARGUMENT;
    memcpy(out+first,base+first,(size_t)(last-first)*8);return OK;
}
API rpx_kkt_diagonal(int32_t n,int32_t nx,int32_t count,int32_t first,int32_t last,const int32_t *diagonal,const int32_t *permutation,double regularization,double *out,double *scales) {
    int32_t i,p;double d;
    if(!bounds(n,first,last)||nx<1||nx>n||count<1||!diagonal||!permutation||!out||!scales)return BAD_ARGUMENT;
    for(i=first;i<last;i++){p=diagonal[i];if(p<0||p>=count||permutation[i]<0||permutation[i]>=n)return BAD_ARGUMENT;
        out[p]=permutation[i]<nx?rpx_fp_add(out[p],regularization):-regularization;if(!isfinite(out[p]))return DOUBLE_OVERFLOW;
        scales[i]=1;d=fabs(out[p]);if(permutation[i]<nx&&d>1e-20)scales[i]=rpx_fp_div(1,rpx_fp_sqrt(d));
    }return OK;
}
API rpx_kkt_scale(int32_t n,int32_t count,int32_t first,int32_t last,const int32_t *ap,const int32_t *ai,const double *scales,double *out,double *original) {
    int32_t i,j,p;
    if(!bounds(n,first,last)||count<1||!ap||!ai||!scales||!out||!original||ap[0]!=0||ap[n]!=count)return BAD_ARGUMENT;
    for(i=first;i<last;i++){if(ap[i]<0||ap[i+1]<ap[i]||ap[i+1]>count)return BAD_ARGUMENT;
        for(p=ap[i];p<ap[i+1];p++){j=ai[p];if(j<0||j>=n)return BAD_ARGUMENT;original[p]=out[p];out[p]=rpx_fp_mul3(out[p],scales[i],scales[j]);if(!isfinite(out[p]))return DOUBLE_OVERFLOW;}
    }return OK;
}
_Static_assert(sizeof(ddcoef)==32,"VBA coefficient layout");
static dd bulk_coef_product(ddcoef a,dd b,int *e) {
    double bh,bl,p;dd out,left={a.hi,a.lo};
    p=rpx_fp_mul(a.hi,b.hi);
    if(!isfinite(p)){*e=DOUBLE_OVERFLOW;return from_double(0);}
    if(fabs(p)>1e290){*e=NONFINITE_PRODUCT;return from_double(0);}
    if(p==0&&a.hi!=0&&b.hi!=0){*e=PRODUCT_UNDERFLOW;return from_double(0);}
    split(b.hi,&bh,&bl);rpx_fp_dd_product_parts(&left,&b,a.sh,a.sl,bh,bl,p,&out);return out;
}
static dd bulk_right_product(dd a,ddcoef b,int *e) {
    ddcoef left;split(a.hi,&left.sh,&left.sl);left.hi=a.hi;left.lo=a.lo;
    /* Split(b.hi) is exactly b.sh/b.sl; term order is the original HDMul. */
    return bulk_coef_product(left,(dd){b.hi,b.lo},e);
}
API rpx_hoperator(int32_t n,int32_t nx,int32_t nb,int32_t ns,int32_t phase,int32_t first,int32_t last,int32_t unshifted,const uintptr_t *args) {
    const int32_t *perm,*dims,*starts,*ip,*ids,*coef_at,*ep,*ei;
    const ddcoef *scale,*coef0,*hh,*eval,*reg_pos,*reg_neg;const dd *z;
    dd *result,*xx0,*vv,*gx,*hg,t;int32_t i,j,k,r,p,base;int e=0;
    if(n<1||nx<1||nx>n||nb<1||ns<1||!args)return BAD_ARGUMENT;
    for(i=0;i<20;i++)if(!args[i])return BAD_ARGUMENT;
    perm=(const int32_t*)args[0];scale=(const ddcoef*)args[1];dims=(const int32_t*)args[2];starts=(const int32_t*)args[3];ip=(const int32_t*)args[4];ids=(const int32_t*)args[5];
    coef_at=(const int32_t*)args[6];coef0=(const ddcoef*)args[7];hh=(const ddcoef*)args[8];ep=(const int32_t*)args[9];ei=(const int32_t*)args[10];eval=(const ddcoef*)args[11];
    reg_pos=(const ddcoef*)args[12];reg_neg=(const ddcoef*)args[13];z=(const dd*)args[14];result=(dd*)args[15];xx0=(dd*)args[16];vv=(dd*)args[17];gx=(dd*)args[18];hg=(dd*)args[19];
    if(phase==-1){memset(xx0,0,(size_t)n*sizeof(dd));memset(vv,0,(size_t)n*sizeof(dd));memset(gx,0,(size_t)ns*sizeof(dd));memset(hg,0,(size_t)ns*sizeof(dd));return OK;}
    if(phase==0){if(!bounds(n,first,last))return BAD_ARGUMENT;for(i=first;i<last;i++){j=perm[i];if(j<0||j>=n)return BAD_ARGUMENT;xx0[j]=bulk_right_product(z[i],scale[i],&e);if(e)return e;}return OK;}
    if(phase==1){if(!bounds(nb,first,last))return BAD_ARGUMENT;for(k=first;k<last;k++){
        p=starts[k];if((dims[k]!=1&&dims[k]!=3)||p<0||p>ns-dims[k])return BAD_ARGUMENT;
        for(r=0;r<dims[k];r++){base=coef_at[k]+r*(ip[k+1]-ip[k]);for(j=0;j<ip[k+1]-ip[k];j++){
            i=ids[ip[k]+j];if(i<0||i>=nx)return BAD_ARGUMENT;t=bulk_coef_product(coef0[base+j],xx0[i],&e);if(e)return e;gx[p+r]=add(gx[p+r],t);}}
        for(r=0;r<dims[k];r++){
            for(j=0;j<dims[k];j++){t=bulk_coef_product(hh[k+nb*(r+3*j)],gx[p+j],&e);if(e)return e;hg[p+r]=add(hg[p+r],t);}
            base=coef_at[k]+r*(ip[k+1]-ip[k]);for(j=0;j<ip[k+1]-ip[k];j++){i=ids[ip[k]+j];t=bulk_coef_product(coef0[base+j],hg[p+r],&e);if(e)return e;vv[i]=add(vv[i],t);}
        }
    }return OK;}
    if(phase==2){if(!bounds(n-nx,first,last))return BAD_ARGUMENT;for(i=first;i<last;i++)for(p=ep[i];p<ep[i+1];p++){
        j=ei[p];if(j<0||j>=nx)return BAD_ARGUMENT;t=bulk_coef_product(eval[p],xx0[j],&e);if(e)return e;vv[nx+i]=add(vv[nx+i],t);
        t=bulk_coef_product(eval[p],xx0[nx+i],&e);if(e)return e;vv[j]=add(vv[j],t);
    }return OK;}
    if(phase==3){if(!bounds(n,first,last))return BAD_ARGUMENT;for(i=first;i<last;i++){
        j=perm[i];if(j<0||j>=n)return BAD_ARGUMENT;if(!unshifted){t=bulk_coef_product(j<nx?*reg_pos:*reg_neg,xx0[j],&e);if(e)return e;vv[j]=add(vv[j],t);}
        result[i]=bulk_coef_product(scale[i],vv[j],&e);if(e)return e;
    }return OK;}return BAD_ARGUMENT;
}
