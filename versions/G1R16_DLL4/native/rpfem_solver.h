/* DLL4 remaining solver vector loops. Same stored-Double and DD boundaries.
 * Pointers are borrowed for the duration of this call only. LGPL-2.1-or-later. */
#define VP(T,I) ((T*)a[I])
extern double rpx_fp_bilinear3(double,double,double,double,double,double);
extern double rpx_fp_abs6(double,double,double,double,double,double);
extern double rpx_fp_discriminant(double,double,double);
extern int32_t rpx_fp_boundary_tiny(double,double);
API rpx_solver_vectors(int32_t op,int32_t n,int32_t first,int32_t last,const uintptr_t *a,double alpha,double *m) {
 int32_t i,p,r;int e=0;dd t;double v,error,ratio;
 if(!bounds(n,first,last)||!a||!m)return BAD_ARGUMENT;
 {int32_t required=0;
  switch(op){case 1:required=4;break;case 2:required=5;break;case 3:case 4:required=alpha!=0?3:2;break;case 5:required=6;break;case 6:required=5;break;case 7:required=13;break;case 8:required=9;break;case 9:required=2;break;case 10:required=3;break;case 11:required=5;break;case 12:case 13:case 15:required=2;break;case 14:required=5;break;case 16:case 17:required=4;break;case 18:required=7;break;case 19:required=9;break;default:return BAD_ARGUMENT;}
  for(i=0;i<required;i++)if(!a[i]&&!(i==1&&(op==3||op==4)&&alpha!=0))return BAD_ARGUMENT;
 }
 if(op==1){ /* HLinear: residual and maximum scaled/original error */
  const dd *b=VP(dd,0),*product=VP(dd,1);const double *scale=VP(double,2);double *res=VP(double,3);
  for(i=first;i<last;i++){t=sub(b[i],product[i]);res[i]=value(t,&e);if(e)return e;
   error=fabs(res[i]);if(error>m[0])m[0]=error;if(scale[i]==0)return DOUBLE_ZERO_DIVISOR;
   error=fabs(rpx_fp_div(res[i],scale[i]));if(!isfinite(error))return DOUBLE_OVERFLOW;if(error>m[0])m[0]=error;}return OK;
 }
 if(op==2){ /* HLinear: keep S*b low component */
  const double *rhs=VP(double,0);const int32_t *perm=VP(int32_t,1);const ddcoef *scale=VP(ddcoef,2);dd *b=VP(dd,3);double *scaled=VP(double,4);
  for(i=first;i<last;i++){p=perm[i];if(p<0||p>=n)return BAD_ARGUMENT;if(!isfinite(rhs[p])||fabs(rhs[p])>1e290)return 322;
   b[i]=bulk_right_product(from_double(rhs[p]),scale[i],&e);if(e)return e;scaled[i]=value(b[i],&e);if(e)return e;}return OK;
 }
 if(op==3||op==4){ /* correction / initial Double to DD */
  dd *x=VP(dd,0);const double *y=VP(double,1);const dd *yd=VP(dd,2);
  for(i=first;i<last;i++){t=(alpha!=0)?yd[i]:from_double(y[i]);x[i]=(op==3)?add(x[i],t):t;}return OK;
 }
 if(op==5){ /* final unscale/unpermute and extraction */
  const dd *best=VP(dd,0);const ddcoef *scale=VP(ddcoef,1);const int32_t *perm=VP(int32_t,2);dd *out=VP(dd,3);double *answer=VP(double,4),*stored=VP(double,5);
  for(i=first;i<last;i++){p=perm[i];if(p<0||p>=n)return BAD_ARGUMENT;stored[i]=value(best[i],&e);if(e)return e;
   out[p]=bulk_right_product(best[i],scale[i],&e);if(e)return e;answer[p]=value(out[p],&e);if(e)return e;}return OK;
 }
 if(op==6){ /* backward diagnostic input conversions */
  const dd *scaled=VP(dd,0);const ddcoef *scale=VP(ddcoef,1);const double *x=VP(double,2);dd *ox=VP(dd,3),*round=VP(dd,4);
  for(i=first;i<last;i++){ox[i]=bulk_right_product(scaled[i],scale[i],&e);if(e)return e;round[i]=from_double(x[i]);}return OK;
 }
 if(op==7){ /* all three diagnostic sparse products, existing row-entry order */
  const int32_t *ap=VP(int32_t,0),*ai=VP(int32_t,1);const ddcoef *val=VP(ddcoef,2),*orig=VP(ddcoef,3);
  const dd *xs=VP(dd,4),*xo=VP(dd,5),*xr=VP(dd,6);dd *p1=VP(dd,7),*p2=VP(dd,8),*p3=VP(dd,9);
  const double *av=VP(double,10),*x=VP(double,11);double *den=VP(double,12);
  for(i=first;i<last;i++)for(p=ap[i];p<ap[i+1];p++){r=ai[p];if(r<0||r>i)return BAD_ARGUMENT;
   t=bulk_coef_product(val[p],xs[i],&e);if(e)return e;p1[r]=add(p1[r],t);
   if(r!=i){t=bulk_coef_product(val[p],xs[r],&e);if(e)return e;p1[i]=add(p1[i],t);}
   t=bulk_coef_product(orig[p],xo[i],&e);if(e)return e;p2[r]=add(p2[r],t);
   if(r!=i){t=bulk_coef_product(orig[p],xo[r],&e);if(e)return e;p2[i]=add(p2[i],t);}
   t=bulk_coef_product(val[p],xr[i],&e);if(e)return e;p3[r]=add(p3[r],t);
   if(r!=i){t=bulk_coef_product(val[p],xr[r],&e);if(e)return e;p3[i]=add(p3[i],t);}
   den[r]=rpx_fp_addmul(den[r],fabs(av[p]),fabs(x[i]));if(r!=i)den[i]=rpx_fp_addmul(den[i],fabs(av[p]),fabs(x[r]));
  }return OK;
 }
 if(op==8){ /* HBackward reductions: metrics persist across slices */
  const double *row=VP(double,0),*x=VP(double,1),*b=VP(double,2),*rhs=VP(double,3);const int32_t *perm=VP(int32_t,4);
  const dd *p1=VP(dd,5),*p2=VP(dd,6);double *den=VP(double,8);
  for(i=first;i<last;i++){
   if(row[i]>m[0])m[0]=row[i];if(fabs(x[i])>m[1])m[1]=fabs(x[i]);if(fabs(b[i])>m[2])m[2]=fabs(b[i]);
   t=sub(from_double(b[i]),p1[i]);error=fabs(value(t,&e));if(e)return e;den[i]=rpx_fp_add(den[i],fabs(b[i]));
   if(error>m[3]){m[3]=error;m[4]=(double)i;}ratio=0;
   if(den[i]>0)ratio=rpx_fp_div(error,den[i]);else if(error>0)ratio=1e300;if(ratio>m[5])m[5]=ratio;
   p=perm[i];if(p<0||p>=n)return BAD_ARGUMENT;t=sub(from_double(rhs[p]),p2[i]);v=fabs(value(t,&e));if(e)return e;if(v>m[6])m[6]=v;
  }return OK;
 }
 if(op==9){const double *b=VP(double,0);const dd *p3=VP(dd,1);for(i=first;i<last;i++){t=sub(from_double(b[i]),p3[i]);v=fabs(value(t,&e));if(e)return e;if(v>m[0])m[0]=v;}return OK;}
 if(op==10){ /* candidate HSD update; live state remains in VBA until audited */
  const double *old=VP(double,0);const dd *d=VP(dd,1);double *out=VP(double,2);
  for(i=first;i<last;i++){t=mul(from_double(alpha),d[i],&e);if(e)return e;out[i]=value(add(from_double(old[i]),t),&e);if(e)return e;if(!isfinite(out[i])||fabs(out[i])>1e290)return 321;}return OK;
 }
 if(op==11){ /* symmetric Double product */
  const int32_t *ap=VP(int32_t,0),*ai=VP(int32_t,1);const double *av=VP(double,2),*x=VP(double,3);double *y=VP(double,4);
  for(i=first;i<last;i++)for(p=ap[i];p<ap[i+1];p++){r=ai[p];if(r<0||r>i)return BAD_ARGUMENT;
   y[r]=rpx_fp_addmul(y[r],av[p],x[i]);if(!isfinite(y[r]))return DOUBLE_OVERFLOW;
   if(r!=i){y[i]=rpx_fp_addmul(y[i],av[p],x[r]);if(!isfinite(y[i]))return DOUBLE_OVERFLOW;}}
  return OK;
 }
 if(op==12){ /* normal refinement residual */
  const double *rhs=VP(double,0);double *res=VP(double,1);
  for(i=first;i<last;i++){res[i]=rpx_fp_sub(rhs[i],res[i]);if(!isfinite(res[i]))return DOUBLE_OVERFLOW;if(fabs(res[i])>m[0])m[0]=fabs(res[i]);}return OK;
 }
 if(op==13){double *x=VP(double,0);const double *delta=VP(double,1);for(i=first;i<last;i++){x[i]=rpx_fp_add(x[i],delta[i]);if(!isfinite(x[i]))return DOUBLE_OVERFLOW;}return OK;}
 if(op==14){ /* unpack normal direction */
  const double *answer=VP(double,0),*scale=VP(double,1);const int32_t *perm=VP(int32_t,2);double *dx=VP(double,3),*dy=VP(double,4);int32_t nx=(int32_t)alpha;
  for(i=first;i<last;i++){p=perm[i];if(p<0||p>=n)return BAD_ARGUMENT;v=rpx_fp_mul(answer[i],scale[i]);if(!isfinite(v))return DOUBLE_OVERFLOW;if(p<nx)dx[p]=v;else dy[p-nx]=v;}return OK;
 }
 if(op==15){ /* normal state update */
  double *x=VP(double,0);const double *d=VP(double,1);for(i=first;i<last;i++){x[i]=rpx_fp_addmul(x[i],alpha,d[i]);if(!isfinite(x[i]))return DOUBLE_OVERFLOW;}return OK;
 }
 if(op==16){ /* step interior, preserving examined count and early exit */
  const int32_t *dims=VP(int32_t,0),*starts=VP(int32_t,1);const double *x=VP(double,2),*d=VP(double,3);double u,w;
  for(i=first;i<last;i++){p=starts[i];m[0]++;v=rpx_fp_addmul(x[p],alpha,d[p]);if(!isfinite(v))return DOUBLE_OVERFLOW;
   if(v<=0){m[1]=0;return OK;}if(dims[i]==3){u=rpx_fp_addmul(x[p+1],alpha,d[p+1]);w=rpx_fp_addmul(x[p+2],alpha,d[p+2]);if(v<=rpx_fp_norm2(u,w)){m[1]=0;return OK;}}
  }return OK;
 }
 if(op==17){ /* ConeBoundaryHit: all roots, branch choices and ties retained */
  const int32_t *dims=VP(int32_t,0),*starts=VP(int32_t,1);const double *x=VP(double,2),*d=VP(double,3);
  double s0,s1,s2,d0,d1,d2,aa,bb,cc,disc,q,root1,root2,cand,t0,xscale;
  for(i=first;i<last;i++){p=starts[i];s0=x[p];d0=d[p];if(s0<=0){m[2]=0;return OK;}
   if(d0<0){cand=rpx_fp_div(-s0,d0);if(cand<=0){m[2]=0;return OK;}if(cand<m[0]){m[0]=cand;m[1]=(double)i;}}
   if(dims[i]==3){s1=x[p+1];s2=x[p+2];d1=d[p+1];d2=d[p+2];cc=rpx_fp_det3(s0,s1,s2);if(!isfinite(cc))return DOUBLE_OVERFLOW;if(cc<=0){m[2]=0;return OK;}
    aa=rpx_fp_det3(d0,d1,d2);bb=rpx_fp_bilinear3(s0,d0,s1,d1,s2,d2);xscale=rpx_fp_abs6(s0,s1,s2,d0,d1,d2);if(xscale<1)xscale=1;
    if(!isfinite(aa)||!isfinite(bb)||!isfinite(xscale))return DOUBLE_OVERFLOW;
    if(rpx_fp_boundary_tiny(aa,xscale)){if(!rpx_fp_boundary_tiny(bb,xscale)){cand=rpx_fp_div(-cc,bb);t0=rpx_fp_addmul(s0,cand,d0);if(cand>0&&t0>=0&&cand<m[0]){m[0]=cand;m[1]=(double)i;}}}
    else{disc=rpx_fp_discriminant(bb,aa,cc);if(!isfinite(disc))return DOUBLE_OVERFLOW;if(disc>=0){disc=rpx_fp_sqrt(disc);q=rpx_fp_adddiv(bb,bb>=0?disc:-disc,-2);
      if(fabs(q)<=1e-300){m[2]=0;return OK;}root1=rpx_fp_div(q,aa);root2=rpx_fp_div(cc,q);
      t0=rpx_fp_addmul(s0,root1,d0);if(root1>0&&t0>=0&&root1<m[0]){m[0]=root1;m[1]=(double)i;}
      t0=rpx_fp_addmul(s0,root2,d0);if(root2>0&&t0>=0&&root2<m[0]){m[0]=root2;m[1]=(double)i;}
    }}
   }
  }return OK;
 }
 if(op==18){ /* HSD border reduction, DD accumulation order retained */
  const double *row=VP(double,0);const int32_t *perm=VP(int32_t,1);const ddcoef *scale=VP(ddcoef,2);const dd *arrow=VP(dd,3),*answer=VP(dd,4);
  dd *den=VP(dd,5),*rhs=VP(dd,6),rowdd;
  for(i=first;i<last;i++){p=perm[i];if(p<0||p>=n)return BAD_ARGUMENT;rowdd=bulk_right_product(from_double(row[p]),scale[i],&e);if(e)return e;
   t=mul(rowdd,arrow[i],&e);if(e)return e;*den=sub(*den,t);m[0]=rpx_fp_add(m[0],fabs(value(t,&e)));if(e)return e;
   t=mul(rowdd,answer[i],&e);if(e)return e;*rhs=sub(*rhs,t);
  }return OK;
 }
 if(op==19){ /* HSD scaled/unscaled direction and Double view */
  dd *scaled=VP(dd,0),*direction=VP(dd,1);const dd *dt=VP(dd,2),*arrow=VP(dd,3),*solution=VP(dd,4);const ddcoef *scale=VP(ddcoef,5);const int32_t *perm=VP(int32_t,6);
  double *dx=VP(double,7),*dy=VP(double,8);int32_t nx=(int32_t)alpha;
  for(i=first;i<last;i++){p=perm[i];if(p<0||p>=n)return BAD_ARGUMENT;t=mul(arrow[i],*dt,&e);if(e)return e;scaled[i]=sub(solution[i],t);
   direction[p]=bulk_right_product(scaled[i],scale[i],&e);if(e)return e;v=value(direction[p],&e);if(e)return e;if(p<nx)dx[p]=v;else dy[p-nx]=v;
  }return OK;
 }
 return BAD_ARGUMENT;
}

API rpx_solver_update(int32_t nx,int32_t nm,int32_t ns,const uintptr_t *a,double alpha) {
 int32_t i,k,total[4]={nx,nm,ns,ns};double v;if(nx<1||nm<0||ns<1||!a)return BAD_ARGUMENT;
 for(k=0;k<8;k++)if(!a[k])return BAD_ARGUMENT;
 for(k=0;k<4;k++){double *x=(double*)a[2*k];const double *d=(const double*)a[2*k+1];
  for(i=0;i<total[k];i++){v=rpx_fp_addmul(x[i],alpha,d[i]);if(!isfinite(v))return DOUBLE_OVERFLOW;x[i]=v;}}
 return OK;
}

extern double rpx_fp_offset_tau(double,double,double);
API rpx_solver_residual(int32_t nb,int32_t ns,int32_t nx,int32_t nm,int32_t phase,int32_t first,int32_t last,const int32_t *dims,const int32_t *starts,const int32_t *ip,const int32_t *cp,const int32_t *ids,const double *coef,const double *offset,const uintptr_t *a,double tau,double *m) {
 int32_t i,p,j,k,r,total;double v;int code;
 if(!a||!m||!dims||!starts||!ip||!cp||!ids||!coef||!offset)return BAD_ARGUMENT;
 for(i=0;i<12;i++)if(!a[i])return BAD_ARGUMENT;
 const double *xx=(const double*)a[0],*xy=(const double*)a[1],*q=(const double*)a[2],*ev=(const double*)a[5],*er=(const double*)a[6],*slack=(const double*)a[7],*dual=(const double*)a[8];
 const int32_t *ep=(const int32_t*)a[3],*ei=(const int32_t*)a[4];double *rg=(double*)a[9],*rd=(double*)a[10],*rp=(double*)a[11];
 if(nb<1||ns<1||nx<1||nm<0)return BAD_ARGUMENT;
 total=nb;if(phase==1||phase==4)total=nx;if(phase==2)total=nm;if(phase==5)total=nm+ns;if(!bounds(total,first,last))return BAD_ARGUMENT;
 if(phase==0){code=rpx_gproduct(nb,ns,nx,first,last,dims,starts,ip,cp,ids,coef,offset,xx,rg,1);if(code)return code;
  for(k=first;k<last;k++)for(r=0;r<dims[k];r++){p=starts[k]+r;rg[p]=rpx_fp_offset_tau(rg[p],tau,offset[p]);rg[p]=rpx_fp_sub(rg[p],slack[p]);if(!isfinite(rg[p]))return DOUBLE_OVERFLOW;}return OK;}
 if(phase==1){for(i=first;i<last;i++){rd[i]=rpx_fp_mul(tau,q[i]);if(!isfinite(rd[i]))return DOUBLE_OVERFLOW;}return OK;}
 if(phase==2){for(i=first;i<last;i++){v=rpx_fp_mul(-tau,er[i]);for(p=ep[i];p<ep[i+1];p++){j=ei[p];if(j<0||j>=nx)return BAD_ARGUMENT;v=rpx_fp_addmul(v,ev[p],xx[j]);rd[j]=rpx_fp_addmul(rd[j],ev[p],xy[i]);if(!isfinite(rd[j]))return DOUBLE_OVERFLOW;}if(!isfinite(v))return DOUBLE_OVERFLOW;rp[i]=v;}return OK;}
 if(phase==3)return rpx_gtadd(nb,ns,nx,first,last,dims,starts,ip,cp,ids,coef,dual,rd,-1);
 if(phase==4){for(i=first;i<last;i++)if(fabs(rd[i])>m[1])m[1]=fabs(rd[i]);return OK;}
 if(phase==5){for(i=first;i<last;i++){v=i<nm?rp[i]:rg[i-nm];if(fabs(v)>m[0])m[0]=fabs(v);}return OK;}
 return BAD_ARGUMENT;
}
#undef VP

/* Full cone direction loops: normal RHS work, normal output, HSD DD output.
 * Common model arrays are packed once and invalidated on numeric updates. */
API rpx_solver_cones(int32_t op,int32_t nb,int32_t ns,int32_t nx,int32_t first,int32_t last,const int32_t *dims,const int32_t *starts,const int32_t *ip,const int32_t *cp,const int32_t *ids,const double *coef,const uintptr_t *a) {
 int32_t i,k,p,r,j,id,base;int e=0;double v;dd t;
 if(!bounds(nb,first,last)||ns<1||nx<1||!dims||!starts||!ip||!cp||!ids||!coef||!a)return BAD_ARGUMENT;
 for(i=0;i<(op==0?4:(op==1?6:12));i++)if(!a[i])return BAD_ARGUMENT;
 for(k=first;k<last;k++){p=starts[k];if((dims[k]!=1&&dims[k]!=3)||p<0||p>ns-dims[k]||ip[k+1]<ip[k])return BAD_ARGUMENT;
  if(op==0){const double *rc=(const double*)a[0],*rg=(const double*)a[1],*hh=(const double*)a[2];double *work=(double*)a[3];
   for(r=0;r<dims[k];r++){v=rc[p+r];for(j=0;j<dims[k];j++)v=rpx_fp_addmul(v,hh[k+nb*(r+3*j)],rg[p+j]);if(!isfinite(v))return DOUBLE_OVERFLOW;work[p+r]=v;}}
  else if(op==1){const double *rc=(const double*)a[0],*rg=(const double*)a[1],*hh=(const double*)a[2],*dx=(const double*)a[3];double *ds=(double*)a[4],*dz=(double*)a[5];
   for(r=0;r<dims[k];r++){v=0;for(j=ip[k];j<ip[k+1];j++){id=ids[j];if(id<0||id>=nx)return BAD_ARGUMENT;v=rpx_fp_addmul(v,coef[cp[k]+r*(ip[k+1]-ip[k])+j-ip[k]],dx[id]);}ds[p+r]=rpx_fp_add(v,rg[p+r]);}
   for(r=0;r<dims[k];r++){v=-rc[p+r];for(j=0;j<dims[k];j++)v=rpx_fp_mulsub(v,hh[k+nb*(r+3*j)],ds[p+j]);if(!isfinite(v))return DOUBLE_OVERFLOW;dz[p+r]=v;}}
  else if(op==2){const double *offset=(const double*)a[0],*rg=(const double*)a[1],*rc=(const double*)a[2];const dd *dt=(const dd*)a[3],*direction=(const dd*)a[4];
   const int32_t *at=(const int32_t*)a[5];const ddcoef *co=(const ddcoef*)a[6],*hh=(const ddcoef*)a[7];dd *ds=(dd*)a[8],*dz=(dd*)a[9];double *ds0=(double*)a[10],*dz0=(double*)a[11];
   for(r=0;r<dims[k];r++){t=mul(from_double(offset[p+r]),*dt,&e);if(e)return e;ds[p+r]=add(t,from_double(rg[p+r]));base=at[k]+r*(ip[k+1]-ip[k]);
    for(j=ip[k];j<ip[k+1];j++){id=ids[j];if(id<0||id>=nx)return BAD_ARGUMENT;t=bulk_coef_product(co[base+j-ip[k]],direction[id],&e);if(e)return e;ds[p+r]=add(ds[p+r],t);}}
   for(r=0;r<dims[k];r++){dz[p+r]=from_double(-rc[p+r]);for(j=0;j<dims[k];j++){t=bulk_coef_product(hh[k+nb*(r+3*j)],ds[p+j],&e);if(e)return e;dz[p+r]=sub(dz[p+r],t);}ds0[p+r]=value(ds[p+r],&e);if(e)return e;dz0[p+r]=value(dz[p+r],&e);if(e)return e;}}
  else return BAD_ARGUMENT;
 }return OK;
}
