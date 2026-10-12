/* DLL5: exact-order DD KKT assembly, using the same checked DD primitives.
 * No pointers survive a call; phase/cursor are caller-owned. */
API rpx_hdd_assemble(int32_t n,int32_t nx,int32_t nb,int32_t nz,int32_t phase,int32_t first,int32_t last,const uintptr_t*a){
 const double*base;dd*out;const int32_t*dims,*ip,*sp,*slots,*at,*diag,*perm,*ap,*ai;const ddcoef*coef0,*hh,*scale;
 int32_t i,j,k,r,q,p,col,index,cnt,slot;int e=0;dd work[3],coefficient,t,total;double wh[3],wl[3];
 if(n<1||nx<1||nx>n||nb<1||nz<1||!a)return BAD_ARGUMENT;for(i=0;i<14;i++)if(!a[i])return BAD_ARGUMENT;
 base=(const double*)a[0];out=(dd*)a[1];dims=(const int32_t*)a[2];ip=(const int32_t*)a[3];sp=(const int32_t*)a[4];slots=(const int32_t*)a[5];at=(const int32_t*)a[6];coef0=(const ddcoef*)a[7];hh=(const ddcoef*)a[8];diag=(const int32_t*)a[9];perm=(const int32_t*)a[10];scale=(const ddcoef*)a[11];ap=(const int32_t*)a[12];ai=(const int32_t*)a[13];
 if(phase==0){if(!bounds(nz,first,last))return BAD_ARGUMENT;for(i=first;i<last;i++)out[i]=from_double(base[i]);return 0;}
 if(phase==1){if(!bounds(nb,first,last))return BAD_ARGUMENT;for(i=first;i<last;i++){
  cnt=ip[i+1]-ip[i];slot=sp[i];if(cnt<0||(dims[i]!=1&&dims[i]!=3)||sp[i+1]-slot!=cnt*(cnt+1)/2)return BAD_ARGUMENT;
  for(j=0;j<cnt;j++){
   for(r=0;r<dims[i];r++){work[r]=from_double(0);for(q=0;q<dims[i];q++){index=at[i]+q*cnt+j;coefficient=from_double(coef0[index].hi);t=mul_known((dd){hh[i+nb*(r+3*q)].hi,hh[i+nb*(r+3*q)].lo},coefficient,coef0[index].sh,coef0[index].sl,&e);if(e)return e;work[r]=add(work[r],t);}split(work[r].hi,&wh[r],&wl[r]);}
   for(k=0;k<=j;k++){total=from_double(0);for(r=0;r<dims[i];r++){index=at[i]+r*cnt+k;t=mul_known((dd){coef0[index].hi,coef0[index].lo},work[r],wh[r],wl[r],&e);if(e)return e;total=add(total,t);}p=slots[slot++];if(p<0||p>=nz)return BAD_ARGUMENT;out[p]=add(out[p],total);}
  }
 }return 0;}
 if(phase==2){if(!bounds(n,first,last))return BAD_ARGUMENT;for(i=first;i<last;i++){p=diag[i];if(p<0||p>=nz||perm[i]<0||perm[i]>=n)return BAD_ARGUMENT;out[p]=perm[i]<nx?add(out[p],from_double(1e-10)):from_double(-1e-10);}return 0;}
 if(phase==3){if(!bounds(n,first,last)||ap[0]!=0||ap[n]!=nz)return BAD_ARGUMENT;for(col=first;col<last;col++){if(ap[col]<0||ap[col+1]<ap[col]||ap[col+1]>nz)return BAD_ARGUMENT;for(p=ap[col];p<ap[col+1];p++){i=ai[p];if(i<0||i>col)return BAD_ARGUMENT;out[p]=bulk_coef_product(scale[col],out[p],&e);if(e)return e;out[p]=bulk_coef_product(scale[i],out[p],&e);if(e)return e;}}return 0;}
 return BAD_ARGUMENT;
}
API rpx_hdd_step(int32_t n,int32_t nx,int32_t nb,int32_t nz,int32_t phase,int32_t*cursor,const uintptr_t*a,int32_t ms){
 int32_t total,last;int code;LARGE_INTEGER clock,freq;LONGLONG deadline;
 if(!cursor||phase<0||phase>3||ms<0||ms>1000)return BAD_ARGUMENT;total=phase==0?nz:phase==1?nb:n;
 if(*cursor<0||*cursor>total)return BAD_ARGUMENT;QueryPerformanceFrequency(&freq);QueryPerformanceCounter(&clock);deadline=clock.QuadPart+freq.QuadPart*ms/1000;
 do{last=*cursor+128;if(last>total)last=total;code=rpx_hdd_assemble(n,nx,nb,nz,phase,*cursor,last,a);if(code)return code;*cursor=last;if(*cursor==total||ms==0||(phase==1&&*cursor%256==0))break;QueryPerformanceCounter(&clock);}while(clock.QuadPart<deadline);return 0;
}
API rpx_krylov_givens(double a,double b,double magnitude,double*out){
 double x,y;int e=rpx_bulk_ready();if(e)return e;if(!out||magnitude==0)return BAD_ARGUMENT;
 x=nt_power(rpx_fp_div(a,magnitude),2);y=nt_power(rpx_fp_div(b,magnitude),2);*out=rpx_fp_mul(magnitude,rpx_fp_sqrt(rpx_fp_add(x,y)));return isfinite(*out)?0:DOUBLE_OVERFLOW;
}
