/* Time-budgeted batches. All array arguments are borrowed only during this call.
 * Internal chunks preserve each existing kernel's iteration and store order.
 * SPDX-License-Identifier: LGPL-2.1-or-later */
#define NOMINMAX
#include <windows.h>
#include <cstring>
#include <cstdint>
#include <algorithm>
struct DD{double hi,lo;};
extern "C" int __stdcall rpx_solver_vectors(int32_t,int32_t,int32_t,int32_t,const uintptr_t*,double,double*);
extern "C" int __stdcall rpx_numeric(int32_t,int32_t,int32_t,const int32_t*,const int32_t*,const double*,const int32_t*,const int32_t*,const int32_t*,const int32_t*,const int32_t*,double*,double*,double*,double*,int32_t*);
extern "C" int __stdcall rpx_dd_numeric(int32_t,int32_t,int32_t,const int32_t*,const int32_t*,const DD*,const int32_t*,const int32_t*,const int32_t*,const int32_t*,const int32_t*,DD*,DD*,DD*,double*,int32_t*);
#define API extern "C" __declspec(dllexport) int __stdcall
struct Clock{LARGE_INTEGER frequency,start,now;explicit Clock(int ms){QueryPerformanceFrequency(&frequency);QueryPerformanceCounter(&start);start.QuadPart+=frequency.QuadPart*ms/1000;}bool due(){QueryPerformanceCounter(&now);return now.QuadPart>=start.QuadPart;}};
API rpx_vector_step(int32_t op,int32_t n,int32_t*cursor,const uintptr_t*args,double alpha,double*metrics,int32_t ms){
 if(!cursor||!args||!metrics||n<0||*cursor<0||*cursor>n||ms<0||ms>1000)return 100;Clock clock(ms);
 do{int32_t last=std::min(n,*cursor+512);int code=rpx_solver_vectors(op,n,*cursor,last,args,alpha,metrics);if(code)return code;*cursor=last;
  if((op==16&&metrics[1]==0)||(op==17&&metrics[2]==0)){*cursor=n;return 0;}
  if(ms==0||*cursor==n)return 0;
 }while(!clock.due());return 0;
}
API rpx_factor_step(int32_t n,int32_t*cursor,const uintptr_t*a,int32_t precise,int32_t ms,double*minimum,int32_t*failure){
 if(!cursor||!a||!minimum||!failure||n<1||*cursor<0||*cursor>=n||(precise!=0&&precise!=1)||ms<0||ms>1000)return 100;
 for(int i=0;i<11;i++)if(!a[i])return 100;Clock clock(ms);
 auto ap=reinterpret_cast<const int32_t*>(a[0]),ai=reinterpret_cast<const int32_t*>(a[1]),cp=reinterpret_cast<const int32_t*>(a[3]),ri=reinterpret_cast<const int32_t*>(a[4]),rp=reinterpret_cast<const int32_t*>(a[5]),rn=reinterpret_cast<const int32_t*>(a[6]),wa=reinterpret_cast<const int32_t*>(a[7]);
 do{int32_t last=std::min(n,*cursor+(precise?32:128));int code;
  if(precise)code=rpx_dd_numeric(n,*cursor,last,ap,ai,reinterpret_cast<const DD*>(a[2]),cp,ri,rp,rn,wa,reinterpret_cast<DD*>(a[8]),reinterpret_cast<DD*>(a[9]),reinterpret_cast<DD*>(a[10]),minimum,failure);
  else code=rpx_numeric(n,*cursor,last,ap,ai,reinterpret_cast<const double*>(a[2]),cp,ri,rp,rn,wa,reinterpret_cast<double*>(a[8]),reinterpret_cast<double*>(a[9]),reinterpret_cast<double*>(a[10]),minimum,failure);
  if(code)return code;*cursor=last;if(ms==0||*cursor==n||(precise&&*cursor%256==0))return 0;
 }while(!clock.due());return 0;
}
API rpx_exact_spans(int32_t count,const uintptr_t*a,const int32_t*lengths,int32_t*equal){
 if(count<1||count>64||!a||!lengths||!equal)return 330;*equal=0;
 for(int i=0;i<count;i++){if(lengths[i]<0||!a[2*i]||!a[2*i+1])return 330;if(std::memcmp(reinterpret_cast<const void*>(a[2*i]),reinterpret_cast<const void*>(a[2*i+1]),size_t(lengths[i])*8)!=0)return 0;}
 *equal=1;return 0;
}
