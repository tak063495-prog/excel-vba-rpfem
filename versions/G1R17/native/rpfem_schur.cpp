/* DLL5 owned Schur layout, local factors and complete checked solve.
 * Order and stopping rules match RPX_Schur; no factor/iterate VBA pointer retained.
 * SPDX-License-Identifier: LGPL-2.1-or-later */
#define NOMINMAX
#include <windows.h>
#include <cstdint>
#include <vector>
#include <unordered_map>
#include <memory>
#include <algorithm>
#include <cmath>
#include <cstring>
#include <new>
extern "C" double rpx_fp_add(double,double),rpx_fp_sub(double,double),rpx_fp_div(double,double),rpx_fp_mulsub(double,double,double),rpx_fp_addmul(double,double,double),rpx_fp_sqrt(double);
extern "C" int __stdcall rpx_backsolve(int32_t,int32_t,int32_t,int32_t,const int32_t*,const int32_t*,const double*,const double*,const double*,double*);
extern "C" int __stdcall rpx_solver_vectors(int32_t,int32_t,int32_t,int32_t,const uintptr_t*,double,double*);
#define API extern "C" __declspec(dllexport) int __stdcall
namespace {
struct Failure{int code;};
double clock(){LARGE_INTEGER q,f;QueryPerformanceCounter(&q);QueryPerformanceFrequency(&f);return double(q.QuadPart)/double(f.QuadPart);}
bool finite(double x){return std::isfinite(x)&&x>=-1e290&&x<=1e290;}
double stored(double x){if(!std::isfinite(x))throw Failure{206};return x;}
struct Context{
 int nf,nr,nb,fullnz,nz,phase=0,cursor=0,solvephase=0,solvecursor=0,checks=0,corrections=0,solves=0;
 uint64_t pattern,epoch=0,serial=0;bool factoring=false,factorready=false,solving=false,check=false,reuse=false,structure=false;int rejection=0;
 std::vector<int32_t> pos,hp,bp,rows,ptr,sp,slots,inv,full,direct,ap,ai,cp,ri;
 std::vector<double> av,chol,bv,u,fullav,lv,dv,rhs,workrhs,answer,out,res,reducedRhs,reduced,z;
 double errors[4]={},times[8]={},error=0;uint64_t local=0;
 Context(int f,int r,int b,int fn,int n,uint64_t pat,const uintptr_t*a,int rowsn,int slotsn):nf(f),nr(r),nb(b),fullnz(fn),nz(n),pattern(pat),av(n),chol(size_t(b)*9),bv(size_t(rowsn)*3),u(size_t(rowsn)*3),rhs(f),workrhs(f),answer(f),out(f),res(f),reducedRhs(r),reduced(r),z(size_t(b)*3){
  auto copy=[&](std::vector<int32_t>&v,int index,size_t count){auto p=reinterpret_cast<const int32_t*>(a[index]);v.assign(p,p+count);};
  copy(pos,0,size_t(b)*3);copy(hp,1,size_t(b)*9);copy(ptr,2,size_t(b)+1);copy(rows,3,rowsn);copy(bp,4,size_t(rowsn)*3);copy(sp,5,size_t(b)+1);copy(slots,6,slotsn);copy(inv,7,r);copy(full,8,r);copy(direct,9,fn);
  if(ptr[0]!=0||ptr[b]!=rowsn||sp[0]!=0||sp[b]!=slotsn)throw Failure{330};
  std::vector<int>seen(r,0);
  for(int i=0;i<r;i++){if(inv[i]<0||inv[i]>=r||seen[inv[i]]++||full[i]<0||full[i]>=f)throw Failure{330};}
  for(int i:pos)if(i<0||i>=f)throw Failure{330};for(int i:hp)if(i<-1||i>=fn)throw Failure{330};for(int i:bp)if(i<-1||i>=fn)throw Failure{330};
  for(int i:rows)if(i<0||i>=r)throw Failure{330};for(int i:slots)if(i<0||i>=n)throw Failure{330};for(int i:direct)if(i<-1||i>=n)throw Failure{330};
  for(int i=0;i<b;i++){int64_t width=int64_t(ptr[i+1])-ptr[i];if(width<0||sp[i+1]<sp[i]||int64_t(sp[i+1])-sp[i]!=width*width)throw Failure{330};}
 }
 void local_solve(int b,double*x){
  size_t at=size_t(b)*9;for(int i=0;i<3;i++){for(int j=0;j<i;j++)x[i]=stored(rpx_fp_mulsub(x[i],chol[at+i+3*j],x[j]));x[i]=stored(rpx_fp_div(x[i],chol[at+i+3*i]));}
  for(int i=2;i>=0;i--){for(int j=i+1;j<3;j++)x[i]=stored(rpx_fp_mulsub(x[i],chol[at+j+3*i],x[j]));x[i]=stored(rpx_fp_div(x[i],chol[at+i+3*i]));}
 }
 size_t bytes()const{return sizeof(*this)+4*(pos.capacity()+hp.capacity()+bp.capacity()+rows.capacity()+ptr.capacity()+sp.capacity()+slots.capacity()+inv.capacity()+full.capacity()+direct.capacity()+ap.capacity()+ai.capacity()+cp.capacity()+ri.capacity())+8*(av.capacity()+chol.capacity()+bv.capacity()+u.capacity()+fullav.capacity()+lv.capacity()+dv.capacity()+rhs.capacity()+workrhs.capacity()+answer.capacity()+out.capacity()+res.capacity()+reducedRhs.capacity()+reduced.capacity()+z.capacity());}
};
std::unordered_map<uintptr_t,std::unique_ptr<Context>> contexts;uintptr_t nextid=0;
Context&get(uintptr_t h,uint64_t pattern){auto i=contexts.find(h);if(i==contexts.end()||i->second->pattern!=pattern)throw Failure{318};return *i->second;}
template<class F>int checked(F f){try{return f();}catch(Failure e){return e.code;}catch(const std::bad_alloc&){return 207;}catch(...){return 208;}}
}
API rpx_schur_create(int32_t nf,int32_t nr,int32_t nb,int32_t fullnz,int32_t nz,int32_t rows,int32_t slots,uint64_t pattern,const uintptr_t*a,uintptr_t*h){
 if(!h||!a||nf<1||nr<1||nr>nf||nb<1||fullnz<1||nz<1||rows<1||slots<1||pattern==0)return 330;*h=0;for(int i=0;i<10;i++)if(!a[i])return 330;
 return checked([&](){auto p=std::make_unique<Context>(nf,nr,nb,fullnz,nz,pattern,a,rows,slots);if(nextid==UINTPTR_MAX)throw Failure{318};uintptr_t id=++nextid;contexts.emplace(id,std::move(p));*h=id;return 0;});
}
API rpx_schur_destroy(uintptr_t h){return checked([&](){auto i=contexts.find(h);if(i==contexts.end())throw Failure{318};contexts.erase(i);return 0;});}
API rpx_schur_bytes(uintptr_t h,uint64_t pattern,uint64_t*out){if(!out)return 330;return checked([&](){*out=get(h,pattern).bytes();return 0;});}
API rpx_schur_factor_begin(uintptr_t h,uint64_t pattern,uint64_t epoch){return checked([&](){auto&p=get(h,pattern);if(!epoch||p.factoring||p.solving)throw Failure{318};p.epoch=epoch;p.factoring=true;p.factorready=false;p.phase=0;p.cursor=0;p.rejection=0;p.local=0;std::fill(p.av.begin(),p.av.end(),0.0);std::fill(p.times,p.times+8,0.0);return 0;});}
API rpx_schur_factor_step(uintptr_t h,uint64_t pattern,uint64_t epoch,const double*matrix,int32_t ms,int32_t*done,int32_t*reason){
 if(!matrix||!done||!reason||ms<0||ms>1000)return 330;*done=0;*reason=0;
 return checked([&](){auto&p=get(h,pattern);if(!p.factoring||p.epoch!=epoch)throw Failure{318};double deadline=clock()+double(ms)/1000;
  do{double start=clock();int total=p.nb;
   if(p.phase==0){total=p.fullnz;int last=std::min(total,p.cursor+512);for(int i=p.cursor;i<last;i++)if(!finite(matrix[i])){p.rejection=1;break;}p.cursor=last;}
   else if(p.phase==1){total=p.fullnz;int last=std::min(total,p.cursor+512);for(int i=p.cursor;i<last;i++)if(p.direct[i]>=0){int at=p.direct[i];p.av[at]=stored(rpx_fp_add(p.av[at],matrix[i]));}p.cursor=last;}
   else if(p.phase==2){int b=p.cursor;size_t at=size_t(b)*9;double x[3];
    for(int i=0;i<3&&!p.rejection;i++)for(int j=0;j<=i;j++){int kpos=p.hp[at+i+3*j];double v=kpos>=0?matrix[kpos]:0.0;for(int k=0;k<j;k++)v=stored(rpx_fp_mulsub(v,p.chol[at+i+3*k],p.chol[at+j+3*k]));
     if(i==j){if(!finite(v)||v<=1e-20){p.rejection=2;break;}p.chol[at+i+3*j]=rpx_fp_sqrt(v);}else p.chol[at+i+3*j]=stored(rpx_fp_div(v,p.chol[at+j+3*j]));}
    if(!p.rejection){p.local+=uint64_t(p.ptr[b+1]-p.ptr[b]);for(int r=p.ptr[b];r<p.ptr[b+1]&&!p.rejection;r++){for(int k=0;k<3;k++){int index=p.bp[size_t(r)*3+k];x[k]=index>=0?matrix[index]:0.0;p.bv[size_t(r)*3+k]=x[k];}p.local_solve(b,x);for(int k=0;k<3;k++){if(!finite(x[k])){p.rejection=3;break;}p.u[size_t(r)*3+k]=x[k];}}}p.cursor++;
   }else if(p.phase==3){int b=p.cursor,width=p.ptr[b+1]-p.ptr[b],begin=p.ptr[b];for(int i=0;i<width;i++)for(int j=0;j<=i;j++){double v=0;for(int k=0;k<3;k++)v=stored(rpx_fp_addmul(v,p.bv[size_t(begin+i)*3+k],p.u[size_t(begin+j)*3+k]));int index=p.slots[p.sp[b]+i+width*j];p.av[index]=stored(rpx_fp_sub(p.av[index],v));}p.cursor++;}
   else if(p.phase==4){total=p.nz;int last=std::min(total,p.cursor+512);for(int i=p.cursor;i<last;i++)if(!finite(p.av[i])){p.rejection=4;break;}p.cursor=last;}
   else throw Failure{318};
   p.times[p.phase==3?2:1]+=clock()-start;
   if(p.rejection){p.factoring=false;*done=1;*reason=p.rejection;return 0;}
   if(p.cursor==total){p.cursor=0;if(++p.phase==5){p.factoring=false;*done=1;return 0;}}
   if(ms==0)return 0;
  }while(clock()<deadline);return 0;
 });
}
API rpx_schur_factor_result(uintptr_t h,uint64_t pattern,uint64_t epoch,double*av,double*metrics){if(!av||!metrics)return 330;return checked([&](){auto&p=get(h,pattern);if(p.epoch!=epoch||p.factoring||p.phase!=5||p.rejection)throw Failure{318};std::copy(p.av.begin(),p.av.end(),av);metrics[0]=double(p.local);metrics[1]=p.times[1];metrics[2]=p.times[2];return 0;});}
API rpx_schur_factor_commit(uintptr_t h,uint64_t pattern,uint64_t epoch,const uintptr_t*a,int32_t factor_nnz){
 if(!a||factor_nnz<0)return 330;for(int i=0;i<7;i++)if(!a[i])return 330;
 return checked([&](){auto&p=get(h,pattern);if(p.epoch!=epoch||p.factoring||p.phase!=5||p.rejection||p.solving)throw Failure{318};
  auto copy=[&](std::vector<int32_t>&v,int index,size_t count){auto q=reinterpret_cast<const int32_t*>(a[index]);if(p.structure){if(v.size()!=count||!std::equal(v.begin(),v.end(),q))throw Failure{318};}else v.assign(q,q+count);};
  copy(p.ap,0,size_t(p.nf)+1);copy(p.ai,1,p.fullnz);copy(p.cp,3,size_t(p.nr)+1);copy(p.ri,4,factor_nnz);
  if(p.ap[0]!=0||p.ap[p.nf]!=p.fullnz||p.cp[0]!=0||p.cp[p.nr]!=factor_nnz)throw Failure{330};
  auto av=reinterpret_cast<const double*>(a[2]),lv=reinterpret_cast<const double*>(a[5]),dv=reinterpret_cast<const double*>(a[6]);
  p.fullav.assign(av,av+p.fullnz);p.lv.assign(lv,lv+factor_nnz);p.dv.assign(dv,dv+p.nr);p.structure=true;p.factorready=true;return 0;
 });
}
API rpx_schur_solve_begin(uintptr_t h,uint64_t pattern,uint64_t epoch,const double*rhs,int32_t check,int32_t reuse,uint64_t*serial){
 if(!rhs||!serial||(check!=0&&check!=1)||(reuse!=0&&reuse!=1))return 330;
 return checked([&](){auto&p=get(h,pattern);if(!p.factorready||p.epoch!=epoch||p.factoring||p.solving||p.serial==UINT64_MAX)throw Failure{318};p.serial++;*serial=p.serial;
  std::copy(rhs,rhs+p.nf,p.rhs.begin());p.workrhs=p.rhs;p.check=check!=0;p.reuse=reuse!=0;p.solving=true;p.solvephase=0;p.solvecursor=0;p.checks=0;p.corrections=0;p.solves=0;p.error=0;std::fill(p.errors,p.errors+4,0.0);std::fill(p.times,p.times+8,0.0);return 0;
 });
}
API rpx_schur_solve_step(uintptr_t h,uint64_t pattern,uint64_t epoch,uint64_t serial,int32_t ms,int32_t*done){
 if(!done||ms<0||ms>1000)return 330;*done=0;
 return checked([&](){auto&p=get(h,pattern);if(!p.solving||!p.factorready||p.epoch!=epoch||p.serial!=serial)throw Failure{318};double deadline=clock()+double(ms)/1000;
  do{double start=clock();int total=p.nb,bucket=4;int i=p.solvecursor,last;double x[3];
   if(p.solvephase==0){std::fill(p.out.begin(),p.out.end(),0.0);for(int r=0;r<p.nr;r++)p.reducedRhs[p.inv[r]]=p.workrhs[p.full[r]];p.solvecursor=total;}
   else if(p.solvephase==1){int b=i;for(int k=0;k<3;k++)x[k]=p.workrhs[p.pos[size_t(b)*3+k]];p.local_solve(b,x);if(p.reuse)for(int k=0;k<3;k++)p.z[size_t(b)*3+k]=x[k];
    for(int r=p.ptr[b];r<p.ptr[b+1];r++){double v=0;for(int k=0;k<3;k++)v=stored(rpx_fp_addmul(v,p.bv[size_t(r)*3+k],x[k]));int index=p.inv[p.rows[r]];p.reducedRhs[index]=stored(rpx_fp_sub(p.reducedRhs[index],v));}p.solvecursor++;
   }else if(p.solvephase>=2&&p.solvephase<=4){bucket=5;total=p.nr;last=std::min(total,i+32);int first=i,phase=p.solvephase-1;if(phase==3){last=total-i;first=std::max(0,last-32);}
    int32_t emptyrow=0;double emptyvalue=0;int code=rpx_backsolve(p.nr,phase,first,last,p.cp.data(),p.ri.empty()?&emptyrow:p.ri.data(),p.lv.empty()?&emptyvalue:p.lv.data(),p.dv.data(),p.reducedRhs.data(),p.reduced.data());if(code)throw Failure{code};p.solvecursor=std::min(total,i+32);
   }else if(p.solvephase==5){bucket=6;total=p.nr;last=std::min(total,i+512);for(int r=i;r<last;r++)p.out[p.full[r]]=p.reduced[p.inv[r]];p.solvecursor=last;
   }else if(p.solvephase==6){bucket=6;int b=i;
    for(int k=0;k<3;k++){x[k]=p.reuse?p.z[size_t(b)*3+k]:p.workrhs[p.pos[size_t(b)*3+k]];for(int r=p.ptr[b];r<p.ptr[b+1];r++)x[k]=stored(rpx_fp_mulsub(x[k],p.reuse?p.u[size_t(r)*3+k]:p.bv[size_t(r)*3+k],p.out[p.full[p.rows[r]]]));}
    if(!p.reuse)p.local_solve(b,x);for(int k=0;k<3;k++)p.out[p.pos[size_t(b)*3+k]]=x[k];p.solvecursor++;
   }else if(p.solvephase==7){bucket=7;p.solves++;if(p.solves==1)p.answer=p.out;else{for(int r=0;r<p.nf;r++)p.answer[r]=stored(rpx_fp_add(p.answer[r],p.out[r]));p.corrections++;}
    if(!p.check){p.solving=false;*done=1;return 0;}for(double v:p.answer)if(!finite(v)){p.error=1e300;p.solving=false;*done=1;return 0;}
    std::fill(p.res.begin(),p.res.end(),0.0);p.error=0;p.solvecursor=total;
   }else if(p.solvephase==8){bucket=7;total=p.nf;last=std::min(total,i+64);uintptr_t args[5]={reinterpret_cast<uintptr_t>(p.ap.data()),reinterpret_cast<uintptr_t>(p.ai.data()),reinterpret_cast<uintptr_t>(p.fullav.data()),reinterpret_cast<uintptr_t>(p.answer.data()),reinterpret_cast<uintptr_t>(p.res.data())};double m[1]={};int code=rpx_solver_vectors(11,p.nf,i,last,args,0,m);if(code)throw Failure{code};p.solvecursor=last;
   }else if(p.solvephase==9){bucket=7;total=p.nf;last=std::min(total,i+512);for(int r=i;r<last;r++){if(p.checks==3&&(!finite(p.answer[r])||!finite(p.res[r]))){p.error=1e300;p.solving=false;*done=1;return 0;}double v=stored(rpx_fp_sub(p.rhs[r],p.res[r]));if(std::fabs(v)>p.error)p.error=std::fabs(v);p.res[r]=v;}p.solvecursor=last;
   }else throw Failure{318};
   p.times[bucket]+=clock()-start;
   if(p.solvecursor==total){p.solvecursor=0;if(p.solvephase==9){p.errors[p.checks++]=p.error;if(p.checks==4){p.solving=false;*done=1;return 0;}p.workrhs=p.res;p.solvephase=0;}else{if(p.solvephase==1)p.reduced=p.reducedRhs;p.solvephase++;}}
   if(ms==0)return 0;
  }while(clock()<deadline);return 0;
 });
}
API rpx_schur_solve_result(uintptr_t h,uint64_t pattern,uint64_t epoch,uint64_t serial,double*out,double*metrics){
 if(!out||!metrics)return 330;return checked([&](){auto&p=get(h,pattern);if(p.solving||!p.factorready||p.epoch!=epoch||p.serial!=serial||p.solves<1)throw Failure{318};
  std::copy(p.answer.begin(),p.answer.end(),out);metrics[0]=p.solves;metrics[1]=p.checks;metrics[2]=p.corrections;metrics[3]=p.error;for(int i=0;i<4;i++)metrics[4+i]=p.errors[i];for(int i=4;i<8;i++)metrics[8+i-4]=p.times[i];return 0;});
}
API rpx_schur_live_handles(void){return static_cast<int>(contexts.size());}
API rpx_schur_finite(int32_t n,const double*matrix,int32_t*invalid){if(n<1||!matrix||!invalid)return 330;*invalid=0;for(int i=0;i<n;i++)if(!finite(matrix[i])){*invalid=1;break;}return 0;}
