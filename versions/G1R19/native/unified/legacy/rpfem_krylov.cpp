/* DLL5: resumable right-preconditioned GMRES, exact VBA HKRepair order.
 * Width 24, three restarts, two compensated MGS passes; checked DD operator.
 * Work vectors are owned. Model/factor arguments are borrowed per Step only.
 * SPDX-License-Identifier: LGPL-2.1-or-later */
#define NOMINMAX
#include <windows.h>
#include <cstdint>
#include <vector>
#include <unordered_map>
#include <memory>
#include <algorithm>
#include <cmath>
#include <new>
struct DD{double hi,lo;};
extern "C" int __stdcall rpx_dd_test(int32_t,const DD*,const DD*,DD*);
extern "C" int __stdcall rpx_hoperator(int32_t,int32_t,int32_t,int32_t,int32_t,int32_t,int32_t,int32_t,const uintptr_t*);
extern "C" int __stdcall rpx_backsolve(int32_t,int32_t,int32_t,int32_t,const int32_t*,const int32_t*,const double*,const double*,const double*,double*);
extern "C" int __stdcall rpx_dd_backsolve(int32_t,int32_t,int32_t,int32_t,const int32_t*,const int32_t*,const DD*,const DD*,const double*,DD*);
extern "C" int __stdcall rpx_lu_retained_solve(uintptr_t,int32_t,int32_t,int32_t,int32_t,int32_t,const double*,double*,int32_t*);
extern "C" int __stdcall rpx_krylov_givens(double,double,double,double*);
extern "C" double rpx_fp_add(double,double),rpx_fp_sub(double,double),rpx_fp_mul(double,double),rpx_fp_div(double,double),rpx_fp_addmul(double,double,double),rpx_fp_mulsub(double,double,double),rpx_fp_dot2(double,double,double,double),rpx_fp_sqrt(double),rpx_fp_mulminus(double,double,double),rpx_fp_subsub(double,double,double);
extern "C" int32_t rpx_fp_greater_product(double,double,double);
#define API extern "C" __declspec(dllexport) int __stdcall
namespace{
struct Failure{int code;};
double clock(){LARGE_INTEGER q,f;QueryPerformanceCounter(&q);QueryPerformanceFrequency(&f);return double(q.QuadPart)/double(f.QuadPart);}
bool finite(double x){return std::isfinite(x)&&std::fabs(x)<=1e290;}
double stored(double x){if(!std::isfinite(x))throw Failure{206};return x;}
DD operation(int op,DD a,DD b){DD out;int c=rpx_dd_test(op,&a,&b,&out);if(c)throw Failure{c};return out;}
double value(DD a){double x=stored(rpx_fp_add(a.hi,a.lo));if(!finite(a.hi)||!finite(a.lo)||!finite(x))throw Failure{204};return x;}
uintptr_t ptr(const void*p){return reinterpret_cast<uintptr_t>(p);}
struct Job{
 int n,nx,nb,ns,kind=0,factor=0,unshifted=0,phase=0,subphase=0,cursor=0,cycle=0,j=0,k=0,pass=0,used=0,steps=0,finitechecks=0;
 uint64_t serial=0;bool active=false,complete=false,final=false;
 double best=0,previous=0,error=0,largest=0,total=0,beta=0,nextnorm=0,dot=0,compensation=0;
 std::vector<DD> bdd,answer,saved,vectorDD,product,ddsolution,opx,opv,opg,oph;
 std::vector<double> basis,preconditioned,residual,vector,solved;
 double hess[25*24]={},cosine[24]={},sine[24]={},rhs[25]={},coefficients[24]={};
 Job(int a,int x,int b,int s):n(a),nx(x),nb(b),ns(s),bdd(a),answer(a),saved(a),vectorDD(a),product(a),ddsolution(a),opx(a),opv(a),opg(s),oph(s),basis(size_t(a)*25),preconditioned(size_t(a)*24),residual(a),vector(a),solved(a){}
 double&H(int r,int col){return hess[r+25*col];}
 size_t bytes()const{return sizeof(*this)+16*(bdd.capacity()+answer.capacity()+saved.capacity()+vectorDD.capacity()+product.capacity()+ddsolution.capacity()+opx.capacity()+opv.capacity()+opg.capacity()+oph.capacity())+8*(basis.capacity()+preconditioned.capacity()+residual.capacity()+vector.capacity()+solved.capacity());}
 void go(int p){phase=p;cursor=0;subphase=0;}
};
std::unordered_map<uintptr_t,std::unique_ptr<Job>> jobs;uintptr_t nextid=0;
Job&get(uintptr_t h){auto i=jobs.find(h);if(i==jobs.end())throw Failure{318};return *i->second;}
template<class F>int checked(F f){try{return f();}catch(Failure e){return e.code;}catch(const std::bad_alloc&){return 207;}catch(...){return 208;}}
}
API rpx_krylov_create(int32_t n,int32_t nx,int32_t nb,int32_t ns,uintptr_t*h){if(!h||n<1||n>1048576||nx<1||nx>n||nb<1||ns<1)return 330;*h=0;return checked([&](){auto p=std::make_unique<Job>(n,nx,nb,ns);if(nextid==UINTPTR_MAX)throw Failure{318};uintptr_t id=++nextid;jobs.emplace(id,std::move(p));*h=id;return 0;});}
API rpx_krylov_destroy(uintptr_t h){return checked([&](){auto i=jobs.find(h);if(i==jobs.end())throw Failure{318};jobs.erase(i);return 0;});}
API rpx_krylov_begin(uintptr_t h,uint64_t serial,int32_t factor,int32_t kind,int32_t unshifted,const DD*bdd,const DD*answer,double best){
 if(!serial||!bdd||!answer||kind<0||kind>2||(unshifted!=0&&unshifted!=1)||!std::isfinite(best))return 330;
 return checked([&](){auto&p=get(h);if(p.active)throw Failure{318};p.serial=serial;p.factor=factor;p.kind=kind;p.unshifted=unshifted;p.best=best;p.previous=best;p.active=true;p.complete=false;p.final=false;p.phase=0;p.cursor=0;p.subphase=0;p.cycle=0;p.steps=0;p.finitechecks=0;
  std::copy(bdd,bdd+p.n,p.bdd.begin());std::copy(answer,answer+p.n,p.answer.begin());p.saved=p.answer;std::fill(p.coefficients,p.coefficients+24,0.0);return 0;});
}
API rpx_krylov_step(uintptr_t h,uint64_t serial,int32_t factor,const uintptr_t*op,const uintptr_t*fa,uintptr_t lu,int32_t ms,int32_t*done){
 if(!op||!fa||!done||ms<0||ms>1000)return 330;*done=0;
 return checked([&](){auto&p=get(h);if(!p.active||p.serial!=serial||p.factor!=factor)throw Failure{318};double deadline=clock()+double(ms)/1000;
  do{int first=p.cursor,last=std::min(p.n,first+512);int c=0;bool advance=false;
   if(p.phase==0||p.phase==9){ // exact original DD operator
    uintptr_t a[20];std::copy(op,op+14,a);a[14]=ptr(p.phase==0?p.answer.data():p.vectorDD.data());a[15]=ptr(p.product.data());a[16]=ptr(p.opx.data());a[17]=ptr(p.opv.data());a[18]=ptr(p.opg.data());a[19]=ptr(p.oph.data());
    if(p.subphase==0){c=rpx_hoperator(p.n,p.nx,p.nb,p.ns,-1,0,0,p.unshifted,a);p.subphase++;p.cursor=0;}
    else{int phase=p.subphase-1,total=p.n;if(phase==1)total=p.nb;if(phase==2)total=p.n-p.nx;last=std::min(total,first+32);if(total)c=rpx_hoperator(p.n,p.nx,p.nb,p.ns,phase,first,last,p.unshifted,a);p.cursor=last;if(last==total){p.cursor=0;if(++p.subphase==5){p.error=0;p.go(p.phase==0?1:10);}}}
   }else if(p.phase==1){
    auto scale=reinterpret_cast<const double*>(fa[4]);for(int i=first;i<last;i++){p.residual[i]=value(operation(1,p.bdd[i],p.product[i]));double v=std::fabs(p.residual[i]);if(v>p.error)p.error=v;if(scale[i]==0)throw Failure{205};v=p.final?stored(rpx_fp_div(v,scale[i])):std::fabs(stored(rpx_fp_div(p.residual[i],scale[i])));if(v>p.error)p.error=v;}
    p.cursor=last;if(last==p.n){if(p.error<p.best){p.best=p.error;p.saved=p.answer;}if(p.final){p.answer=p.saved;p.active=false;p.complete=true;*done=1;return 0;}if(p.best<=1e-11){p.final=true;p.go(0);}else{p.largest=0;p.total=0;p.go(3);}}
   }else if(p.phase==3||p.phase==13){
    auto&v=p.phase==3?p.residual:p.vector;for(int i=first;i<last;i++){if(!finite(v[i]))throw Failure{339};if(std::fabs(v[i])>p.largest)p.largest=std::fabs(v[i]);}p.cursor=last;if(last==p.n){if(p.largest==0){if(p.phase==3){p.final=true;p.go(0);}else{p.nextnorm=0;p.H(p.j+1,p.j)=0;p.go(16);}}else{p.total=0;p.go(p.phase==3?4:14);}}
   }else if(p.phase==4||p.phase==14){
    auto&v=p.phase==4?p.residual:p.vector;for(int i=first;i<last;i++){double x=stored(rpx_fp_div(v[i],p.largest));p.total=stored(rpx_fp_addmul(p.total,x,x));}p.cursor=last;if(last==p.n){double norm=stored(rpx_fp_mul(p.largest,rpx_fp_sqrt(p.total)));if(p.phase==4){p.beta=norm;if(norm==0){p.final=true;p.go(0);}else{std::fill(p.rhs,p.rhs+25,0.0);p.rhs[0]=norm;p.used=0;p.j=0;p.go(5);}}else{p.nextnorm=norm;p.H(p.j+1,p.j)=norm;p.go(norm>0?15:16);}}
   }else if(p.phase==5){for(int i=first;i<last;i++)p.basis[i]=stored(rpx_fp_div(p.residual[i],p.beta));advance=true;if(last==p.n)p.go(6);
   }else if(p.phase==6){for(int i=first;i<last;i++)p.vector[i]=p.basis[size_t(p.j)*p.n+i];advance=true;if(last==p.n)p.go(7);
   }else if(p.phase==7){
    int end=std::min(p.n,first+32),begin=first;if((p.kind==2&&p.subphase==1)||(p.kind!=2&&p.subphase==3)){end=p.n-first;begin=std::max(0,end-32);}
    if(p.kind==2)c=rpx_lu_retained_solve(lu,p.factor,p.n,p.subphase,begin,end,p.vector.data(),p.solved.data(),&p.finitechecks);
    else if(p.kind==1)c=rpx_dd_backsolve(p.n,p.subphase,begin,end,reinterpret_cast<const int32_t*>(fa[0]),reinterpret_cast<const int32_t*>(fa[1]),reinterpret_cast<const DD*>(fa[2]),reinterpret_cast<const DD*>(fa[3]),p.vector.data(),p.ddsolution.data());
    else c=rpx_backsolve(p.n,p.subphase,begin,end,reinterpret_cast<const int32_t*>(fa[0]),reinterpret_cast<const int32_t*>(fa[1]),reinterpret_cast<const double*>(fa[2]),reinterpret_cast<const double*>(fa[3]),p.vector.data(),p.solved.data());
    p.cursor=std::min(p.n,first+32);if(p.cursor==p.n){p.cursor=0;if(++p.subphase==(p.kind==2?3:4))p.go(8);}
   }else if(p.phase==8){for(int i=first;i<last;i++){if(p.kind==1)p.solved[i]=value(p.ddsolution[i]);if(!finite(p.solved[i]))throw Failure{339};p.preconditioned[size_t(p.j)*p.n+i]=p.solved[i];p.vectorDD[i]={p.solved[i],0.0};}advance=true;if(last==p.n)p.go(9);
   }else if(p.phase==10){for(int i=first;i<last;i++)p.vector[i]=value(p.product[i]);advance=true;if(last==p.n){for(int k=0;k<=p.j+1;k++)p.H(k,p.j)=0;p.pass=0;p.k=0;p.dot=0;p.compensation=0;p.go(11);}
   }else if(p.phase==11){for(int i=first;i<last;i++){double adjusted=stored(rpx_fp_mulminus(p.vector[i],p.basis[size_t(p.k)*p.n+i],p.compensation)),next=stored(rpx_fp_add(p.dot,adjusted));p.compensation=stored(rpx_fp_subsub(next,p.dot,adjusted));p.dot=next;}advance=true;if(last==p.n){p.H(p.k,p.j)=stored(rpx_fp_add(p.H(p.k,p.j),p.dot));p.go(12);}
   }else if(p.phase==12){for(int i=first;i<last;i++)p.vector[i]=stored(rpx_fp_mulsub(p.vector[i],p.dot,p.basis[size_t(p.k)*p.n+i]));advance=true;if(last==p.n){p.k++;if(p.k>p.j){p.k=0;p.pass++;}if(p.pass==2){p.largest=0;p.go(13);}else{p.dot=0;p.compensation=0;p.go(11);}}
   }else if(p.phase==15){for(int i=first;i<last;i++)p.basis[size_t(p.j+1)*p.n+i]=stored(rpx_fp_div(p.vector[i],p.nextnorm));advance=true;if(last==p.n)p.go(16);
   }else if(p.phase==16){
    for(int k=0;k<p.j;k++){double v=stored(rpx_fp_dot2(p.cosine[k],p.H(k,p.j),p.sine[k],p.H(k+1,p.j)));p.H(k+1,p.j)=stored(rpx_fp_dot2(-p.sine[k],p.H(k,p.j),p.cosine[k],p.H(k+1,p.j)));p.H(k,p.j)=v;}
    double magnitude=std::max(std::fabs(p.H(p.j,p.j)),std::fabs(p.H(p.j+1,p.j)));if(magnitude==0){p.go(17);}else{double v;c=rpx_krylov_givens(p.H(p.j,p.j),p.H(p.j+1,p.j),magnitude,&v);if(c)throw Failure{c};if(v==0)throw Failure{205};p.cosine[p.j]=stored(rpx_fp_div(p.H(p.j,p.j),v));p.sine[p.j]=stored(rpx_fp_div(p.H(p.j+1,p.j),v));p.H(p.j,p.j)=v;p.H(p.j+1,p.j)=0;p.rhs[p.j+1]=stored(rpx_fp_mul(-p.sine[p.j],p.rhs[p.j]));p.rhs[p.j]=stored(rpx_fp_mul(p.cosine[p.j],p.rhs[p.j]));p.used=p.j+1;p.steps++;
     if(p.nextnorm==0||!rpx_fp_greater_product(std::fabs(p.rhs[p.j+1]),1e-14,p.beta)||p.j==23)p.go(17);else{p.j++;p.go(6);}}
   }else if(p.phase==17){bool broken=p.used==0;for(int k=p.used-1;k>=0&&!broken;k--){double v=p.rhs[k];for(int j=k+1;j<p.used;j++)v=stored(rpx_fp_mulsub(v,p.H(k,j),p.coefficients[j]));if(p.H(k,k)==0)broken=true;else p.coefficients[k]=stored(rpx_fp_div(v,p.H(k,k)));}if(broken){p.final=true;p.go(0);}else p.go(18);
   }else if(p.phase==18){for(int i=first;i<last;i++)for(int j=0;j<p.used;j++){DD a={p.coefficients[j],0},b={p.preconditioned[size_t(j)*p.n+i],0};p.answer[i]=operation(0,p.answer[i],operation(2,a,b));}advance=true;if(last==p.n){if(++p.cycle==3)p.final=true;p.go(0);}
   }else throw Failure{318};
   if(c)throw Failure{c};if(advance&&last!=p.n)p.cursor=last;
   if(ms==0)return 0;
  }while(clock()<deadline);return 0;
 });
}
API rpx_krylov_result(uintptr_t h,uint64_t serial,DD*out,double*metrics){if(!out||!metrics)return 330;return checked([&](){auto&p=get(h);if(p.active||!p.complete||p.serial!=serial)throw Failure{318};std::copy(p.answer.begin(),p.answer.end(),out);metrics[0]=p.best;metrics[1]=p.steps;metrics[2]=p.best<=1e-11?1:0;metrics[3]=p.finitechecks;return 0;});}
API rpx_krylov_bytes(uintptr_t h,uint64_t*out){if(!out)return 330;return checked([&](){*out=get(h).bytes();return 0;});}
API rpx_krylov_live_handles(void){return static_cast<int>(jobs.size());}
