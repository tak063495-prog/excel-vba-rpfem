/* DLL4 resumable ordinary LDL solve + checked refinement.
 * Owns workspace, never retains borrowed VBA pointers. LGPL-2.1-or-later. */
#define NOMINMAX
#include <windows.h>
#include <cstdint>
#include <vector>
#include <unordered_map>
#include <new>
#include <cstring>
#include <algorithm>
extern "C" int __stdcall rpx_backsolve(int32_t,int32_t,int32_t,int32_t,const int32_t*,const int32_t*,const double*,const double*,const double*,double*);
extern "C" int __stdcall rpx_solver_vectors(int32_t,int32_t,int32_t,int32_t,const uintptr_t*,double,double*);
#define API extern "C" __declspec(dllexport) int __stdcall
namespace {
struct Refine {
 int32_t n,stage=0,cursor=0,checks=0,corrections=0;uint64_t epoch=0;
 double residual=0;bool active=false;
 std::vector<double> rhs,answer,res,delta;
 explicit Refine(int32_t size):n(size),rhs(size),answer(size),res(size),delta(size){}
};
std::unordered_map<uintptr_t,Refine*> jobs;uintptr_t next_id=0;
struct Failure {int code;};
Refine* get(uintptr_t id){auto i=jobs.find(id);if(i==jobs.end())throw Failure{318};return i->second;}
template<class F>int checked(F f){try{return f();}catch(Failure e){return e.code;}catch(const std::bad_alloc&){return 207;}catch(...){return 208;}}
}
API rpx_refine_create(int32_t n,uintptr_t *handle){if(!handle||n<1)return 100;*handle=0;
 return checked([&](){Refine*p=new Refine(n);try{if(next_id==UINTPTR_MAX)throw Failure{318};uintptr_t id=++next_id;jobs.emplace(id,p);*handle=id;}catch(...){delete p;throw;}return 0;});}
API rpx_refine_destroy(uintptr_t handle){return checked([&](){auto p=get(handle);jobs.erase(handle);delete p;return 0;});}
API rpx_refine_begin(uintptr_t handle,int32_t n,uint64_t epoch,const double *rhs){if(!rhs)return 100;
 return checked([&](){auto p=get(handle);if(p->n!=n||epoch==0||p->active)throw Failure{318};
  std::memcpy(p->rhs.data(),rhs,size_t(n)*8);std::memcpy(p->answer.data(),rhs,size_t(n)*8);
  p->epoch=epoch;p->stage=1;p->cursor=0;p->checks=0;p->corrections=0;p->residual=0;p->active=true;return 0;});}
API rpx_refine_step(uintptr_t handle,int32_t n,uint64_t epoch,const uintptr_t *args,int32_t milliseconds,int32_t *done){
 if(!args||!done||milliseconds<1||milliseconds>1000)return 100;for(int i=0;i<7;i++)if(!args[i])return 100;*done=0;
 return checked([&](){auto p=get(handle);if(p->n!=n||p->epoch!=epoch||!p->active)throw Failure{318};LARGE_INTEGER freq,clock;QueryPerformanceFrequency(&freq);QueryPerformanceCounter(&clock);LONGLONG deadline=clock.QuadPart+freq.QuadPart*milliseconds/1000;
  auto ap=reinterpret_cast<const int32_t*>(args[0]),ai=reinterpret_cast<const int32_t*>(args[1]);auto av=reinterpret_cast<const double*>(args[2]);
  auto cp=reinterpret_cast<const int32_t*>(args[3]),ri=reinterpret_cast<const int32_t*>(args[4]);auto lv=reinterpret_cast<const double*>(args[5]),dv=reinterpret_cast<const double*>(args[6]);
  do {
   int32_t first=p->cursor,last=std::min(n,first+32);int code=0;double metric[1]={p->residual};
   if(p->stage>=1&&p->stage<=3){int phase=p->stage;double*x=p->answer.data();if(phase==3){last=n-first;first=std::max(0,last-32);}code=rpx_backsolve(n,phase,first,last,cp,ri,lv,dv,p->rhs.data(),x);}
   else if(p->stage==4){std::fill(p->res.begin(),p->res.end(),0.0);p->residual=0;p->cursor=n;}
   else if(p->stage==5){uintptr_t a[5]={reinterpret_cast<uintptr_t>(ap),reinterpret_cast<uintptr_t>(ai),reinterpret_cast<uintptr_t>(av),reinterpret_cast<uintptr_t>(p->answer.data()),reinterpret_cast<uintptr_t>(p->res.data())};code=rpx_solver_vectors(11,n,first,last,a,0,metric);}
   else if(p->stage==6){uintptr_t a[2]={reinterpret_cast<uintptr_t>(p->rhs.data()),reinterpret_cast<uintptr_t>(p->res.data())};code=rpx_solver_vectors(12,n,first,last,a,0,metric);p->residual=metric[0];}
   else if(p->stage==7){std::copy(p->res.begin(),p->res.end(),p->delta.begin());p->cursor=n;}
   else if(p->stage>=8&&p->stage<=10){int phase=p->stage-7;double*x=p->delta.data();if(phase==3){last=n-first;first=std::max(0,last-32);}code=rpx_backsolve(n,phase,first,last,cp,ri,lv,dv,p->res.data(),x);}
   else if(p->stage==11){uintptr_t a[2]={reinterpret_cast<uintptr_t>(p->answer.data()),reinterpret_cast<uintptr_t>(p->delta.data())};code=rpx_solver_vectors(13,n,first,last,a,0,metric);}
   else throw Failure{318};
   if(code){p->active=false;throw Failure{code};}
   p->cursor=std::min(n,p->cursor+32);
   if(p->cursor==n){p->cursor=0;
    if(p->stage==6){p->checks++;if(p->residual<1e-11){p->active=false;*done=1;return 0;}}
    if(p->stage==11){p->corrections++;if(p->checks==3){p->active=false;*done=1;return 0;}p->stage=4;}
    else p->stage++;
   }
   QueryPerformanceCounter(&clock);
  }while(clock.QuadPart<deadline);
  return 0;
 });
}
API rpx_refine_result(uintptr_t handle,int32_t n,uint64_t epoch,double *out,int32_t *checks,int32_t *corrections,double *residual){
 if(!out||!checks||!corrections||!residual)return 100;return checked([&](){auto p=get(handle);if(p->n!=n||p->epoch!=epoch||p->active||p->checks<1)throw Failure{318};
  std::memcpy(out,p->answer.data(),size_t(n)*8);*checks=p->checks;*corrections=p->corrections;*residual=p->residual;return 0;});}
API rpx_refine_live_handles(void){return static_cast<int>(jobs.size());}
