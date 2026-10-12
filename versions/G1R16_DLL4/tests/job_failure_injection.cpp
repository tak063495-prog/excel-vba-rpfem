// Test-only allocation and unknown-exception injection for persistent jobs.
#include <cstdlib>
#include <new>
#include <cstdio>
#include <vector>
struct InjectedUnknown {};
static long long fail_at=-1,calls=0,live=0;static bool unknown=false;
void* operator new(size_t n){auto at=calls++;if(at==fail_at){if(unknown)throw InjectedUnknown{};throw std::bad_alloc();}void*p=std::malloc(n?n:1);if(!p)throw std::bad_alloc();live++;return p;}
void operator delete(void*p) noexcept {if(p){live--;std::free(p);}}
void operator delete(void*p,size_t) noexcept {::operator delete(p);}
void*operator new[](size_t n){return ::operator new(n);}
void operator delete[](void*p) noexcept {::operator delete(p);}
void operator delete[](void*p,size_t) noexcept {::operator delete(p);}
#ifdef TEST_HLINEAR
#include "../native/rpfem_hlinear.cpp"
#else
#include "../native/rpfem_refine.cpp"
#endif
static int scenario(bool partial=false) {
 constexpr int n=17;
 int32_t ap[n+1],ai[n],cp[n+1]={},dummy=0;
 double av[n],diag[n],rhs[n],out[n],scale[n],residual=0;int32_t done=0,checks=0,corrections=0;
 uintptr_t h=0;
 for(int i=0;i<n;i++){ap[i]=i;ai[i]=i;av[i]=1;diag[i]=1;rhs[i]=i+.25;scale[i]=1;}ap[n]=n;
#ifdef TEST_HLINEAR
 (void)checks;(void)corrections;(void)residual;(void)out;
 int32_t dims[n],starts[n],ip[n+1];
 struct Coef {double hi,lo,sh,sl;};Coef co[n],hh[n*9]={},sc[n],reg={};DD best[n],bdd[n];double metrics[14]={};
 for(int i=0;i<n;i++){dims[i]=1;starts[i]=i;ip[i]=i;co[i]={1,0,1,0};hh[i]={1,0,1,0};sc[i]=co[i];}ip[n]=n;
 uintptr_t op[14]={ptr(ai),ptr(sc),ptr(dims),ptr(starts),ptr(ip),ptr(ai),ptr(starts),ptr(co),ptr(hh),ptr(&dummy),ptr(&dummy),ptr(&reg),ptr(&reg),ptr(&reg)};
 uintptr_t init[3]={ptr(rhs),ptr(ai),ptr(sc)},factor[5]={ptr(cp),ptr(&dummy),ptr(av),ptr(diag),ptr(scale)};
 int rc=rpx_hlinear_create(n,n,n,n,&h);if(!rc)rc=rpx_hlinear_begin(h,1,42,0,0,init);
 while(!rc&&!done){rc=rpx_hlinear_step(h,1,42,op,factor,0,1,&done);if(partial)break;}
 if(!rc&&!partial){rc=rpx_hlinear_result(h,1,best,out,bdd,metrics);if(!rc){for(int i=0;i<n;i++)if(best[i].hi!=rhs[i]||best[i].lo!=0)return 999;}}
 if(h){int code=rpx_hlinear_destroy(h);if(code)return code;}
#else
 (void)scale;
 uintptr_t args[7]={reinterpret_cast<uintptr_t>(ap),reinterpret_cast<uintptr_t>(ai),reinterpret_cast<uintptr_t>(av),reinterpret_cast<uintptr_t>(cp),reinterpret_cast<uintptr_t>(&dummy),reinterpret_cast<uintptr_t>(av),reinterpret_cast<uintptr_t>(diag)};
 int rc=rpx_refine_create(n,&h);if(!rc)rc=rpx_refine_begin(h,n,1,rhs);
 while(!rc&&!done){rc=rpx_refine_step(h,n,1,args,1,&done);if(partial)break;}
 if(!rc&&!partial){rc=rpx_refine_result(h,n,1,out,&checks,&corrections,&residual);if(!rc){for(int i=0;i<n;i++)if(out[i]!=rhs[i])return 999;}}
 if(h){int code=rpx_refine_destroy(h);if(code)return code;}
#endif
 return rc;
}
static int handles(){
#ifdef TEST_HLINEAR
 return rpx_hlinear_live_handles();
#else
 return rpx_refine_live_handles();
#endif
}
int main(){
 if(scenario())return 1;auto before=live;calls=0;if(scenario()||live!=before)return 2;auto count=calls;
 for(long long i=0;i<count;i++){calls=0;fail_at=i;int rc=scenario();fail_at=-1;if(rc!=207||live!=before||handles()){std::printf("FAIL allocation %lld rc=%d live=%lld/%lld\n",i,rc,live,before);return 3;}}
 unknown=true;for(long long i:{0LL,1LL,count/2,count-1}){calls=0;fail_at=i;int rc=scenario();fail_at=-1;if(rc!=208||live!=before||handles())return 4;}unknown=false;
 if(scenario(true)||live!=before||handles())return 5;
 uintptr_t old=0,second=0;
#ifdef TEST_HLINEAR
 if(rpx_hlinear_create(2,2,2,2,&old)||rpx_hlinear_destroy(old)||rpx_hlinear_create(2,2,2,2,&second))return 6;
 if(old==second||rpx_hlinear_destroy(old)!=318||rpx_hlinear_destroy(second)||handles())return 7;
#else
 if(rpx_refine_create(2,&old)||rpx_refine_destroy(old)||rpx_refine_create(2,&second))return 6;
 if(old==second||rpx_refine_destroy(old)!=318||rpx_refine_destroy(second)||handles())return 7;
#endif
 std::printf("PASS %lld allocation points207; 4 unknown points208; partial cleanup; stale handles rejected; exact identity solve\n",count);return 0;
}
