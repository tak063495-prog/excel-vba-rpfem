// Real native allocation/unknown injection; global C++ allocator is test-only.
#include <cstdlib>
#include <new>
#include <cstdio>
struct InjectedUnknown{};
static long long fail_at=-1,calls=0,live=0;static bool unknown=false;
void*operator new(size_t n){auto at=calls++;if(at==fail_at){if(unknown)throw InjectedUnknown{};throw std::bad_alloc();}void*p=std::malloc(n?n:1);if(!p)throw std::bad_alloc();live++;return p;}
void operator delete(void*p)noexcept{if(p){live--;std::free(p);}}
void operator delete(void*p,size_t)noexcept{::operator delete(p);}
void*operator new[](size_t n){return ::operator new(n);}
void operator delete[](void*p)noexcept{::operator delete(p);}
void operator delete[](void*p,size_t)noexcept{::operator delete(p);}
#if defined(TEST_COMPILE)
#include "../native/rpfem_compile.cpp"
static int handles(){return rpx_compile_live_handles();}
static int scenario(bool partial=false){
 int32_t ip[]={0,3,6},ids[]={0,2,1,0,2,1},ep[]={0,2},ei[]={0,2},perm[4]={},inv[4]={},done=0;double ev[]={.5,-.2};uintptr_t h=0;
 uintptr_t args[]={reinterpret_cast<uintptr_t>(ip),reinterpret_cast<uintptr_t>(ids),reinterpret_cast<uintptr_t>(ep),reinterpret_cast<uintptr_t>(ei),reinterpret_cast<uintptr_t>(ev)};
 int rc=rpx_compile_create(3,1,2,6,2,args,&h);
 while(!rc&&done!=1){rc=rpx_compile_step(h,0,&done);if(partial)break;}
 if(!rc&&!partial){rc=rpx_compile_order_begin(h,perm,inv,0);done=0;while(!rc&&done!=3)rc=rpx_compile_step(h,0,&done);if(!rc)rc=rpx_compile_permute_begin(h);done=0;while(!rc&&done!=2)rc=rpx_compile_step(h,0,&done);}
 if(h){int c=rpx_compile_destroy(h);if(c)return c;}return rc;
}
#elif defined(TEST_SCHUR)
#include "../native/rpfem_schur.cpp"
static int handles(){return rpx_schur_live_handles();}
static int scenario(bool partial=false){
 int32_t pos[]={0,1,2},hp[]={0,1,3,1,2,4,3,4,5},ptr0[]={0,2},rows[]={0,1},bp[]={6,7,8,10,11,12},sp[]={0,4},slots[]={0,1,1,2},inv[]={0,1},full[]={3,4},direct[15];
 int32_t ap[]={0,1,3,6,10,15},ai[]={0,0,1,0,1,2,0,1,2,3,0,1,2,3,4},cp[]={0,0,0},dummy=0,done=0,reason=0;
 double matrix[15]={},reduced[3]={},m[12]={},empty=0,dv[]={1,2},rhs[]={4,6,6,4,10},out[5]={};uintptr_t h=0;uint64_t serial=0;
 std::fill(direct,direct+15,-1);direct[9]=0;direct[13]=1;direct[14]=2;matrix[0]=4;matrix[2]=3;matrix[5]=2;matrix[9]=1;matrix[14]=2;
 auto address=[](const void*p){return reinterpret_cast<uintptr_t>(p);};
 uintptr_t args[]={address(pos),address(hp),address(ptr0),address(rows),address(bp),address(sp),address(slots),address(inv),address(full),address(direct)};
 int rc=rpx_schur_create(5,2,1,15,3,2,4,1,args,&h);if(!rc)rc=rpx_schur_factor_begin(h,1,1);
 while(!rc&&!done){rc=rpx_schur_factor_step(h,1,1,matrix,0,&done,&reason);if(partial)break;}
 if(!rc&&!partial){rc=rpx_schur_factor_result(h,1,1,reduced,m);uintptr_t factor[]={address(ap),address(ai),address(matrix),address(cp),address(&dummy),address(&empty),address(dv)};
  if(!rc)rc=rpx_schur_factor_commit(h,1,1,factor,0);if(!rc)rc=rpx_schur_solve_begin(h,1,1,rhs,1,1,&serial);done=0;
  while(!rc&&!done)rc=rpx_schur_solve_step(h,1,1,serial,0,&done);if(!rc)rc=rpx_schur_solve_result(h,1,1,serial,out,m);
  if(!rc){for(int i=0;i<5;i++)if(out[i]!=double(i+1))rc=999;if(m[3]!=0)rc=999;}
 }
 if(h){int c=rpx_schur_destroy(h);if(c)return c;}return rc;
}
#else
#include "../native/rpfem_krylov.cpp"
static int handles(){return rpx_krylov_live_handles();}
static int scenario(bool partial=false){
 constexpr int n=17;int32_t cp[n+1]={},idx[n],ip[n+1],dims[n],dummy=0,done=0;double diag[n],rhs0[n],scale[n],empty=0,m[4]={};DD bdd[n],answer[n]={};
 struct Coef{double hi,lo,sh,sl;};Coef coef[n],hh[n*9]={},sc[n],reg={};
 for(int i=0;i<n;i++){idx[i]=i;ip[i]=i;dims[i]=1;diag[i]=1;rhs0[i]=i+.25;bdd[i]={rhs0[i],0};scale[i]=1;coef[i]={1,0,1,0};hh[i]=coef[i];sc[i]=coef[i];}ip[n]=n;
 uintptr_t op[]={ptr(idx),ptr(sc),ptr(dims),ptr(idx),ptr(ip),ptr(idx),ptr(idx),ptr(coef),ptr(hh),ptr(&dummy),ptr(&dummy),ptr(&reg),ptr(&reg),ptr(&reg)},fa[]={ptr(cp),ptr(&dummy),ptr(&empty),ptr(diag),ptr(scale)};
 uintptr_t h=0;int rc=rpx_krylov_create(n,n,n,n,&h);if(!rc)rc=rpx_krylov_begin(h,1,7,0,0,bdd,answer,1e300);
 while(!rc&&!done){rc=rpx_krylov_step(h,1,7,op,fa,0,0,&done);if(partial)break;}
 if(!rc&&!partial){rc=rpx_krylov_result(h,1,answer,m);if(!rc&&m[2]!=1)rc=999;}
 if(h){int c=rpx_krylov_destroy(h);if(c)return c;}return rc;
}
#endif
int main(){
 if(scenario())return 1;auto before=live;calls=0;if(scenario()||live!=before||handles())return 2;auto count=calls;
 for(long long i=0;i<count;i++){calls=0;fail_at=i;int rc=scenario();fail_at=-1;if(rc!=207||live!=before||handles()){std::printf("FAIL allocation %lld rc=%d live=%lld/%lld handles=%d\n",i,rc,live,before,handles());return 3;}}
 unknown=true;for(long long i:{0LL,1LL,count/2,count-1}){calls=0;fail_at=i;int rc=scenario();fail_at=-1;if(rc!=208||live!=before||handles())return 4;}unknown=false;
 if(scenario(true)||live!=before||handles())return 5;
 std::printf("PASS %lld allocation points207; 4 unknown points208; partial cleanup; live handles0\n",count);return 0;
}
