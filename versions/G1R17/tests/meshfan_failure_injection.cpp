// Test-only allocator: real DLL planner bad_alloc/unknown exceptions and cleanup.
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
#include "../native/rpfem_meshfan.cpp"
static int scenario(bool partial=false){
 double xy[]={0,0,1,0,-1,1},cx[1]={},cy[1]={};int32_t tri[]={0,1,2},eligible[]={1},body[]={1};
 int32_t edges[]={0,-1,0,1,-1,-1,1,0,-1,1,2,-1,-1,0,0,-1,2,0,-1,-1,0},done=0,count=0,split[1]={};uintptr_t h=0;
 int rc=rpx_meshfan_create(3,1,3,xy,tri,eligible,body,edges,&h);
 while(!rc&&!done){rc=rpx_meshfan_step(h,0,&done);if(partial)break;}
 if(!rc&&!partial){rc=rpx_meshfan_result(h,1,split,cx,cy,&count);if(!rc&&(count!=1||split[0]!=0||cx[0]<=0||cy[0]<=0))rc=999;}
 if(h){int c=rpx_meshfan_destroy(h);if(c)return c;}return rc;
}
int main(){
 if(scenario())return 1;auto before=live;calls=0;if(scenario()||live!=before||rpx_meshfan_live_handles())return 2;auto count=calls;
 for(long long i=0;i<count;i++){calls=0;fail_at=i;int rc=scenario();fail_at=-1;if(rc!=207||live!=before||rpx_meshfan_live_handles()){std::printf("FAIL allocation %lld rc=%d live=%lld/%lld\n",i,rc,live,before);return 3;}}
 unknown=true;for(long long i:{0LL,1LL,count/2,count-1}){calls=0;fail_at=i;int rc=scenario();fail_at=-1;if(rc!=208||live!=before||rpx_meshfan_live_handles())return 4;}unknown=false;
 if(scenario(true)||live!=before||rpx_meshfan_live_handles())return 5;
 std::printf("PASS %lld allocation points207; 4 unknown points208; partial cleanup; live handles0\n",count);return 0;
}
