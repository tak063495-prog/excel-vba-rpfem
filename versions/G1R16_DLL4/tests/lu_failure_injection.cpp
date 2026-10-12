// Test-only allocation and unknown-exception injection, never in the DLL.
#include <cstdlib>
#include <new>
#include <cstdio>
struct InjectedUnknown {};
static long long fail_at=-1,calls=0,live=0;
static bool unknown=false;
void* operator new(size_t n) {long long at=calls++;if(at==fail_at){if(unknown)throw InjectedUnknown{};throw std::bad_alloc();}void*p=std::malloc(n?n:1);if(!p)throw std::bad_alloc();live++;return p;}
void operator delete(void*p) noexcept {if(p){live--;std::free(p);}}
void operator delete(void*p,size_t) noexcept {::operator delete(p);}
void* operator new[](size_t n) {return ::operator new(n);}
void operator delete[](void*p) noexcept {::operator delete(p);}
void operator delete[](void*p,size_t) noexcept {::operator delete(p);}
#include "../native/rpfem_lu.cpp"
static int scenario(int stream,int stop=-1) {
    constexpr int n=17;int32_t ap[n+1]={0},ai[n*n]={0},permutation[n],written=0,lc=0,uc=0;double av[n*n]={0};int nz=0;uintptr_t handle=0;
    for(int c=0;c<n;c++){permutation[c]=n-c-1;for(int r=0;r<n;r++)if(r==c||(r*13+c*17)%11<5){ai[nz]=r;av[nz++]=r==c?20:double((r+c)%5)-2;}ap[c+1]=nz;}
    int rc=rpx_lu_create(n,nz,1e9,0,42,permutation,stream,.1,&handle);
    for(int phase=0;phase<5&&rc==0;phase++)rc=rpx_lu_input(handle,phase,0,n,ap,ai,av,0,1);
    while(rc==0&&written<n){rc=rpx_lu_step(handle,3,&written);if(stop>=0&&written>=stop)break;}
    if(rc==0&&stop<0)rc=rpx_lu_finish(handle,&lc,&uc);
    if(handle){int d=rpx_lu_destroy(handle);if(d!=0)return d;}
    return rc;
}
int main() {
    long long total=0;int unknown_points=0;
    for(int stream=0;stream<=1;stream++){
        if(scenario(stream)!=0)return 1;long long before=live;calls=0;
        if(scenario(stream)!=0||live!=before)return 2;long long count=calls;total+=count;
        for(long long at=0;at<count;at++){calls=0;fail_at=at;int rc=scenario(stream);fail_at=-1;if(rc!=207||live!=before||rpx_lu_live_handles()!=0){std::printf("FAIL allocation %lld code=%d live=%lld/%lld\n",at,rc,live,before);return 3;}}
        unknown=true;
        for(long long at:{0LL,1LL,10LL,count-1}){calls=0;fail_at=at;int rc=scenario(stream);fail_at=-1;if(rc!=208||live!=before||rpx_lu_live_handles()!=0)return 4;unknown_points++;}
        unknown=false;for(int stop=0;stop<17;stop+=3)if(scenario(stream,stop)!=0||live!=before||rpx_lu_live_handles()!=0)return 5;
    }
    int32_t order[2]={0,1},ap[3]={0,1,2},ai[2]={0,1};double av[2]={1,1};uintptr_t h=0;
    if(rpx_lu_create(2,2,1,0,1,order,0,0,&h)!=311||h)return 6;
    if(rpx_lu_create(2,2,1e9,0,1,order,0,1.1,&h)!=317||h)return 7;
    order[1]=0;if(rpx_lu_create(2,2,1e9,0,1,order,0,0,&h)!=314||h)return 8;order[1]=1;
    if(rpx_lu_create(2,2,1e9,0,1,order,0,0,&h)!=0)return 9;
    ai[1]=2;if(rpx_lu_input(h,0,0,2,ap,ai,av,0,0)!=312)return 10;
    if(rpx_lu_destroy(h)!=0||rpx_lu_destroy(h)!=318||rpx_lu_live_handles()!=0)return 11;
    std::printf("PASS %lld allocation-failure points code207; %d unknown points code208; scoped context no leaks; partial-step cleanup; memory/order/storage/threshold rejection\n",total,unknown_points);
    return 0;
}
