// Failure injection for the unchanged production ordering source.
// Test-only global allocator fails at each allocation; it is never in the DLL.
#include <cstdlib>
#include <new>
#include <cstdio>
struct InjectedUnknown {};
static long long failAt=-1,calls=0,live=0;
static bool unknownFailure=false;
void* operator new(size_t size) {
    long long at=calls++;
    if(at==failAt) {
        if(unknownFailure) throw InjectedUnknown{};
        throw std::bad_alloc();
    }
    void* p=std::malloc(size?size:1); if(!p) throw std::bad_alloc();
    ++live; return p;
}
void operator delete(void* p) noexcept { if(p) { --live; std::free(p); } }
void operator delete(void* p,size_t) noexcept { ::operator delete(p); }
void* operator new[](size_t size) { return ::operator new(size); }
void operator delete[](void* p) noexcept { ::operator delete(p); }
void operator delete[](void* p,size_t) noexcept { ::operator delete(p); }
#include "../native/rpfem_order.cpp"
static int scenario(int mode,int stopAfter=-1) {
    constexpr int n=17; int ap[n+1]={0},ai[n*n]={0},out[n]={0},size=0,written=0;
    for(int i=0;i<n;++i) {
        for(int j=n-1;j>=0;--j) if((i*13+j*17)%11<5) ai[size++]=j;
        ap[i+1]=size;
    }
    void* handle=nullptr; int rc=rpx_order_create(n,mode,&handle);
    if(rc==0) rc=rpx_order_feed(handle,0,n,ap,n+1,ai,size);
    if(rc==0) rc=rpx_order_start(handle);
    while(rc==0 && written<n) {
        rc=rpx_order_step(handle,3,&written,out,n);
        if(stopAfter>=0 && written>=stopAfter) break;
    }
    rpx_order_destroy(handle); return rc;
}
int main() {
    long long total=0;
    for(int mode=0;mode<=1;++mode) {
        calls=0; const long long before=live;
        int initial=scenario(mode);
        if(initial!=0 || live!=before) { std::printf("FAIL initial mode=%d rc=%d live=%lld before=%lld calls=%lld\n",mode,initial,live,before,calls); return 1; }
        const long long count=calls; total+=count;
        for(long long at=0;at<count;++at) {
            calls=0; failAt=at;
            int rc=scenario(mode); failAt=-1;
            if(rc!=207 || live!=before) {
                std::printf("FAIL mode=%d allocation=%lld code=%d live=%lld expected=%lld\n",mode,at,rc,live,before); return 2;
            }
        }
        unknownFailure=true;
        for(long long at: {0LL,1LL,10LL,count-1}) {
            calls=0; failAt=at;
            int rc=scenario(mode); failAt=-1;
            if(rc!=208 || live!=before) return 15;
        }
        unknownFailure=false;
        for(int stop=0;stop<17;stop+=3) if(scenario(mode,stop)!=0 || live!=before) return 3;
    }
    void* handle=nullptr; int ap[3]={0,1,1},ai[1]={2},out[2]={0},written=0;
    if(rpx_order_create(0,0,&handle)!=100 || handle) return 4;
    if(rpx_order_create(2,2,&handle)!=100 || handle) return 5;
    if(rpx_order_create(2,0,&handle)!=0) return 6;
    if(rpx_order_start(handle)!=100) return 7;
    if(rpx_order_feed(handle,0,2,ap,3,ai,1)!=100) return 8;
    rpx_order_destroy(handle); handle=nullptr;
    ap[1]=0; ap[2]=0;
    if(rpx_order_create(2,0,&handle)!=0) return 9;
    if(rpx_order_feed(handle,0,2,ap,3,ai,0)!=0 || rpx_order_start(handle)!=0) return 10;
    if(rpx_order_step(handle,3,&written,out,1)!=100 || written) return 11;
    if(rpx_order_step(handle,0,&written,out,2)!=100 || written) return 12;
    if(rpx_order_step(handle,3,&written,out,2)!=0 || written!=2) return 13;
    rpx_order_destroy(handle);
    if(live!=0) return 14;
    std::printf("PASS %lld allocation-failure points code207; 8 unknown exceptions code208; no leaks; cleanup after partial steps; invalid contracts rejected\n",total);
    return 0;
}
