/* DLL5 typed sparse model compiler. Keeps dictionary insertion order, cone
 * pattern ownership, normalized equality addition order and exact cache key.
 * Ordering/tie rules are the unchanged DLL3 kernel. No numeric factor reuse.
 * LGPL-2.1-or-later (ordering); no borrowed pointers survive a call. */
#define NOMINMAX
#include <windows.h>
#include <cstdint>
#include <vector>
#include <unordered_map>
#include <map>
#include <memory>
#include <algorithm>
#include <string>
#include <new>
#include <cstring>
#include <cmath>
extern "C" int __stdcall rpx_order_create(int,int,void**),rpx_order_feed(void*,int,int,const int*,int,const int*,int),rpx_order_start(void*),rpx_order_step(void*,int,int*,int*,int),rpx_order_destroy(void*);
extern "C" double rpx_fp_add(double,double);
#define API extern "C" __declspec(dllexport) int __stdcall
namespace {
struct Failure{int code;};
double clock(){LARGE_INTEGER q,f;QueryPerformanceCounter(&q);QueryPerformanceFrequency(&f);return double(q.QuadPart)/double(f.QuadPart);}
struct Column{
 std::vector<int32_t> keys;std::unordered_map<int32_t,int32_t> at;
 void add(int32_t k){if(at.find(k)!=at.end())return;at.emplace(k,int32_t(keys.size()));keys.push_back(k);}
};
struct Context{
 int n,nx,nm,nb,phase=0,cursor=0,nnz=0,written=0;bool pending=false,complete=false,hit=false,convertPending=false;void*order=nullptr;
 std::vector<int32_t> ip,ids,ep,ei,pattern,rawap,rawai,perm,inv,ap,ai,diag;
 std::vector<double> eval,base;std::vector<Column> cols,ordered;
 std::map<std::vector<int32_t>,int32_t> patterns;std::wstring signature;
 Context(int x,int m,int b,const int32_t*cip,const int32_t*cids,int cn,const int32_t*eap,const int32_t*eid,const double*ev,int en):n(x+m),nx(x),nm(m),nb(b),ip(cip,cip+b+1),ids(cids,cids+cn),ep(eap,eap+m+1),ei(eid,eid+en),pattern(b),rawap(n+1),perm(n),inv(n),ap(n+1),diag(n),eval(ev,ev+en),cols(n),ordered(n){
  if(ip[0]!=0||ip[nb]!=cn||ep[0]!=0||ep[nm]!=en)throw Failure{330};
  for(int i=0;i<nb;i++)if(ip[i]<0||ip[i+1]<ip[i])throw Failure{330};for(int i=0;i<nm;i++)if(ep[i]<0||ep[i+1]<ep[i])throw Failure{330};
  for(int i:ids)if(i<0||i>=nx)throw Failure{330};for(int i:ei)if(i<0||i>=nx)throw Failure{330};
  for(int i=0;i<n;i++)cols[i].add(i);
 }
 ~Context(){if(order)rpx_order_destroy(order);}
 void go(int p){phase=p;cursor=0;}
 size_t bytes()const{
  size_t out=sizeof(*this)+4*(ip.capacity()+ids.capacity()+ep.capacity()+ei.capacity()+pattern.capacity()+rawap.capacity()+rawai.capacity()+perm.capacity()+inv.capacity()+ap.capacity()+ai.capacity()+diag.capacity())+8*(eval.capacity()+base.capacity())+2*signature.capacity();
  auto count=[&](const std::vector<Column>&v){size_t b=v.capacity()*sizeof(Column);for(auto&c:v)b+=4*c.keys.capacity()+32*c.at.size()+sizeof(void*)*c.at.bucket_count();return b;};return out+count(cols)+count(ordered)+64*patterns.size()+4*ids.size();
 }
};
std::unordered_map<uintptr_t,std::unique_ptr<Context>> contexts;uintptr_t nextid=0;
Context&get(uintptr_t h){auto i=contexts.find(h);if(i==contexts.end())throw Failure{318};return *i->second;}
template<class F>int checked(F f){try{return f();}catch(Failure e){return e.code;}catch(const std::bad_alloc&){return 207;}catch(...){return 208;}}
void require(int c){if(c)throw Failure{c};}
}
API rpx_compile_create(int32_t nx,int32_t nm,int32_t nb,int32_t cn,int32_t en,const uintptr_t*a,uintptr_t*h){
 if(nx<1||nm<0||int64_t(nx)+nm>1048576||nb<1||cn<0||en<0||!a||!h)return 330;*h=0;for(int i=0;i<5;i++)if(!a[i])return 330;
 return checked([&](){auto p=std::make_unique<Context>(nx,nm,nb,reinterpret_cast<const int32_t*>(a[0]),reinterpret_cast<const int32_t*>(a[1]),cn,reinterpret_cast<const int32_t*>(a[2]),reinterpret_cast<const int32_t*>(a[3]),reinterpret_cast<const double*>(a[4]),en);if(nextid==UINTPTR_MAX)throw Failure{318};uintptr_t id=++nextid;contexts.emplace(id,std::move(p));*h=id;return 0;});
}
API rpx_compile_destroy(uintptr_t h){return checked([&](){auto i=contexts.find(h);if(i==contexts.end())throw Failure{318};contexts.erase(i);return 0;});}
API rpx_compile_step(uintptr_t h,int32_t ms,int32_t*done){
 if(!done||ms<0||ms>1000)return 330;*done=0;
 return checked([&](){auto&p=get(h);if(p.pending){*done=1;return 0;}if(p.convertPending){*done=3;return 0;}if(p.complete){*done=2;return 0;}double deadline=clock()+double(ms)/1000;int units=0;
  do{int i=p.cursor;
   if(p.phase==0){std::vector<int32_t>key(p.ids.begin()+p.ip[i],p.ids.begin()+p.ip[i+1]);auto found=p.patterns.find(key);
    if(found!=p.patterns.end())p.pattern[i]=found->second;else{p.patterns.emplace(key,i);p.pattern[i]=i;for(size_t j=0;j<key.size();j++)for(size_t k=0;k<=j;k++){int id=key[k],col=key[j];if(id>col)std::swap(id,col);p.cols[col].add(id);}}
    if(++p.cursor==p.nb)p.go(1);
   }else if(p.phase==1){if(p.nm>0)for(int at=p.ep[i];at<p.ep[i+1];at++)p.cols[p.nx+i].add(p.ei[at]);if(p.nm==0||++p.cursor==p.nm)p.go(2);
   }else if(p.phase==2){p.rawap[i]=int32_t(p.rawai.size());for(int k:p.cols[i].keys)p.rawai.push_back(k);auto sorted=p.cols[i].keys;std::sort(sorted.begin(),sorted.end());if(i==0)p.signature=std::to_wstring(p.n)+L":";else p.signature+=L";";
    if(sorted.empty())p.signature+=L"-";else for(size_t k=0;k<sorted.size();k++){if(k)p.signature+=L",";p.signature+=std::to_wstring(sorted[k]);}
    if(++p.cursor==p.n){if(p.rawai.size()>100000000)throw Failure{207};p.rawap[p.n]=int32_t(p.rawai.size());p.pending=true;p.go(3);*done=1;return 0;}
   }else if(p.phase==4){int last=std::min(p.n,i+128);require(rpx_order_feed(p.order,i,last,p.rawap.data(),p.n+1,p.rawai.data(),int32_t(p.rawai.size())));p.cursor=last;if(last==p.n)p.go(5);
   }else if(p.phase==5){require(rpx_order_start(p.order));p.go(6);
   }else if(p.phase==6){require(rpx_order_step(p.order,64,&p.written,p.perm.data(),p.n));if(p.written==p.n){for(int k=0;k<p.n;k++)p.inv[p.perm[k]]=k;require(rpx_order_destroy(p.order));p.order=nullptr;p.go(7);p.convertPending=true;*done=3;return 0;}
   }else if(p.phase==7){for(int key:p.cols[i].keys){int col=p.inv[i],id=p.inv[key];if(id>col)std::swap(id,col);p.ordered[col].add(id);}p.cols[i]=Column();if(++p.cursor==p.n)p.go(8);
   }else if(p.phase==8){p.ap[i]=int32_t(p.ai.size());for(int key:p.ordered[i].keys){if(key==i)p.diag[i]=int32_t(p.ai.size());p.ai.push_back(key);}if(++p.cursor==p.n){if(p.ai.size()>100000000)throw Failure{207};p.ap[p.n]=int32_t(p.ai.size());p.nnz=int32_t(p.ai.size());p.base.resize(p.nnz,0.0);p.go(9);}
   }else if(p.phase==9){if(p.nm>0)for(int at=p.ep[i];at<p.ep[i+1];at++){int col=p.inv[p.nx+i],id=p.inv[p.ei[at]];if(id>col)std::swap(id,col);auto found=p.ordered[col].at.find(id);if(found==p.ordered[col].at.end())throw Failure{318};int pos=p.ap[col]+found->second;p.base[pos]=rpx_fp_add(p.base[pos],p.eval[at]);if(!std::isfinite(p.base[pos]))throw Failure{206};}if(p.nm==0||++p.cursor==p.nm){p.complete=true;*done=2;return 0;}
   }else throw Failure{318};
   if(ms==0&&(++units>=128||p.phase==4||p.phase==5||p.phase==6))return 0;
  }while(ms==0||clock()<deadline);return 0;
 });
}
API rpx_compile_signature(uintptr_t h,wchar_t*out,int32_t capacity,int32_t*written){if(!written||capacity<0)return 330;return checked([&](){auto&p=get(h);if(!p.pending)throw Failure{318};if(p.signature.size()>INT32_MAX)throw Failure{206};*written=int32_t(p.signature.size());if(capacity==0&&!out)return 0;if(!out||capacity<*written)throw Failure{330};std::copy(p.signature.begin(),p.signature.end(),out);return 0;});}
API rpx_compile_order_begin(uintptr_t h,const int32_t*perm,const int32_t*inv,int32_t hit){if(hit!=0&&hit!=1)return 330;return checked([&](){auto&p=get(h);if(!p.pending||p.phase!=3)throw Failure{318};p.hit=hit!=0;p.pending=false;if(hit){if(!perm||!inv)throw Failure{330};std::vector<bool>seen(p.n);for(int k=0;k<p.n;k++){if(perm[k]<0||perm[k]>=p.n||seen[perm[k]]||inv[perm[k]]!=k)throw Failure{330};seen[perm[k]]=true;}p.perm.assign(perm,perm+p.n);p.inv.assign(inv,inv+p.n);p.go(7);p.convertPending=true;}else{require(rpx_order_create(p.n,0,&p.order));p.go(4);}return 0;});}
API rpx_compile_permute_begin(uintptr_t h){return checked([&](){auto&p=get(h);if(!p.convertPending||p.phase!=7)throw Failure{318};p.convertPending=false;return 0;});}
API rpx_compile_counts(uintptr_t h,int32_t*out){if(!out)return 330;return checked([&](){auto&p=get(h);if(!p.complete)throw Failure{318};out[0]=p.nnz;out[1]=p.hit?1:0;return 0;});}
API rpx_compile_export(uintptr_t h,int32_t nz,const uintptr_t*a){if(!a)return 330;for(int i=0;i<7;i++)if(!a[i])return 330;return checked([&](){auto&p=get(h);if(!p.complete||nz!=p.nnz)throw Failure{318};auto copy=[&](const std::vector<int32_t>&v,int index){std::copy(v.begin(),v.end(),reinterpret_cast<int32_t*>(a[index]));};copy(p.perm,0);copy(p.inv,1);copy(p.ap,2);copy(p.ai,3);copy(p.diag,4);std::copy(p.base.begin(),p.base.end(),reinterpret_cast<double*>(a[5]));copy(p.pattern,6);return 0;});}
API rpx_compile_slots(uintptr_t h,int32_t cone,int32_t size,int32_t*out){if(!out||size<0)return 330;return checked([&](){auto&p=get(h);if(!p.complete||cone<0||cone>=p.nb)throw Failure{318};int count=p.ip[cone+1]-p.ip[cone];if(int64_t(count)*count!=size||p.pattern[cone]!=cone)throw Failure{330};for(int j=0;j<count;j++)for(int k=0;k<=j;k++){int col=p.inv[p.ids[p.ip[cone]+j]],id=p.inv[p.ids[p.ip[cone]+k]];if(id>col)std::swap(id,col);auto found=p.ordered[col].at.find(id);if(found==p.ordered[col].at.end())throw Failure{318};out[j+count*k]=p.ap[col]+found->second;}return 0;});}
API rpx_compile_bytes(uintptr_t h,uint64_t*out){if(!out)return 330;return checked([&](){*out=get(h).bytes();return 0;});}
API rpx_compile_live_handles(void){return int(contexts.size());}
