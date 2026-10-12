/* G1R16 DLL_3: exact-order port of RPX_HSDLu's sparse row-pivoted LU.
 * SPDX-License-Identifier: LGPL-2.1-or-later
 * A scoped context owns temporary graph storage only. Compressed factors are
 * exported to VBA and the context is destroyed before returning to the solver.
 */
#include <cstdint>
#include <vector>
#include <unordered_map>
#include <algorithm>
#include <cmath>
#include <cstring>
#include <new>
#include <limits>
#pragma fp_contract(off)
extern "C" double rpx_fp_mul(double,double);
extern "C" double rpx_fp_mul3(double,double,double);
extern "C" double rpx_fp_div(double,double);
extern "C" double rpx_fp_mulsub(double,double,double);
#define API extern "C" __declspec(dllexport) int __stdcall
namespace {
enum { OK=0,PIVOT=1,BAD=100,ALLOC=207,UNKNOWN=208,MEMORY=311,STORAGE=312,SIZE=313,
 ORDER=314,MATRIX=315,FACTOR=316,THRESHOLD=317,CONTEXT=318,RHS=319,SOLUTION=320,FP_OVERFLOW=206,FP_DIVZERO=205 };
struct Error { int code; };
bool finite(double x) {return x>=-1e290&&x<=1e290&&std::isfinite(x);}
struct Lu {
    int32_t n,nnz,epoch,capacity=0,used=0,free_head=0,live=0,occupied=0,mask=2047,k=0;
    int32_t extra_used=0,extra_capacity=0,u_used=0,u_capacity=0,stamp=0;
    bool streaming=false,pair_ready=false,ready=false,pivot_failed=false;
    double budget,overhead,pair_bytes=0,bytes=0,updates=0,checks=0,min_pivot=1e300,max_l=0,input_max=0,factor_max=0,threshold;
    double scan_nodes=0,scan_rows=0;
    std::vector<int32_t> nr,nc,rn,rp,cn,cp,rh,ch,slots,row_at,position,order,inverse,row_index,row_stamp,seen,pair;
    std::vector<double> nv,row_scale,col_scale;
    std::vector<int32_t> lp,li,up,ui,extra_head,extra_next,extra_col,u_col;
    std::vector<double> lv,uv,diagonal,extra_value,u_val;
    Lu(int32_t n0,int32_t nz,double maximum,double existing,int32_t epoch0,const int32_t *permutation,bool stream,double th)
      :n(n0),nnz(nz),epoch(epoch0),streaming(stream),budget(maximum),overhead(existing+8.0*n0),threshold(th) {
        if(n<1||n>1048576)throw Error{SIZE};if(nnz<0)throw Error{STORAGE};if(th<0||th>1)throw Error{THRESHOLD};
        memory(1024,2048);
        rh.resize(n);ch.resize(n);row_at.resize(n);position.resize(n);order.resize(n);inverse.resize(n);row_index.resize(n);row_stamp.resize(n);seen.resize(n);
        row_scale.resize(n);col_scale.resize(n);slots.resize(2048);u_col.resize(n+1);u_val.resize(n+1);
        for(int32_t i=0;i<n;i++){int32_t o=permutation[i];if(o<0||o>=n||seen[o])throw Error{ORDER};seen[o]=1;order[i]=o;inverse[o]=i;row_at[i]=i;position[i]=i;}
        std::fill(seen.begin(),seen.end(),-1);
        uint64_t count=uint64_t(n)*uint64_t(n);
        if(count<=2147483647ULL&&double(count)*4<=67108864.0&&double(count)*4+140000000.0<=budget){pair_bytes=double(count)*4;pair.resize(size_t(count));pair_ready=true;}
        memory(1024,2048);grow();if(streaming){extra_head.resize(n);up.resize(n+1);diagonal.resize(n);}
    }
    void memory(int32_t cap,int32_t hash) {
        double per_node=streaming?40.0:44.0;
        double factors=streaming?28.0*(extra_capacity+1.0)+24.0*(u_capacity+1.0):0;
        bytes=overhead+(cap+1.0)*per_node+hash*8.0+n*128.0+pair_bytes+factors;
        if(bytes>budget)throw Error{MEMORY};
    }
    int32_t bucket(int32_t r,int32_t c) const {
        double mixed=rpx_fp_mul(double(int64_t(r)*n+c+1),0.618033988749895);
        return int32_t(std::trunc((mixed-std::trunc(mixed))*double(mask+1)));
    }
    int32_t locate(int32_t r,int32_t c,int32_t &at) const {
        int32_t h=bucket(r,c),tomb=-1;
        for(int32_t probe=0;probe<=mask;probe++){
            int32_t id=slots[h];if(id==0){at=tomb>=0?tomb:h;return 0;}
            if(id==-1){if(tomb<0)tomb=h;}else if(nr[id]==r&&nc[id]==c){at=h;return id;}
            h=(h+1)&mask;
        }throw Error{STORAGE};
    }
    void rehash(int32_t size) {
        memory(capacity,size);slots.assign(size,0);mask=size-1;occupied=0;
        for(int32_t id=1;id<=used;id++)if(nr[id]>=0){int32_t h=bucket(nr[id],nc[id]);while(slots[h]!=0)h=(h+1)&mask;slots[h]=id;occupied++;}
    }
    void grow() {
        int64_t next=int64_t(capacity)*2;if(next<1024)next=1024;
        double node=streaming?40.0:44.0,factors=streaming?28.0*(extra_capacity+1.0)+24.0*(u_capacity+1.0):0;
        double affordable=std::trunc((budget-overhead-(mask+1)*8.0-n*128.0-pair_bytes-factors)/node)-1.0;
        if(affordable<double(next)){if(affordable<=capacity)throw Error{MEMORY};next=int64_t(affordable);}
        if(next>INT32_MAX-1)throw Error{FP_OVERFLOW};memory(int32_t(next),mask+1);
        size_t size=size_t(next)+1;nr.resize(size);nc.resize(size);nv.resize(size);rn.resize(size);rp.resize(size);cn.resize(size);cp.resize(size);capacity=int32_t(next);
    }
    void put(int32_t r,int32_t c,double value) {
        updates++;checks++;if(!finite(value))throw Error{FACTOR};int32_t h=0,id=locate(r,c,h),a,b;
        if(id>0){if(value!=0){nv[id]=value;return;}
            a=rp[id];b=rn[id];if(a>0)rn[a]=b;else rh[r]=b;if(b>0)rp[b]=a;
            a=cp[id];b=cn[id];if(a>0)cn[a]=b;else ch[c]=b;if(b>0)cp[b]=a;
            if(pair_ready)pair[size_t(r)*n+c]=0;slots[h]=-1;nr[id]=-1;rn[id]=free_head;free_head=id;live--;
        }else if(value!=0){
            if(occupied*10.0>=(mask+1)*7.0){int32_t size=mask+1;if(live*10.0>=size*4.0){if(size>INT32_MAX/2)throw Error{FP_OVERFLOW};size*=2;}rehash(size);id=locate(r,c,h);}
            if(free_head>0){id=free_head;free_head=rn[id];}else{if(used==capacity)grow();id=++used;}
            nr[id]=r;nc[id]=c;nv[id]=value;rp[id]=0;rn[id]=rh[r];if(rh[r]>0)rp[rh[r]]=id;rh[r]=id;
            cp[id]=0;cn[id]=ch[c];if(ch[c]>0)cp[ch[c]]=id;ch[c]=id;
            if(slots[h]==0)occupied++;slots[h]=id;live++;if(pair_ready)pair[size_t(r)*n+c]=id;
        }
    }
    void input(int32_t phase,int32_t first,int32_t last,const int32_t *ap,const int32_t *ai,const double *av,bool symmetric,bool equilibrate) {
        if(first<0||last<first||last>n||!ap||!ai||!av||ap[0]!=0||ap[n]!=nnz)throw Error{STORAGE};
        if(phase==1){for(int32_t r=first;r<last;r++)row_scale[r]=(equilibrate&&row_scale[r]>0)?rpx_fp_div(1,row_scale[r]):1;return;}
        if(phase==3){for(int32_t c=first;c<last;c++)col_scale[c]=(equilibrate&&col_scale[c]>0)?rpx_fp_div(1,col_scale[c]):1;return;}
        for(int32_t c=first;c<last;c++){
            if(ap[c]<0||ap[c+1]<ap[c]||ap[c+1]>nnz)throw Error{STORAGE};
            for(int32_t p=ap[c];p<ap[c+1];p++){
                int32_t r=ai[p];if(r<0||r>=n||(symmetric&&r>c))throw Error{STORAGE};double value=std::fabs(av[p]);
                if(phase==0){if(value>row_scale[r])row_scale[r]=value;if(symmetric&&value>row_scale[c])row_scale[c]=value;}
                else if(phase==2){value=std::fabs(rpx_fp_mul(av[p],row_scale[r]));if(value>col_scale[c])col_scale[c]=value;
                    if(symmetric){value=std::fabs(rpx_fp_mul(av[p],row_scale[c]));if(value>col_scale[r])col_scale[r]=value;}}
                else if(phase==4){if(seen[r]==c)throw Error{STORAGE};seen[r]=c;value=av[p];checks++;if(!finite(value))throw Error{MATRIX};
                    double mapped=rpx_fp_mul3(value,row_scale[r],col_scale[c]);input_max=std::max(input_max,std::fabs(mapped));put(r,inverse[c],mapped);
                    if(symmetric&&r!=c){mapped=rpx_fp_mul3(value,row_scale[c],col_scale[r]);input_max=std::max(input_max,std::fabs(mapped));put(c,inverse[r],mapped);}}
                else throw Error{STORAGE};
            }
        }
        if(phase==4&&last==n)factor_max=input_max;
    }
    void stream_lower(int32_t row,int32_t column,double value) {
        if(extra_used>=extra_capacity){int64_t cap=int64_t(extra_capacity)*2;if(cap<32768)cap=32768;if(cap>INT32_MAX-1)throw Error{FP_OVERFLOW};extra_capacity=int32_t(cap);memory(capacity,mask+1);extra_next.resize(size_t(cap)+1);extra_col.resize(size_t(cap)+1);extra_value.resize(size_t(cap)+1);}
        int32_t id=++extra_used;extra_next[id]=extra_head[row];extra_head[row]=id;extra_col[id]=column;extra_value[id]=value;
    }
    void stream_upper(int32_t column,double value) {
        if(u_used>=u_capacity){int64_t cap=int64_t(u_capacity)*2;if(cap<32768)cap=32768;if(cap>INT32_MAX)throw Error{FP_OVERFLOW};u_capacity=int32_t(cap);memory(capacity,mask+1);ui.resize(size_t(cap));uv.resize(size_t(cap));}
        ui[u_used]=column;uv[u_used]=value;u_used++;
    }
    bool eliminate(int32_t limit) {
        for(int32_t iter=0;iter<limit&&k<n;iter++,k++){
            double largest=0;int32_t best=-1,id=ch[k],h=0;
            while(id>0){int32_t r=nr[id];if(position[r]>=k){double value=std::fabs(nv[id]);if(value>largest||(value==largest&&(best<0||r<best))){largest=value;best=r;}}id=cn[id];}
            if(threshold>0&&largest>0){int32_t r=order[k];id=position[r]>=k?locate(r,k,h):0;if(id>0&&std::fabs(nv[id])>=rpx_fp_mul(threshold,largest))best=r;}
            if(best<0||largest==0){pivot_failed=true;pair.clear();return false;}
            int32_t old=row_at[k],p=position[best];row_at[k]=best;position[best]=k;row_at[p]=old;position[old]=p;
            id=locate(best,k,h);double pivot=nv[id];min_pivot=std::min(min_pivot,std::fabs(pivot));int32_t un=0;
            for(int32_t j=rh[best];j>0;j=rn[j])if(nc[j]>k){u_col[++un]=nc[j];u_val[un]=nv[j];}
            if(streaming){diagonal[k]=pivot;up[k]=u_used;for(int32_t j=1;j<=un;j++)stream_upper(u_col[j],u_val[j]);up[k+1]=u_used;}
            id=ch[k];while(id>0){int32_t next=cn[id],r=nr[id];if(position[r]>k){
                double mult=rpx_fp_div(nv[id],pivot);if(!std::isfinite(mult))throw Error{FP_OVERFLOW};checks++;if(!finite(mult))throw Error{FACTOR};
                if(streaming){stream_lower(r,k,mult);put(r,k,0);}else nv[id]=mult;max_l=std::max(max_l,std::fabs(mult));
                if(!pair_ready){if(stamp==INT32_MAX){std::fill(row_stamp.begin(),row_stamp.end(),0);stamp=0;}stamp++;scan_rows++;
                    for(int32_t scan=rh[r];scan>0;scan=rn[scan]){scan_nodes++;row_index[nc[scan]]=scan;row_stamp[nc[scan]]=stamp;}}
                for(int32_t at=1;at<=un;at++){int32_t c=u_col[at],target=pair_ready?pair[size_t(r)*n+c]:(row_stamp[c]==stamp?row_index[c]:0);
                    double value=target>0?nv[target]:0;value=rpx_fp_mulsub(value,mult,u_val[at]);if(!std::isfinite(value))throw Error{FP_OVERFLOW};factor_max=std::max(factor_max,std::fabs(value));
                    if(target>0&&value!=0){checks++;if(!finite(value))throw Error{FACTOR};nv[target]=value;updates++;}else put(r,c,value);
                }
            }id=next;}
            if(streaming){id=rh[best];while(id>0){int32_t next=rn[id];put(best,nc[id],0);id=next;}}
        }return true;
    }
    void sort_pairs(std::vector<int32_t>&ix,std::vector<double>&v,int32_t first,int32_t last) {
        if(first>=last)return;int32_t a=first,b=last,pivot=ix[(first+last)/2];
        while(a<=b){while(ix[a]<pivot)a++;while(ix[b]>pivot)b--;if(a<=b){std::swap(ix[a],ix[b]);std::swap(v[a],v[b]);a++;b--;}}
        if(first<b)sort_pairs(ix,v,first,b);if(a<last)sort_pairs(ix,v,a,last);
    }
    void compress() {
        if(k!=n||pivot_failed)throw Error{CONTEXT};lp.assign(size_t(n)+1,0);
        if(streaming){for(int32_t row=0;row<n;row++){int32_t total=lp[row];for(int32_t id=extra_head[row_at[row]];id>0;id=extra_next[id])total++;lp[row+1]=total;}
            memory(capacity,mask+1);li.resize(size_t(lp[n])+1);lv.resize(size_t(lp[n])+1);
            for(int32_t row=0;row<n;row++){int32_t at=lp[row];for(int32_t id=extra_head[row_at[row]];id>0;id=extra_next[id]){li[at]=extra_col[id];lv[at]=extra_value[id];at++;}
                sort_pairs(li,lv,lp[row],lp[row+1]-1);sort_pairs(ui,uv,up[row],up[row+1]-1);}
        }else{up.assign(size_t(n)+1,0);diagonal.resize(n);
            for(int32_t row=0;row<n;row++){int32_t ln=lp[row],un=up[row];for(int32_t id=rh[row_at[row]];id>0;id=rn[id]){int32_t c=nc[id];if(c<row)ln++;if(c>row)un++;if(c==row)diagonal[row]=nv[id];}lp[row+1]=ln;up[row+1]=un;}
            li.resize(size_t(lp[n])+1);lv.resize(size_t(lp[n])+1);ui.resize(size_t(up[n])+1);uv.resize(size_t(up[n])+1);
            for(int32_t row=0;row<n;row++){int32_t a=lp[row],b=up[row];for(int32_t id=rh[row_at[row]];id>0;id=rn[id]){int32_t c=nc[id];if(c<row){li[a]=c;lv[a++]=nv[id];}if(c>row){ui[b]=c;uv[b++]=nv[id];}}
                sort_pairs(li,lv,lp[row],lp[row+1]-1);sort_pairs(ui,uv,up[row],up[row+1]-1);}
        }
        ready=true;
        /* Release graph before VBA allocates its copies of compressed factors. */
        std::vector<int32_t>().swap(nr);std::vector<int32_t>().swap(nc);std::vector<double>().swap(nv);std::vector<int32_t>().swap(rn);std::vector<int32_t>().swap(rp);std::vector<int32_t>().swap(cn);std::vector<int32_t>().swap(cp);
        std::vector<int32_t>().swap(rh);std::vector<int32_t>().swap(ch);std::vector<int32_t>().swap(slots);std::vector<int32_t>().swap(pair);
        std::vector<int32_t>().swap(extra_head);std::vector<int32_t>().swap(extra_next);std::vector<int32_t>().swap(extra_col);std::vector<double>().swap(extra_value);
    }
};
std::unordered_map<uintptr_t,Lu*> live_handles;
uintptr_t next_handle=0;
Lu* get(uintptr_t h){auto i=live_handles.find(h);if(i==live_handles.end())throw Error{CONTEXT};return i->second;}
template<class F> int checked(F f){try{return f();}catch(const Error&e){return e.code;}catch(const std::bad_alloc&){return ALLOC;}catch(...){return UNKNOWN;}}
}
API rpx_lu_create(int32_t n,int32_t nnz,double maximum,double existing,int32_t epoch,const int32_t *order,int32_t streaming,double threshold,uintptr_t *handle) {
    if(!handle||!order)return BAD;*handle=0;return checked([&](){Lu*p=new Lu(n,nnz,maximum,existing,epoch,order,streaming!=0,threshold);try{if(next_handle==UINTPTR_MAX)throw Error{CONTEXT};uintptr_t id=++next_handle;live_handles.emplace(id,p);*handle=id;}catch(...){delete p;throw;}return OK;});
}
API rpx_lu_input(uintptr_t handle,int32_t phase,int32_t first,int32_t last,const int32_t *ap,const int32_t *ai,const double *av,int32_t symmetric,int32_t equilibrate) {
    return checked([&](){get(handle)->input(phase,first,last,ap,ai,av,symmetric!=0,equilibrate!=0);return OK;});
}
API rpx_lu_step(uintptr_t handle,int32_t limit,int32_t *written) {
    if(limit<1||!written)return BAD;return checked([&](){Lu*p=get(handle);bool ok=p->eliminate(limit);*written=p->k;return ok?OK:PIVOT;});
}
API rpx_lu_finish(uintptr_t handle,int32_t *lcount,int32_t *ucount) {
    if(!lcount||!ucount)return BAD;return checked([&](){Lu*p=get(handle);p->compress();*lcount=p->lp[p->n];*ucount=p->up[p->n];return OK;});
}
API rpx_lu_metrics(uintptr_t handle,double *out) {
    if(!out)return BAD;return checked([&](){Lu*p=get(handle);out[0]=p->bytes;out[1]=p->input_max>0?rpx_fp_div(p->factor_max,p->input_max):0;out[2]=p->min_pivot;out[3]=p->max_l;out[4]=p->updates;out[5]=p->capacity;out[6]=p->checks;out[7]=p->scan_nodes;out[8]=p->scan_rows;out[9]=p->pair_ready?1:0;return OK;});
}
API rpx_lu_export(uintptr_t handle,int32_t array,void *out,int32_t size) {
    if(!out||size<0)return BAD;return checked([&](){Lu*p=get(handle);if(!p->ready)throw Error{CONTEXT};const void *data=nullptr;size_t bytes=0;
        switch(array){case 0:data=p->order.data();bytes=p->order.size()*4;break;case 1:data=p->row_at.data();bytes=p->row_at.size()*4;break;case 2:data=p->row_scale.data();bytes=p->row_scale.size()*8;break;case 3:data=p->col_scale.data();bytes=p->col_scale.size()*8;break;
        case 4:data=p->lp.data();bytes=p->lp.size()*4;break;case 5:data=p->li.data();bytes=size_t(p->lp[p->n])*4;break;case 6:data=p->lv.data();bytes=size_t(p->lp[p->n])*8;break;case 7:data=p->up.data();bytes=p->up.size()*4;break;
        case 8:data=p->ui.data();bytes=size_t(p->up[p->n])*4;break;case 9:data=p->uv.data();bytes=size_t(p->up[p->n])*8;break;case 10:data=p->diagonal.data();bytes=p->diagonal.size()*8;break;default:return BAD;}
        if(bytes>size_t(size))return BAD;if(bytes)std::memcpy(out,data,bytes);return OK;});
}
API rpx_lu_destroy(uintptr_t handle) {
    auto it=live_handles.find(handle);if(it==live_handles.end())return CONTEXT;auto p=it->second;live_handles.erase(it);delete p;return OK;
}
API rpx_lu_backsolve(int32_t n,int32_t phase,int32_t first,int32_t last,const uintptr_t *args,int32_t lcount,int32_t ucount,int32_t *checks) {
    if(n<1||first<0||last<first||last>n||!args||!checks||lcount<0||ucount<0)return BAD;
    for(int32_t i=0;i<14;i++)if(!args[i])return BAD;
    auto order=reinterpret_cast<const int32_t*>(args[0]);auto row_at=reinterpret_cast<const int32_t*>(args[1]);auto rs=reinterpret_cast<const double*>(args[2]);auto cs=reinterpret_cast<const double*>(args[3]);
    auto lp=reinterpret_cast<const int32_t*>(args[4]);auto li=reinterpret_cast<const int32_t*>(args[5]);auto lv=reinterpret_cast<const double*>(args[6]);auto up=reinterpret_cast<const int32_t*>(args[7]);auto ui=reinterpret_cast<const int32_t*>(args[8]);auto uv=reinterpret_cast<const double*>(args[9]);auto d=reinterpret_cast<const double*>(args[10]);auto rhs=reinterpret_cast<const double*>(args[11]);auto work=reinterpret_cast<double*>(args[12]);auto out=reinterpret_cast<double*>(args[13]);
    if(phase==0){for(int32_t i=first;i<last;i++){if(row_at[i]<0||row_at[i]>=n||lp[i]<0||lp[i+1]<lp[i]||lp[i+1]>lcount)return BAD;double v=rpx_fp_mul(rhs[row_at[i]],rs[row_at[i]]);
        for(int32_t p=lp[i];p<lp[i+1];p++){if(li[p]<0||li[p]>=i)return BAD;v=rpx_fp_mulsub(v,lv[p],work[li[p]]);}if(!std::isfinite(v))return FP_OVERFLOW;work[i]=v;}return OK;}
    if(phase==1){for(int32_t i=last-1;i>=first;i--){if(up[i]<0||up[i+1]<up[i]||up[i+1]>ucount)return BAD;double v=work[i];for(int32_t p=up[i];p<up[i+1];p++){if(ui[p]<=i||ui[p]>=n)return BAD;v=rpx_fp_mulsub(v,uv[p],work[ui[p]]);}if(d[i]==0)return FP_DIVZERO;work[i]=rpx_fp_div(v,d[i]);if(!std::isfinite(work[i]))return FP_OVERFLOW;(*checks)++;if(!finite(work[i]))return SOLUTION;}return OK;}
    if(phase==2){for(int32_t i=first;i<last;i++){if(order[i]<0||order[i]>=n)return BAD;out[order[i]]=rpx_fp_mul(cs[order[i]],work[i]);if(!std::isfinite(out[order[i]]))return FP_OVERFLOW;}return OK;}return BAD;
}
API rpx_lu_live_handles(void) { return int(live_handles.size()); }



/* DLL4: factors and forward/backsolve workspace remain owned by this context.
 * Borrowed VBA pointers exist only during this call. Epoch and dimension must
 * agree; every release/park/mode switch invalidates the scoped handle. */
API rpx_lu_retained_solve(uintptr_t handle,int32_t epoch,int32_t n,int32_t phase,int32_t first,int32_t last,const double *rhs,double *out,int32_t *checks) {
    if(!rhs||!out||!checks)return BAD;
    return checked([&](){Lu*p=get(handle);if(!p->ready||p->epoch!=epoch||p->n!=n)throw Error{CONTEXT};
        int32_t zero_index=0;double zero_value=0;
        uintptr_t args[14]={reinterpret_cast<uintptr_t>(p->order.data()),reinterpret_cast<uintptr_t>(p->row_at.data()),
            reinterpret_cast<uintptr_t>(p->row_scale.data()),reinterpret_cast<uintptr_t>(p->col_scale.data()),
            reinterpret_cast<uintptr_t>(p->lp.data()),reinterpret_cast<uintptr_t>(p->li.data()),reinterpret_cast<uintptr_t>(p->lv.data()),
            reinterpret_cast<uintptr_t>(p->up.data()),reinterpret_cast<uintptr_t>(p->ui.empty()?&zero_index:p->ui.data()),reinterpret_cast<uintptr_t>(p->uv.empty()?&zero_value:p->uv.data()),
            reinterpret_cast<uintptr_t>(p->diagonal.data()),reinterpret_cast<uintptr_t>(rhs),reinterpret_cast<uintptr_t>(p->u_val.data()),reinterpret_cast<uintptr_t>(out)};
        return rpx_lu_backsolve(n,phase,first,last,args,p->lp[n],p->up[n],checks);
    });
}
