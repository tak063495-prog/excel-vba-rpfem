/* Exact port of G1R15 RPX_Order / RPX_ColumnOrder. No floating arithmetic.
 * SPDX-License-Identifier: Apache-2.0
 * Stable key insertion order, snapshot enumeration, heap tie rules, and the
 * n > 20000 degree branch are deliberately identical to the VBA source.
 * The context owns only its graph. No VBA pointer survives an API call.
 */
#include <vector>
#include <unordered_map>
#include <memory>
#include <stdexcept>
#include <new>
#include <cstdint>
#include <limits>
#include <algorithm>
#define API extern "C" __declspec(dllexport) int __stdcall
enum { OK=0, BAD=100, ORDER_OVERFLOW=206, OOM=207, UNKNOWN=208 };
static int checked(int64_t value) {
    if(value<INT32_MIN || value>INT32_MAX) throw std::overflow_error("Long overflow");
    return static_cast<int>(value);
}
struct OrderedSet {
    struct Entry { int key; bool live; };
    std::vector<Entry> entries;
    std::unordered_map<int,size_t> index;
    void add(int k) {
        if(index.find(k)!=index.end()) return;
        index.emplace(k,entries.size()); entries.push_back({k,true});
    }
    bool remove(int k) {
        auto it=index.find(k); if(it==index.end()) return false;
        entries[it->second].live=false; index.erase(it);
        // Compaction preserves the exact insertion order of the live keys.
        if(entries.size()>2*index.size()+256) {
            size_t pos=0;
            for(const auto& e:entries) if(e.live) {
                entries[pos]=e; index.find(e.key)->second=pos; ++pos;
            }
            entries.resize(pos);
        }
        return true;
    }
    int count() const { return checked(static_cast<int64_t>(index.size())); }
    std::vector<int> keys() const {
        std::vector<int> out; out.reserve(index.size());
        for(const auto& e:entries) if(e.live) out.push_back(e.key);
        return out;
    }
    void clear() { entries.clear(); index.clear(); }
};
struct OrderContext {
    int n, mode, nextRow=0, heapSize, written=0, tag=0, neighborEpoch=0, cliqueEpoch=0;
    bool started=false;
    std::vector<OrderedSet> adjacency,incident;
    std::vector<std::unique_ptr<OrderedSet>> clique;
    std::vector<int> selected,externalDegree,neighborMark,cliqueMark,heap,heapAt,priority;
    OrderContext(int size,int kind):n(size),mode(kind),heapSize(size),
        adjacency(size),incident(size),clique(kind?checked(2LL*size):size),
        selected(size),externalDegree(clique.size()),neighborMark(size),
        cliqueMark(clique.size()),heap(size),heapAt(size),priority(size) {}
    void heapSwap(int a,int b) {
        std::swap(heap[a],heap[b]); heapAt[heap[a]]=a; heapAt[heap[b]]=b;
    }
    void heapUpdate(int v) {
        int p=heapAt[v];
        while(p>0) {
            int par=(p-1)/2;
            if(priority[heap[par]]<=priority[v]) break;
            heapSwap(p,par); p=par;
        }
        for(;;) {
            int64_t child=2LL*p+1; if(child>=heapSize) break;
            int j=static_cast<int>(child);
            if(j+1<heapSize && priority[heap[j+1]]<priority[heap[j]]) ++j;
            if(priority[v]<=priority[heap[j]]) break;
            heapSwap(p,j); p=j;
        }
    }
    const OrderedSet& getClique(int el) const {
        if(el<0 || static_cast<size_t>(el)>=clique.size() || !clique[el]) throw std::logic_error("missing clique");
        return *clique[el];
    }
    void feed(int first,int last,const int* ap,int apSize,const int* ai,int aiSize) {
        if(started || first!=nextRow || first<0 || last<first || last>n || !ap || apSize!=n+1 || aiSize<0 || (aiSize && !ai)) throw std::invalid_argument("feed range");
        for(int i=first;i<last;++i) {
            if(ap[i]<0 || ap[i+1]<ap[i] || ap[i+1]>aiSize) throw std::invalid_argument("row extent");
            if(mode) clique[n+i]=std::make_unique<OrderedSet>();
            for(int p=ap[i];p<ap[i+1];++p) {
                int j=ai[p]; if(j<0 || j>=n) throw std::invalid_argument("column index");
                if(mode) { clique[n+i]->add(j); incident[j].add(n+i); }
                else if(j!=i) { adjacency[i].add(j); adjacency[j].add(i); }
            }
        }
        nextRow=last;
    }
    void start() {
        if(started || nextRow!=n) throw std::invalid_argument("initialization incomplete");
        for(int i=0;i<n;++i) { heap[i]=i; heapAt[i]=i; priority[i]=adjacency[i].count(); }
        if(mode) for(int i=0;i<n;++i) {
            tag=checked(static_cast<int64_t>(tag)+1); selected[i]=tag; int degree=0;
            for(int el:incident[i].keys()) for(int k:getClique(el).keys())
                if(selected[k]!=tag) { selected[k]=tag; degree=checked(static_cast<int64_t>(degree)+1); }
            priority[i]=degree;
        }
        for(int i=n-1;i>=0;--i) heapUpdate(heap[i]);
        started=true;
    }
    void eliminate(int* output) {
        int v=heap[0]; auto neighbors=std::make_unique<OrderedSet>();
        for(int j:adjacency[v].keys()) neighbors->add(j);
        const auto inc=incident[v].keys(); // VBA .keys is a snapshot, even during Remove.
        for(int el:inc) for(int j:getClique(el).keys()) neighbors->add(j);
        neighbors->remove(v);
        for(int el:inc) {
            for(int j:getClique(el).keys()) if(!incident[j].remove(el)) throw std::logic_error("incident mismatch");
            clique[el].reset();
        }
        --heapSize;
        if(heapSize>0) { heap[0]=heap[heapSize]; heapAt[heap[0]]=0; heapUpdate(heap[0]); }
        output[written++]=v;
        const auto nbr=neighbors->keys(); const int nbrN=checked(static_cast<int64_t>(nbr.size()));
        neighborEpoch=checked(static_cast<int64_t>(neighborEpoch)+1);
        for(int j:nbr) neighborMark[j]=neighborEpoch;
        for(int j:nbr) {
            adjacency[j].remove(v);
            for(int k:adjacency[j].keys()) if(neighborMark[k]==neighborEpoch) adjacency[j].remove(k);
            incident[j].add(v);
        }
        adjacency[v].clear(); incident[v].clear(); clique[v]=std::move(neighbors);
        if(n>20000) {
            for(int j:nbr) {
                tag=checked(static_cast<int64_t>(tag)+1); selected[j]=tag; int degree=0;
                for(int k:adjacency[j].keys()) if(selected[k]!=tag) { selected[k]=tag; degree=checked(static_cast<int64_t>(degree)+1); }
                for(int el:incident[j].keys()) for(int k:getClique(el).keys())
                    if(selected[k]!=tag) { selected[k]=tag; degree=checked(static_cast<int64_t>(degree)+1); }
                priority[j]=degree; heapUpdate(j);
            }
        } else {
            cliqueEpoch=checked(static_cast<int64_t>(cliqueEpoch)+1);
            for(int j:nbr) for(int el:incident[j].keys()) if(cliqueMark[el]!=cliqueEpoch) {
                cliqueMark[el]=cliqueEpoch; int degree=0;
                for(int k:getClique(el).keys()) if(neighborMark[k]!=neighborEpoch) degree=checked(static_cast<int64_t>(degree)+1);
                externalDegree[el]=degree;
            }
            for(int j:nbr) {
                int degree=checked(static_cast<int64_t>(adjacency[j].count())+nbrN-1);
                for(int el:incident[j].keys()) degree=checked(static_cast<int64_t>(degree)+externalDegree[el]);
                if(degree>heapSize-1) degree=heapSize-1;
                priority[j]=degree; heapUpdate(j);
            }
        }
    }
};
#define CATCH_ERRORS catch(const std::bad_alloc&) { return OOM; } \
    catch(const std::overflow_error&) { return ORDER_OVERFLOW; } \
    catch(const std::length_error&) { return ORDER_OVERFLOW; } \
    catch(const std::invalid_argument&) { return BAD; } \
    catch(...) { return UNKNOWN; }
API rpx_order_create(int n,int mode,void** handle) {
    if(!handle) return BAD; *handle=nullptr;
    if(n<=0 || n==INT32_MAX || (mode!=0 && mode!=1)) return BAD;
    try { *handle=new OrderContext(n,mode); return OK; } CATCH_ERRORS
}
API rpx_order_feed(void* handle,int first,int last,const int* ap,int apSize,const int* ai,int aiSize) {
    if(!handle) return BAD;
    try { static_cast<OrderContext*>(handle)->feed(first,last,ap,apSize,ai,aiSize); return OK; } CATCH_ERRORS
}
API rpx_order_start(void* handle) {
    if(!handle) return BAD;
    try { static_cast<OrderContext*>(handle)->start(); return OK; } CATCH_ERRORS
}
API rpx_order_step(void* handle,int limit,int* written,int* output,int outputSize) {
    if(!handle || !written || !output || limit<=0) return BAD;
    auto& c=*static_cast<OrderContext*>(handle);
    if(!c.started || outputSize!=c.n) return BAD;
    try {
        int last=static_cast<int>(std::min(static_cast<int64_t>(c.n),static_cast<int64_t>(c.written)+limit));
        while(c.written<last) c.eliminate(output);
        *written=c.written; return OK;
    } CATCH_ERRORS
}
API rpx_order_destroy(void* handle) { delete static_cast<OrderContext*>(handle); return OK; }
