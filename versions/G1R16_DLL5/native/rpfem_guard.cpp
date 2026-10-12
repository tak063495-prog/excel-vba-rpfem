/* DLL5 input identity. Borrow SAFEARRAYs only during the call; retain nothing.
 * Preserve VBA type/value/formula semantics, including signed-zero equality.
 * Unsupported values return 3 so the original VBA predicate handles them.
 * SPDX-License-Identifier: LGPL-2.1-or-later */
#define NOMINMAX
#include <windows.h>
#include <oleauto.h>
#include <cmath>
#include <cstring>
#include <cstdint>
#define API extern "C" __declspec(dllexport) int __stdcall
namespace {
struct View {
 SAFEARRAY* a=nullptr;void* data=nullptr;LONG rows=0,cols=0,lo=0;
 ~View(){if(data)SafeArrayUnaccessData(a);}
 bool open(const VARIANT* v,VARTYPE type,int dimensions){
  if(!v||v->vt!=(VT_ARRAY|type)||!v->parray||SafeArrayGetDim(v->parray)!=UINT(dimensions))return false;
  a=v->parray;LONG lower=0,upper=0;
  if(FAILED(SafeArrayGetLBound(a,1,&lower))||FAILED(SafeArrayGetUBound(a,1,&upper))||upper<lower)return false;
  lo=lower;rows=upper-lower+1;
  if(dimensions==2){LONG lc=0,uc=0;if(FAILED(SafeArrayGetLBound(a,2,&lc))||FAILED(SafeArrayGetUBound(a,2,&uc))||lc!=1||uc<lc||lower!=1)return false;cols=uc-lc+1;}
  return SUCCEEDED(SafeArrayAccessData(a,&data));
 }
};
int cell(const VARIANT& a,const VARIANT& b){
 if(a.vt!=b.vt)return 1;
 switch(a.vt){
 case VT_EMPTY:return 0;
 case VT_R8:case VT_DATE:if(!std::isfinite(a.dblVal)||!std::isfinite(b.dblVal))return 3;return a.dblVal!=b.dblVal;
 case VT_R4:if(!std::isfinite(a.fltVal)||!std::isfinite(b.fltVal))return 3;return a.fltVal!=b.fltVal;
 case VT_I1:return a.cVal!=b.cVal;
 case VT_UI1:return a.bVal!=b.bVal;
 case VT_I2:return a.iVal!=b.iVal;
 case VT_I4:return a.lVal!=b.lVal;
 case VT_BOOL:if((a.boolVal!=0&&a.boolVal!=-1)||(b.boolVal!=0&&b.boolVal!=-1))return 3;return a.boolVal!=b.boolVal;
 case VT_CY:return a.cyVal.int64!=b.cyVal.int64;
 case VT_ERROR:return a.scode!=b.scode;
 case VT_BSTR:{UINT n=SysStringLen(a.bstrVal);if(n!=SysStringLen(b.bstrVal))return 1;return n&&std::memcmp(a.bstrVal,b.bstrVal,size_t(n)*sizeof(wchar_t))!=0;}
 default:return 3;
 }
}
}
API rpx_snapshot_compare(const VARIANT*live,const VARIANT*forms,const VARIANT*saved,const VARIANT*oldforms,const VARIANT*columns,int32_t n,int32_t rows,int32_t*changed){
 if(!changed)return 330;*changed=0;View a,b,c,d,cols;
 if(n<1||rows<1||!a.open(live,VT_VARIANT,2)||!b.open(forms,VT_VARIANT,2)||!c.open(saved,VT_VARIANT,2)||!d.open(oldforms,VT_VARIANT,2)||!cols.open(columns,VT_I4,1))return 330;
 if(a.rows!=rows||b.rows!=rows||c.rows!=rows||d.rows!=rows||a.cols!=b.cols||a.cols!=c.cols||a.cols!=d.cols||cols.rows<n)return 330;
 auto ids=static_cast<const LONG*>(cols.data);auto av=static_cast<const VARIANT*>(a.data),bv=static_cast<const VARIANT*>(b.data),cv=static_cast<const VARIANT*>(c.data),dv=static_cast<const VARIANT*>(d.data);
 for(int k=0;k<n;k++)if(ids[k]<1||ids[k]>a.cols)return 330;
 for(LONG r=0;r<rows;r++)for(int k=0;k<n;k++){
  size_t at=size_t(ids[k]-1)*size_t(rows)+size_t(r);int x=cell(av[at],cv[at]),y=cell(bv[at],dv[at]);
  if(x==3||y==3)return 3;if(x||y){*changed=1;return 0;}
 }
 return 0;
}
