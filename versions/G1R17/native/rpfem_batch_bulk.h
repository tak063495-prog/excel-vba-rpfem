/* DLL5: time-budgeted complete bulk loops. The original kernels, chunk order,
 * stored FP boundaries and rejection semantics remain unchanged. */
#define BP(T,I) ((T*)a[I])
API rpx_bulk_step(int32_t op,const int32_t*d,const uintptr_t*a,int32_t total,int32_t*cursor,double alpha,double*m,int32_t ms){
 int32_t first,last,chunk;int code=0;LARGE_INTEGER q,f;LONGLONG deadline;
 if(!d||!a||!cursor||!m||total<1||*cursor<0||*cursor>total||ms<0||ms>1000)return BAD_ARGUMENT;
 chunk=256;if(op==1)chunk=64;if(op==4||op==5||op==6)chunk=128;if(op==8)chunk=512;if(op==9)chunk=65536;
 QueryPerformanceFrequency(&f);QueryPerformanceCounter(&q);deadline=q.QuadPart+f.QuadPart*ms/1000;
 do{first=*cursor;last=first+chunk;if(last>total)last=total;
  switch(op){
   case 1:code=rpx_scaling(d[0],d[1],first,last,d[5],BP(int32_t,0),BP(int32_t,1),BP(double,2),BP(double,3),BP(double,4),BP(double,5),BP(double,6),BP(double,7),BP(double,8),BP(int16_t,9),BP(dd,10),BP(int32_t,11));break;
   case 2:code=rpx_gproduct(d[0],d[1],d[2],first,last,BP(int32_t,0),BP(int32_t,1),BP(int32_t,2),BP(int32_t,3),BP(int32_t,4),BP(double,5),BP(double,6),BP(double,7),BP(double,8),d[5]);break;
   case 3:code=rpx_gtadd(d[0],d[1],d[2],first,last,BP(int32_t,0),BP(int32_t,1),BP(int32_t,2),BP(int32_t,3),BP(int32_t,4),BP(double,5),BP(double,6),BP(double,7),alpha);break;
   case 4:code=rpx_kkt_cones(d[0],d[6],first,last,BP(int32_t,0),BP(int32_t,1),BP(int32_t,2),BP(int32_t,3),BP(int32_t,4),BP(double,5),BP(double,6),BP(double,7));break;
   case 5:code=rpx_corrector(d[0],d[1],d[4],first,last,BP(int32_t,0),BP(int32_t,1),alpha,BP(uintptr_t,2));break;
   case 6:code=rpx_hoperator(d[2]+d[3],d[2],d[0],d[1],d[4],first,last,d[5],BP(uintptr_t,0));break;
   case 7:code=rpx_solver_cones(d[4],d[0],d[1],d[2],first,last,BP(int32_t,0),BP(int32_t,1),BP(int32_t,2),BP(int32_t,3),BP(int32_t,4),BP(double,5),BP(uintptr_t,6));break;
   case 8:code=rpx_solver_residual(d[0],d[1],d[2],d[3],d[4],first,last,BP(int32_t,0),BP(int32_t,1),BP(int32_t,2),BP(int32_t,3),BP(int32_t,4),BP(double,5),BP(double,6),BP(uintptr_t,7),alpha,m);break;
   case 9:code=rpx_kkt_init(d[6],first,last,BP(double,0),BP(double,1));break;
   case 10:code=rpx_kkt_diagonal(d[2]+d[3],d[2],d[6],first,last,BP(int32_t,0),BP(int32_t,1),alpha,BP(double,2),BP(double,3));break;
   case 11:code=rpx_kkt_scale(d[2]+d[3],d[6],first,last,BP(int32_t,0),BP(int32_t,1),BP(double,2),BP(double,3),BP(double,4));break;
   default:return BAD_ARGUMENT;
  }
  if(code)return code;*cursor=last;if(ms==0||last==total)return 0;QueryPerformanceCounter(&q);
 }while(q.QuadPart<deadline);return 0;
}
#undef BP
