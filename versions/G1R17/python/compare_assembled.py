"""Actual VBA assembled lower SOCP vs independent Python, same binary mesh."""
import json,struct,numpy as np
from scipy import sparse
from support import ROOT,write
from diagnose import mesh_from_arrays,assemble_lower
def read_program(path):
 raw=path.read_bytes();assert raw[:8]==b'P06VBA01';nx,nm,nb=struct.unpack_from('<iii',raw,8);pos=20
 q=np.frombuffer(raw,'<f8',nx,pos).copy();pos+=8*nx
 def row():
  nonlocal pos
  n=struct.unpack_from('<i',raw,pos)[0];v=struct.unpack_from('<d',raw,pos+4)[0];pos+=12
  d=np.frombuffer(raw,np.dtype([('id','<i4'),('v','<f8')]),n,pos);pos+=12*n
  return v,[(int(a['id']),float(a['v']))for a in d]
 def matrix(rows):
  ii=[];jj=[];vv=[]
  for i,r in enumerate(rows):
   for j,v in r:ii.append(i);jj.append(j);vv.append(v)
  return sparse.coo_matrix((vv,(ii,jj)),shape=(len(rows),nx)).tocsc()
 er=[];b=[]
 for _ in range(nm):v,r=row();b.append(v);er.append(r)
 cr=[];h=[];dims=[]
 for _ in range(nb):
  dim=struct.unpack_from('<i',raw,pos)[0];pos+=4;dims.append(dim)
  for _ in range(dim):v,r=row();h.append(v);cr.append(r)
 assert pos==len(raw)
 E=matrix(er);norm=np.sqrt(np.asarray(E.power(2).sum(axis=1)).ravel());assert np.all(norm>0)
 return sparse.diags(1/norm)@E,np.asarray(b)/norm,matrix(cr),np.asarray(h),dims,q
def main():
 rows=[];folder=ROOT/'results/geometry25'
 for case in ['P06N','P06A']:
  rec=json.loads((ROOT/'data/inputs'/(case+'.json')).read_text(encoding='utf8'))
  for version in ['old','new']:
   prefix=folder/(case+'_'+version+'_model');raw=prefix.with_name(prefix.name+'_mesh.bin').read_bytes();nn,ne=struct.unpack_from('<ii',raw)
   xy=np.frombuffer(raw,'<f8',nn*2,8).copy().reshape(nn,2);tri=np.frombuffer(raw,'<i4',ne*4,8+16*nn).copy().reshape(ne,4)
   m=mesh_from_arrays(xy,tri[:,:3],tri[:,3],rec);E,b,G,h,dims,q=read_program(prefix.with_name(prefix.name+'_assembled.bin'))
   E1,b1,G1,h1,dims1,zero=assemble_lower(m,.5)
   assert E.shape==E1.shape and G.shape==G1.shape and dims==dims1,(case,version,E.shape,E1.shape,G.shape,G1.shape)
   errors={'E':float(np.max(np.abs((E-E1).data),initial=0)),'b':float(np.max(abs(b-b1),initial=0)),'G':float(np.max(np.abs((G-G1).data),initial=0)),'h':float(np.max(abs(h-h1),initial=0))}
   assert max(errors.values())<2e-13,(case,version,errors)
   assert q[-1]==-1 and np.count_nonzero(q)==1
   rows.append({'case':case,'version':version,'nodes':nn,'elements':ne,'Fs':.5,'variables':E.shape[1],'equations':E.shape[0],'cones':len(dims),'max_coefficient_errors':errors,'PASS':True})
   print(case,version,'ASSEMBLY PASS',errors,flush=True)
 write(ROOT/'results/assembled_python_comparison.json',{'rows':rows,'PASS':True,'scope':'Four native VBA exports. Exact binary geometry; same MC/Davis/Fs. Compare normalized equations because Python normalizes them.'})
if __name__=='__main__':main()
