"""Audit the native field, then solve the identical exported mesh independently."""
import json,struct,numpy as np
from support import ROOT,REPO,write
from diagnose import mesh_from_arrays,solve
from common import physical_audit

def read(prefix,rec):
 raw=prefix.with_name(prefix.name+'_mesh.bin').read_bytes();nn,ne=struct.unpack_from('<ii',raw);xy=np.frombuffer(raw,'<f8',nn*2,8).copy().reshape(nn,2);t=np.frombuffer(raw,'<i4',ne*4,8+16*nn).copy().reshape(ne,4)
 m=mesh_from_arrays(xy,t[:,:3],t[:,3],rec)
 raw=prefix.with_name(prefix.name+'_field.bin').read_bytes();assert raw[:8]==b'P06X0001';n=struct.unpack_from('<i',raw,8)[0];x=np.frombuffer(raw,'<f8',n,12).copy()
 return m,x

def main():
 folder=ROOT/'results/fixed_native';meta=json.loads((folder/'summary.json').read_text(encoding='utf8'));rec=json.loads((ROOT/'data/inputs/P06N.json').read_text(encoding='utf8'));out=[]
 for row in meta['rows']:
  side=row['side'];fs=row['Fs'];prefix=folder/(side+'_'+str(fs));m,x=read(prefix,rec);audit=physical_audit(m,x[:12*len(m.tri)] if side=='upper' else x,side,fs,8);assert audit['pass'],audit
  native=audit['value'];entry={'side':side,'Fs':fs,'mesh_id':m.mesh_id,'native_value':native,'native_field_independent_audit':audit,'python_trials':[]};out.append(entry)
  for reduced in [False,True]:
   r,y,pr,z=solve(m,fs,side,reduce=reduced);entry['python_trials'].append(r)
   print(side,fs,'reduce',reduced,'native',native,'python',r['objective'] if side=='upper' else r['load'],'gap',r['dual_gap'],'audit',r.get('physical_audit',{}).get('pass'),flush=True)
  write(ROOT/'results/fixed_python_comparison.json',{'points':out,'scope':'Actual native exported geometry; no reconstructed/rounded mesh key used as identity proof'})
if __name__=='__main__':main()
