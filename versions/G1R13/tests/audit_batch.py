from pathlib import Path
import json,struct,sys,math
import numpy as np
from prepare_inputs import REPO
sys.path.insert(0,str(REPO/'versions/G1R12/validation_runtime/python'))
from common import Mesh,physical_audit
from independent_audit import audit
from zero_phi import audit_zero
from audit_precision import audit_precision
def read_state(path,policy,phase):
    raw=path.read_bytes();assert raw[:8]==b'RP25M001';pos=8
    def take(dtype,count):
        nonlocal pos
        out=np.frombuffer(raw,dtype=dtype,count=count,offset=pos).copy();pos+=out.nbytes;return out
    nn,ne,nm,nr,q=map(int,take('<i4',5));assert nr==0
    scale,factor,fs,objective,native_audit=take('<f8',5)
    xy=take('<f8',nn*2).reshape(nn,2);tr=take('<i4',ne*5).reshape(ne,5)
    mats=take('<f8',nm*4).reshape(nm,4);body=take('<f8',ne*4).reshape(ne,2,2)
    nb=int(take('<i4',1)[0]);bd={};kinds=['load','fixed','roller_x','roller_y']
    for _ in range(nb):
        a,b,k,r=map(int,take('<i4',4));loads=take('<f8',4).reshape(2,2).T
        bd[tuple(sorted((a,b))) ]={'kind':kinds[k],'rigid':r,'loads':loads}
    nf=int(take('<i4',1)[0]);x=take('<f8',nf);assert pos==len(raw)
    assert phase=='reference_total' and factor==fs and np.all(tr[:,4]<0)
    m=Mesh(xy,tr[:,:3],mats,tr[:,3],tr[:,4],body,bd,float(scale),'native_batch25',load_phase=phase,provenance={'psi_policy':policy}).validate()
    return m,x,{'nodes':nn,'elements':ne,'q':q,'factor':factor,'fs':fs,'objective':objective,'native_audit':native_audit,'scale':scale}
def audit_case(case,folder,rec,native):
    rows=[];c=rec['source_case']
    for i,side in enumerate(['lower','upper']):
        m,x,meta=read_state(folder/f'witness_{side}.bin',c['policy'],native['witness_phase'][i])
        assert meta['q']==native['witness_q'][i]
        assert abs(meta['fs']-native[side+'_fs'])<=1e-12*meta['fs']
        mats=np.array([[a['c'],a['phi'],a['psi'],a['gamma']] for a in c['materials']]);assert np.array_equal(m.materials,mats)
        p=m.xy[m.tri];det=(p[:,1,0]-p[:,0,0])*(p[:,2,1]-p[:,0,1])-(p[:,2,0]-p[:,0,0])*(p[:,1,1]-p[:,0,1]);areas=det/2
        actual=[float(areas[m.material_id==j].sum()) for j in range(len(mats))]
        assert np.allclose(actual,rec['region_areas'],rtol=1e-9,atol=1e-8),(case,actual,rec['region_areas'])
        assert np.all(m.body[:,:,0]==0),case
        assert np.array_equal(m.body[:,1,1],-m.materials[m.material_id,3]) and np.all(m.body[:,0,1]==0),case
        weight=float(-areas@m.body[:,1,1]);assert math.isclose(weight,rec['total_weight'],rel_tol=1e-9)
        # Verify the solved native boundary/load against the explicitly prepared input.
        tol=1e-10*max(np.ptp(m.xy,axis=0).max(),1.)
        def on_segment(p,a,b):
            a=np.array(a);b=np.array(b);d=b-a;v=p-a
            t=float(v@d/(d@d))
            return -1e-10<=t<=1+1e-10 and abs(d[0]*v[1]-d[1]*v[0])<=tol*np.linalg.norm(d)
        expected_kinds={'自由':'load','固定XY':'fixed','固定X':'roller_x','固定Y':'roller_y'}
        for key,edge in m.boundary.items():
            a,b=m.xy[list(key)]
            matches=[f for f in rec['boundary_explicit'] if on_segment(a,f['a'],f['b']) and on_segment(b,f['a'],f['b'])]
            assert len(matches)==1,(case,key,a,b,matches)
            wanted=matches[0]
            assert edge['kind']==expected_kinds[wanted['kind']],(case,key,edge['kind'],wanted)
            assert np.array_equal(edge['loads'][:,0],np.zeros(2))
            assert np.array_equal(edge['loads'][:,1],np.array(wanted['load0'])+np.array(wanted['load1'])),(case,key,edge,wanted)
        if side=='upper' and np.all(m.materials[:,1]==0):
            au=audit_zero(m,x,meta['fs'],meta['q'],meta['objective'])
        else:
            au=audit(m,x,side,meta['fs'],meta['q'],meta['objective'],18*len(m.tri))
        precision=None
        if side=='upper':
            m.provenance['case']=case
            precision=audit_precision(m,x,meta['fs'],'upper')
            # Record ordinary Double diagnostics; final decision uses independent 50-digit integration.
            if np.all(m.materials[:,0]==0):
                au['ordinary_double_pass']=au['pass']
                au['pass']=bool(precision['accepted'])
                au['acceptance_basis']='50-digit independent physical audit; unchanged 1e-7'
        raw={k:max(0.,float(v))*m.stress_scale for k,v in au['errors'].items()} if side=='lower' else {}
        raw_pass=max(raw.values(),default=0.)<=1e-7
        directional = float(x[18*len(m.tri)])>=1.-1e-8 if side=='lower' else au['value']<=1.+1e-8
        rows.append({'side':side,'metadata':meta,'physical_audit':au,'precision_audit':precision,'raw_lower_errors':raw,'raw_lower_pass':bool(raw_pass),'endpoint_direction_pass':bool(directional),'material_areas':actual,'weight':weight,'pass':bool(au['pass'] and raw_pass and directional)})
    return {'pass':all(r['pass'] for r in rows),'scope':'independent raw-equation/geometry/weight/material audits of native witnesses','results':rows}
