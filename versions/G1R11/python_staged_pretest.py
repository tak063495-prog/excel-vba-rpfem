from python_mesh_experiments import *
from common import load_mesh
from model_support.p2 import gradients
base=saved_mesh()
# Reconstruct the q4 upper field used before the final tightening in the legacy run.
row=probe(base,'legacy781_q4',.826330730099349,'upper',4)
assert row['accepted_witness'],row
old=load_mesh('P06N',path=R/'results/legacy781_q4_upper_0.826330730099349_4.npz')
with np.load(R/'results/legacy781_q4_upper_0.826330730099349_4.npz') as z:velocity=z['x'].reshape(-1,6,2)
area,grad,edges=topology(base);score=np.zeros(len(base.tri))
for q in range(6):
    a,w=(.445948490915965,.223381589678011) if q<3 else (.091576213509771,.109951743655322)
    lam=np.full(3,a);lam[q%3]=1-2*a
    for e in range(len(base.tri)):
        g=gradients(lam,grad[e]);ex=velocity[e,:,0]@g[:,0];ey=velocity[e,:,1]@g[:,1];sh=velocity[e,:,0]@g[:,1]+velocity[e,:,1]@g[:,0]
        score[e]+=area[e]*w*np.sqrt(ex*ex+ey*ey+.5*sh*sh)
for edge in edges:
    e,f=edge.elements
    if f>=0:
        for s,w in [(0.5-np.sqrt(.15),5/18),(.5,4/9),(.5+np.sqrt(.15),5/18)]:
            weights=np.array([(1-s)*(1-2*s),s*(2*s-1),4*s*(1-s)])
            value=edge.length*w*np.linalg.norm(weights@(velocity[e,edge.loc[0]]-velocity[f,edge.loc[1]]))/2
            score[e]+=value;score[f]+=value
raw=score.copy()
for edge in edges:
    e,f=edge.elements
    if f>=0 and base.material_id[e]!=base.material_id[f]:score[e]+=.25*raw[f];score[f]+=.25*raw[e]
incidence={}
for t in base.tri:
    for a,b in zip(t,np.roll(t,-1)):
        k=tuple(sorted((int(a),int(b))));incidence[k]=incidence.get(k,0)+1
order=np.argsort(-score,kind='stable')
variants={};records=[row]
for cap in [994,1000,1050]:
    marks=set();selected=[];ne=len(base.tri)
    for e in order[:round(.3*len(base.tri))]:
        t=base.tri[e];keys={tuple(sorted((int(a),int(b)))) for a,b in zip(t,np.roll(t,-1))}
        extra=sum(incidence[k] for k in keys-marks)
        if ne+extra>cap:break
        marks|=keys;selected.append(int(e));ne+=extra
    mesh=refine(base,set(selected),True);assert len(mesh.tri)==ne and ne<=cap
    label=f'staged_{cap}';save_mesh(R/'results'/(label+'_mesh.npz'),mesh)
    np.savez(R/'results'/(label+'_selection.npz'),score=score,velocity=velocity,selected=np.array(selected))
    print('MESH',label,'elements',len(mesh.tri),'selected',len(selected),flush=True)
    for fs,mode,q in [(.655,'lower',4),(.66,'lower',4),(.77,'upper',8)]:
        records.append(probe(mesh,label,fs,mode,q))
        (R/'results/staged_pretest.json').write_text(json.dumps(records,ensure_ascii=False,indent=2),encoding='utf-8')
assert any(r['accepted_witness'] for r in records if r['mesh']=='staged_1000' and r['mode']=='lower')
assert any(r['accepted_witness'] for r in records if r['mesh']=='staged_1000' and r['mode']=='upper')
print('Python-first staged policy candidate PASS; audited witnesses only, not full Fs roots',flush=True)
