from pathlib import Path
import json,re,math,hashlib,collections,csv,shutil,os
ROOT=Path(__file__).resolve().parent
REPO=Path(os.environ.get('RPFEM_REPO','C:/Users/link_/Downloads/excel-vba-rpfem_G1R11_20261008'))
CATALOG=Path(os.environ.get('RPFEM_CATALOG','C:/Users/link_/Downloads/Test_Case_Catalog.md'))
SOURCE=REPO/'versions/G1R12/RPFEM_20261009_G1R12_P06N_IssueFix.xlsm'
SOURCE_SHA='fe02ff61687cff8e0f0fc775867c8a244a83f758a3c5b9a9f80b6a2ab0389520'
CASES=json.loads((REPO/'versions/G1R12/validation_runtime/data/cases.json').read_text(encoding='utf-8'))
def area(p):
    x0,y0=p[0]
    return .5*sum((a[0]-x0)*(b[1]-y0)-(b[0]-x0)*(a[1]-y0) for a,b in zip(p,p[1:]+p[:1]))
def pairs(s):return [[float(x.strip()) for x in p.split(',')] for p in re.findall(r'\(([^)]+)\)',s)]
def between(p,a,b):
    dx,dy=b[0]-a[0],b[1]-a[1];px,py=p[0]-a[0],p[1]-a[1]
    dd=dx*dx+dy*dy;t=(px*dx+py*dy)/dd
    return 0<t<1 and abs(px*dy-py*dx)<=1e-12*dd
def subdivide(poly,points):
    out=[]
    for a,b in zip(poly,poly[1:]+poly[:1]):
        dx,dy=b[0]-a[0],b[1]-a[1]
        cuts=[p for p in points if between(p,a,b)]
        out.append(a)
        out.extend(sorted(cuts,key=lambda p:((p[0]-a[0])*dx+(p[1]-a[1])*dy)))
    assert len(set(out))==len(out)
    return out
def emit(c):
    points=[]
    def add(p):
        p=tuple(map(float,p))
        if p not in points:points.append(p)
        return p
    for p in c['outline']:add(p)
    for reg in c['regions']:
        for p in reg['xy']:add(p)
    if c['surcharge']:
        assert c['id']=='D07A';add((2,10));add((8,10))
    regions=[subdivide([tuple(p) for p in reg['xy']],points) for reg in c['regions']]
    exterior=subdivide([tuple(p) for p in c['outline']],points)
    counts=collections.Counter(tuple(sorted((a,b))) for poly in regions for a,b in zip(poly,poly[1:]+poly[:1]))
    declared=set(tuple(sorted((a,b))) for a,b in zip(exterior,exterior[1:]+exterior[:1]))
    assert declared=={k for k,n in counts.items() if n==1} and all(n in [1,2] for n in counts.values())
    assert math.isclose(sum(area(p) for p in regions),c['area'],rel_tol=1e-12)
    xmin=min(p[0] for p in points);xmax=max(p[0] for p in points);ymin=min(p[1] for p in points)
    faces=[];boundary=[]
    for ri,poly in enumerate(regions):
        for j,(a,b) in enumerate(zip(poly,poly[1:]+poly[:1])):
            if counts[tuple(sorted((a,b)))]!=1:continue
            kind='自由';load=[0.,0.]
            if a[1]==b[1]==ymin:kind='固定XY'
            elif a[0]==b[0] and a[0] in [xmin,xmax]:kind='固定X'
            if c['surcharge'] and set([a,b])=={(2.,10.),(8.,10.)}:load=[0.,-20.]
            faces.append([len(faces)+1,ri+1,points.index(a)+1,points.index(b)+1,int(j==len(poly)-1),kind,0,*load,0.,0.])
            boundary.append({'a':a,'b':b,'kind':kind,'load0':load,'load1':[0,0],'region':ri+1})
    if c['id'].startswith('P07'):assert not any(f['a'][0]==f['b'][0]==xmax for f in boundary)
    if c['surcharge']:assert sum(f['load0'][1]*math.dist(f['a'],f['b']) for f in boundary)==-120
    names=[m['name'] for m in c['materials']]
    regs=[[i+1,names.index(c['regions'][i]['name'])+1,0,','.join(str(points.index(p)+1) for p in poly)] for i,poly in enumerate(regions)]
    mats=[[i+1,m['c'],m['phi'],m['gamma'],m['psi']] for i,m in enumerate(c['materials'])]
    settings={'MODE':'Fs','TARGET':1536,'GRAVITY':'固定','ADAPT':1,'CYCLES':3,'QUADRATURE':8,'FS_TOL':.0001,'GAP':.01,'C0_AUTO':0,'RAW_LOWER_AUDIT':1,'TEST_INJECTION':0,'HSD_CAPTURE':0,'DIAGNOSTIC_MODE':1,'ANALYSIS_POLICY':'ASSOCIATED' if c['policy']=='equal_reduced_phi' else 'DAVIS_EQUIVALENT'}
    return {'case':c['id'],'source_case':c,'points':[[i+1,*p] for i,p in enumerate(points)],'regions':regs,'materials':mats,'faces':faces,'boundary_explicit':boundary,'settings':settings,'region_areas':[area(p) for p in regions],'total_weight':sum(area(poly)*c['materials'][names.index(reg['name'])]['gamma'] for poly,reg in zip(regions,c['regions'])),'scope_notes':['G1R12 production TARGET1536/adapt3/q8; not catalogue M0/M1/M2 mesh convergence','E/nu not variables of rigid-plastic formulation; D04/D09 cannot test elastic sensitivity','N denotes Davis equivalent associated formulation, not direct nonassociated elastoplastic FEM']}
def prepare():
    assert hashlib.sha256(SOURCE.read_bytes()).hexdigest()==SOURCE_SHA
    for folder in ['inputs','cases','results','tests']: (ROOT/folder).mkdir(exist_ok=True)
    text=CATALOG.read_text(encoding='utf-8')
    checked=[]
    for c in CASES:
        if c['pore']:continue
        section=re.search(r'^### '+c['id']+r' .*?(?=^### |\Z)',text,re.M|re.S)[0]
        assert pairs(re.search(r'外形（反時計回り）: `([^`]+)`',section)[1])==c['outline']
        mats=re.findall(r'^\| ([^|]+) \| ([\d.e+-]+) \| ([\d.e+-]+) \| ([\d.e+-]+) \| ([\d.e+-]+) \| ([\d.e+-]+) \| ([\d.e+-]+) \|',section,re.M)
        assert len(mats)==len(c['materials'])
        for row,m in zip(mats,c['materials']):assert row[0].strip()==m['name'] and list(map(float,row[1:]))==[m[k] for k in ['E','nu','gamma','c','phi','psi']]
        regs=re.findall(r'^- `([^`]+)`: `([^`]+)`',section,re.M)
        assert [{'name':name,'xy':pairs(coords)} for name,coords in regs]==c['regions']
        rec=emit(c)
        p=ROOT/'inputs'/f"{c['id']}.json";p.write_text(json.dumps(rec,ensure_ascii=False,indent=2),encoding='utf-8',newline='\n')
        checked.append({'case':c['id'],'points':len(rec['points']),'regions':len(rec['regions']),'boundary_faces':len(rec['faces']),'area':c['area'],'weight':rec['total_weight'],'json_md_match':True,'sha256':hashlib.sha256(p.read_bytes()).hexdigest()})
    assert len(checked)==25
    if CATALOG.resolve()!=(ROOT/'Test_Case_Catalog.md').resolve():shutil.copy2(CATALOG,ROOT/'Test_Case_Catalog.md')
    identity={'path':str(SOURCE),'sha256':SOURCE_SHA,'repo_commit':'48a839654fd72d045f5d4e85cdf238dd217816b5','catalog_sha256':hashlib.sha256(CATALOG.read_bytes()).hexdigest(),'cases':checked}
    (ROOT/'source_identity.json').write_text(json.dumps(identity,ensure_ascii=False,indent=2),encoding='utf-8')
    with (ROOT/'results/progress.csv').open('w',encoding='utf-8-sig',newline='') as f:
        w=csv.writer(f);w.writerow(['case','status','fs_obtained','lower_fs','upper_fs','target_met','seconds','error'])
        for r in checked:w.writerow([r['case'],'NOT_RUN','','','','','',''])
    print('Catalog/JSON/explicit boundaries/areas/weights PASS:',len(checked))
if __name__=='__main__':prepare()
