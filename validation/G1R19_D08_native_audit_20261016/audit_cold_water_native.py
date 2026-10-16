"""Solve-free independent audit of the fresh coupled D08 Excel final fields.

The RP25 body snapshot is checked against the common piecewise hydrostatic
force. Internal waterline mesh constraints are geometric protection only; they
are never converted into physical fixed boundaries. Saved Fs/x are unchanged.
"""
from pathlib import Path
import argparse, collections, csv, hashlib, importlib.util, json, math, os, sys, zipfile
import xml.etree.ElementTree as ET

for name in ['OPENBLAS_NUM_THREADS','OMP_NUM_THREADS','MKL_NUM_THREADS']:
    os.environ[name]='1'
ROOT=Path(__file__).resolve().parent
REPO=ROOT.parents[1]
spec=importlib.util.spec_from_file_location('dry_binary_primitives',REPO/'validation/G1R19_D02_route_audit_20261016/audit_cold_native.py')
base=importlib.util.module_from_spec(spec);spec.loader.exec_module(base)
core,np,sparse=base.core,base.np,base.sparse
require,exact,one=base.require,base.exact,base.one


def sha(path):return hashlib.sha256(Path(path).read_bytes()).hexdigest()
def load_module(name,path):
    spec=importlib.util.spec_from_file_location(name,path)
    module=importlib.util.module_from_spec(spec);spec.loader.exec_module(module);return module


def common_body(xy,tri,materials,ids,body):
    y=xy[tri,1]
    require(not np.any((y.min(axis=1)<3.)&(y.max(axis=1)>3.)),'Triangle crosses the prescribed waterline')
    expected=np.zeros_like(body)
    expected[:,1,1]=-materials[ids,3]+np.where(y.max(axis=1)<=3.,9.81,0.)
    np.testing.assert_allclose(body,expected,rtol=0,atol=1e-12)
    return expected


def constraint_kind(key,xy,boundary):
    if key in boundary:return 'exterior'
    require(all(xy[i,1]==3. for i in key),'Interior geometric constraint is not on the waterline')
    return 'waterline_geometric'


def exact_model(model,state,mesh):
    tokens=model.split('|');nn,ne,nm=len(mesh.xy),len(mesh.tri),len(mesh.materials)
    area,_,edges=core.topology(mesh);header=tokens[0].split(',')
    require(header[:7]==[str(v) for v in [nn,ne,nm,0,len(edges)]]+['fixed','reference_total'],'Native header differs')
    require(exact(header[7])==mesh.stress_scale and header[8:]==['ASSOCIATED','raw_audit=True'],'Native policy/scale differs')
    pos=1
    for point in mesh.xy:
        np.testing.assert_array_equal([exact(tokens[pos][:16]),exact(tokens[pos][16:])],point);pos+=1
    for cell,body in zip(state['cells'],mesh.body):
        fields=tokens[pos].split(',');np.testing.assert_array_equal(list(map(int,fields[:5])),cell)
        np.testing.assert_array_equal([[exact(v[:16]),exact(v[16:])] for v in fields[5:]],body);pos+=1
    for mat in mesh.materials:
        np.testing.assert_array_equal([exact(tokens[pos][j:j+16]) for j in range(0,64,16)],mat);pos+=1
    expected={(edge.a,edge.b):edge for edge in edges};seen=set()
    for row in tokens[pos:pos+len(edges)]:
        fields=row.split(',');a,b,e,other=map(int,fields[:4]);key=tuple(sorted((a,b)))
        require(key in expected and key not in seen,'Unexpected/duplicate native edge');seen.add(key)
        edge=expected[key]
        require(sorted([e,other])==sorted(edge.elements) and fields[4]==edge.kind and int(fields[5])==edge.rigid,
                'Physical edge topology/restraint differs')
        np.testing.assert_array_equal([[exact(v[:16]),exact(v[16:])] for v in fields[6:]],edge.loads)
    require(len(seen)==len(edges),'Native physical edge omitted');pos+=len(edges)
    constraints={};counts=collections.Counter()
    for row in tokens[pos:]:
        if not row:continue
        text,a,b=row.split(',');a,b=int(a),int(b);key=tuple(sorted((a,b)))
        require(text==':'.join(map(str,key)) and key in expected and key not in constraints,'Invalid geometric constraint')
        kind=constraint_kind(key,mesh.xy,mesh.boundary);constraints[key]=kind;counts[kind]+=1
    require({key for key,kind in constraints.items() if kind=='exterior'}==set(mesh.boundary),'Exterior geometric constraint omitted')
    water_edges={key for key in expected if key not in mesh.boundary and all(mesh.xy[i,1]==3. for i in key)}
    require({key for key,kind in constraints.items() if kind=='waterline_geometric'}==water_edges,'Waterline protection differs')
    return dict(exact_binary_model_and_rebuilt_physical_topology_match=True,
                exact_native_model_sha256=hashlib.sha256(model.encode()).hexdigest(),
                geometric_constraint_counts=dict(counts),interior_waterline_constraints_are_not_physical_BCs=True)


def checkpoint(folder,name,side,record):
    binary=folder/'checkpoints'/(name+'.bin');metadata=binary.with_suffix('.meta.tsv')
    state,meta=base.state(binary.read_bytes()),base.metadata(metadata)
    require(one(meta,'schema')=='OWN18C01' and one(meta,'side')==side,'Wrong checkpoint schema/side')
    require(one(meta,'policy')=='ASSOCIATED' and one(meta,'phase')=='reference_total' and int(one(meta,'epoch'))==1,'Wrong policy/phase/epoch')
    require(int(one(meta,'q'))==8 and int(one(meta,'nodes'))==len(state['xy']) and int(one(meta,'elements'))==len(state['cells']),'Size/q differs')
    require(exact(one(meta,'value_exact'))==state['fs'] and exact(one(meta,'objective_exact'))==state['objective'],'Exact Fs/objective differs')
    cells=state['cells'];mats=state['materials'];xy=state['xy'];tri=cells[:,:3];ids=cells[:,3]
    np.testing.assert_array_equal(mats,np.array([[m[k] for k in ['c','phi','psi','gamma']] for m in record['source_case']['materials']]))
    require(state['scale']==max(1.,float(mats[:,0].max())),'Canonical stress scale differs')
    ref=load_module('canonical_water_boundary',REPO/'versions/G1R19_WaterCommon_20261016/tests/test_water_common_geometry.py')
    want=ref.canonical_boundary(xy,tri,record);require(set(want)==set(state['boundary']),'Canonical exterior boundary differs')
    for edge,values in want.items():
        require((state['boundary'][edge]['kind'],state['boundary'][edge]['rigid'])==values[:2],'Physical restraint differs')
        np.testing.assert_array_equal(state['boundary'][edge]['loads'],np.array(values[2:]).reshape(2,2).T)
    common_body(xy,tri,mats,ids,state['body'])
    identity=hashlib.sha256(json.dumps(dict(source_case=record['source_case'],pressure='9.81*max(3-y,0)',
                                          continuum_body='common piecewise hydrostatic on waterline-conforming mesh'),sort_keys=True).encode()).hexdigest()
    mesh=core.Mesh(xy,tri,mats,ids,cells[:,4],state['body'],state['boundary'],state['scale'],identity,
                   provenance={'case':'D08A','psi_policy':record['source_case']['policy']}).validate()
    area,_,_=core.topology(mesh)
    require(abs(float(area.sum())-sum(record['region_areas']))<1e-9,'Domain/material-region area differs')
    np.testing.assert_allclose(np.bincount(ids,weights=area,minlength=len(mats)),record['region_areas'],rtol=0,atol=1e-9)
    np.testing.assert_allclose(area@state['body'][:,:,1],[0.,-2571.4],rtol=0,atol=1e-8)
    model_check=exact_model(one(meta,'original_model'),state,mesh)
    ne=len(tri);field=state['field'].copy()
    if side=='lower':
        starts=dict(map(int,value.split(',')) for value in meta['lower_stress_start'])
        require(len(field)==18*ne+1 and int(one(meta,'lower_load_variable'))==18*ne and
                set(starts)==set(range(ne)) and all(starts[e]==18*e for e in starts),'Lower layout differs')
    else:require(len(field)==12*ne,'Upper velocity layout differs')
    before=hashlib.sha256(field.tobytes()).hexdigest()
    audit=core.evaluate(core.Prepared(mesh,state['fs'],side,8),field)
    require(before==hashlib.sha256(field.tobytes()).hexdigest() and np.array_equal(field,state['field']),'Saved field mutated/scaled')
    require(audit['physical_pass'] and audit['witness'],'Original q8 physical/witness gate failed')
    comparable=-audit['objective'] if side=='lower' else audit['objective']
    require(abs(comparable-state['objective'])<=1e-10,'Saved native objective differs from independently rebuilt objective')
    result=dict(side=side,fs=state['fs'],nodes=len(xy),elements=ne,q=8,physical_model_id=identity,
                source=one(meta,'source'),mesh_id=mesh.mesh_id,physical_audit=audit,
                checkpoint=binary.relative_to(REPO).as_posix(),checkpoint_sha256=sha(binary),metadata_sha256=sha(metadata),
                field_sha256=before,field_unchanged=True,fs_unchanged=True,original_model_check=model_check,
                root=base.root_contract(meta,side,state['fs']),area=float(area.sum()),reference_body_force=(area@state['body'][:,:,1]).tolist(),
                common_load=True,canonical_material_boundary_exact=True,waterline_crossing_triangles=0)
    return result,mesh


def workbook_water(path,summary):
    ns={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    relns='{http://schemas.openxmlformats.org/package/2006/relationships}'
    with zipfile.ZipFile(path) as archive:
        wb=ET.fromstring(archive.read('xl/workbook.xml'))
        relations={r.attrib['Id']:r.attrib['Target'] for r in ET.fromstring(archive.read('xl/_rels/workbook.xml.rels'))}
        rid=next(s.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'] for s in wb.findall('s:sheets/s:sheet',ns) if s.attrib['name']=='浸透結果')
        target=relations[rid];target=target.lstrip('/') if target.startswith('/') else 'xl/'+target
        sheet=ET.fromstring(archive.read(target));values={}
        for cell in sheet.findall('.//s:sheetData/s:row/s:c',ns):
            value=cell.find('s:v',ns)
            if value is not None and cell.attrib.get('t','n') in ['n','b']:
                values[cell.attrib['r']]=float(value.text)
    n=int(values['K7']);require(values['K2']==1. and n==int(summary['prescribed_water']['result_K2_K10'][5][0]),'Seep result qualification/nodes differ')
    require(n>=3,'Missing prescribed result nodes')
    coords=[]
    for row in range(2,n+2):
        coords.append([values['B'+str(row)],values['C'+str(row)]])
        require(values['A'+str(row)]==row-1 and values['D'+str(row)]==3.,'Prescribed result node ID/head differs')
        require(abs(values['E'+str(row)]-(3.-values['C'+str(row)]))<=1e-12,'Prescribed pressure head differs')
        require(abs(values['F'+str(row)]-9.81*(3.-values['C'+str(row)]))<=1e-10,'Prescribed signed water pressure differs')
    return dict(prescribed_result_nodes=n,all_saved_node_heads_exactly_3=True,
                signed_pressure_9p81_times_head_minus_y=True,analysis_suction_clamp=0,
                workbook_sha256=sha(path),node_coordinates_sha256=hashlib.sha256(np.array(coords).tobytes()).hexdigest())


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--run',type=Path,default=REPO/'validation/G1R19_WaterCommon_excel_20261016/results/owned_excel_1792152100656896200')
    parser.add_argument('--lower',default='00007_lower');parser.add_argument('--upper',default='00008_upper')
    parser.add_argument('--output',type=Path,default=ROOT/'COLD_NATIVE_D08_INDEPENDENT_AUDIT.json');args=parser.parse_args()
    folder=args.run.resolve();summary_path=folder/'summary.json';summary=json.loads(summary_path.read_text(encoding='utf8'))
    record_path=REPO/'docs/benchmarks/G1R17_native26_20261013/data/inputs/D08A.json';record=json.loads(record_path.read_text(encoding='utf8'))
    require(json.loads((folder/'input_record.json').read_text(encoding='utf8'))==record,'Canonical full input differs')
    require(summary['case']=='D08A' and summary['status']=='PASS' and summary['target_met'] and summary['owned_excel_exited'] is True,'Cold run not complete')
    require(summary['imported_field_or_root'] is False and summary['solver_threads']==1 and summary['kind']=='cold_initial_input','Fresh cold provenance differs')
    require(summary['analysis_seconds']<=900. and summary['budget_seconds']==900.,'Cold budget exceeded')
    require(summary['input_sha256']==sha(record_path) and summary['dll_sha256']==sha(folder/'rpfem_native_x64.dll'),'Input/DLL SHA differs')
    require(sha(folder/'WATER_SOURCE_MANIFEST.json')=='e25c0748e7a0c1b8127a2baa540d29448fe6d788ed04691cd601e931946e2f5c','Frozen water overlay manifest differs')
    protected=[record_path,summary_path,folder/'input_record.json',folder/'WATER_SOURCE_MANIFEST.json',folder/'rpfem_native_x64.dll',folder/'Adaptive_Probe.xlsm',folder/'harness_snapshot.py',Path(__file__)]
    require(sha(folder/'harness_snapshot.py')==summary['harness_sha256'],'Frozen harness differs')
    # The harness saves analysis output into this disposable book. Its
    # configured SHA records the earlier pre-analysis checkpoint, whereas the
    # final file inventory binds the actual post-analysis book read below.
    require(sha(folder/'Adaptive_Probe.xlsm')==summary['files']['Adaptive_Probe.xlsm']['sha256'],'Final workbook differs')
    for name,expected in summary['source_sha256'].items():
        path=folder/'source_snapshot'/(name+'.bas');require(sha(path)==expected,'Frozen source differs: '+name);protected.append(path)
    water=summary['prescribed_water']
    require(summary['prescribed_result'].startswith('PASS fresh prescribed mode2') and water['new_prescribed_mode2_result'] is True and water['seep_FEM_solve'] is False,'Wrong prescribed-water provenance')
    require(water['table']==[[0.,3.],[20.,3.]] and water['input_B3_F4'][1]==[9.81,'連成する',1.,'サクション',0.],'Water input differs')
    book_check=workbook_water(folder/'Adaptive_Probe.xlsm',summary)
    for name in [args.lower,args.upper]:protected.extend([folder/'checkpoints'/(name+'.bin'),folder/'checkpoints'/(name+'.meta.tsv')])
    before={str(p.relative_to(REPO)):sha(p) for p in protected}
    sides={};meshes={}
    for side in ['lower','upper']:sides[side],meshes[side]=checkpoint(folder,getattr(args,side),side,record)
    lower,upper=sides['lower']['fs'],sides['upper']['fs'];gap=2*(upper-lower)/(upper+lower)
    require(0<lower<=upper and gap<=.02,'2% physical bound gap failed')
    cold=summary['cold_result']
    require(cold[:10]==['AUDITED',True,lower,upper,sides['lower']['elements'],sides['upper']['elements'],8,8,True,True] and
            cold[10:14]==[sides['lower']['root']['a'],sides['lower']['root']['b'],sides['upper']['root']['a'],sides['upper']['root']['b']] and cold[14] is False,
            'Final result and own root metadata differ')
    require(sides['lower']['physical_model_id']==sides['upper']['physical_model_id'],'Physical load model differs')
    after={str(p.relative_to(REPO)):sha(p) for p in protected};require(before==after,'Audited original bytes changed')
    result=dict(status='INDEPENDENT_COMMON_WATER_PRIMAL_Q8_AND_NATIVE_ROOT_CONTRACT_PASS',case='D08A',solver_calls=0,
                lower_fs=lower,upper_fs=upper,gap=gap,gap_percent=100*gap,target_gap=.02,target_met=True,
                root_complete=True,field_or_fs_scaled=False,analysis_seconds=summary['analysis_seconds'],
                normal_cold_execution=True,owned_excel_exited=True,physical_model_id=sides['lower']['physical_model_id'],
                sides=sides,prescribed_water_result=book_check,source_sha256=before,source_sha256_after=after,
                original_bytes_unchanged=True,run=folder.relative_to(REPO).as_posix(),
                literal_original_exports_count=len(list((folder/'RPFEM_logs').glob('*.accuracy_*.bin'))),
                literal_replay_scope='Native harness reports 59 literal CSR replays; this checker independently rebuilds the final RP25 physical fields and complete exact model. It does not claim a second independent replay of every CSR or every rejected root endpoint dual.',
                root_proof_limit='Own metadata identity and normalized width plus fresh full primal audit at accepted endpoint. Opposite endpoint signs are retained native certificates, not independently re-solved.',
                model_discretization='Common waterline-conforming hydrostatic continuum body; internal waterline RConstraints are geometric only; exterior physical boundary remains canonical')
    core.write(args.output,result)
    print(json.dumps({k:result[k] for k in ['status','lower_fs','upper_fs','gap_percent','root_complete','target_met','analysis_seconds']},ensure_ascii=False))


if __name__=='__main__':main()
