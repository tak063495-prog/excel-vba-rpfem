"""Independent prescribed-pressure geometry/load replay; never solves Fs."""
from pathlib import Path
import collections,importlib.util,json,sys
import numpy as np

REPO=Path(__file__).resolve().parents[3]
spec=importlib.util.spec_from_file_location('mesh_reader',REPO/'versions/G1R19_2pct/tests/run_mesh_export_excel.py')
reader=importlib.util.module_from_spec(spec);spec.loader.exec_module(reader)
spec=importlib.util.spec_from_file_location('water_reference',REPO/'validation/G1R19_python26_water_20261016/python/waterline_conform.py')
ref=importlib.util.module_from_spec(spec);spec.loader.exec_module(ref)
sys.path.insert(0,str(REPO/'validation/accuracy_20261013/python'))
import experiment_core as core


def alternative_diagonals(xy,actual,expected,parents):
    different=np.flatnonzero(np.any(actual!=expected,axis=1));rows=[]
    def outer(cells):
        count=collections.Counter(tuple(sorted((int(a),int(b)))) for t in cells for a,b in zip(t,np.roll(t,-1)))
        assert max(count.values())<=2
        return {e for e,n in count.items() if n==1}
    def quality(cells):
        p=xy[cells];u=p[:,1]-p[:,0];v=p[:,2]-p[:,0]
        return float(np.min((u[:,0]*v[:,1]-u[:,1]*v[:,0])/np.sum((p-np.roll(p,1,axis=1))**2,axis=(1,2))))
    for parent in np.unique(np.asarray(parents)[different]):
        ids=different[np.asarray(parents)[different]==parent]
        assert len(ids)==2
        a=actual[ids];b=expected[ids]
        assert len(np.unique(a))==4 and set(a.ravel())==set(b.ravel())
        assert outer(a)==outer(b) and len(outer(a))==4
        from scipy.spatial import ConvexHull
        hull=ConvexHull(xy[np.unique(a)]);assert len(hull.vertices)==4
        aa,_=ref.triangle_moments(xy,a);ba,_=ref.triangle_moments(xy,b)
        np.testing.assert_allclose(aa.sum(),hull.volume,rtol=0,atol=1e-12)
        np.testing.assert_allclose(ba.sum(),hull.volume,rtol=0,atol=1e-12)
        qa,qb=quality(a),quality(b)
        assert qa>=qb-8*np.finfo(float).eps
        rows.append({'parent':int(parent),'rows':ids.tolist(),'actual_min_quality':qa,'reference_min_quality_at_actual_coordinates':qb,'same_convex_quad_boundary':True})
    return rows


def audit(folder):
    folder=Path(folder)
    record=json.loads((REPO/'docs/benchmarks/G1R17_native26_20261013/data/inputs/D08A.json').read_text(encoding='utf8'))
    spec=importlib.util.spec_from_file_location('canonical_water_boundary',REPO/'versions/G1R19_WaterCommon_20261016/tests/test_water_common_geometry.py')
    boundary_ref=importlib.util.module_from_spec(spec);spec.loader.exec_module(boundary_ref)
    rows={};force=None;moment=None;diagonals=[];partition=None
    base=reader.read_geometry(folder/'base_raw.bin')
    expected=ref.split_at_waterline(base[0],base[1][:,:3],base[1][:,3],3.)
    for path in sorted(folder.glob('*.bin')):
        xy,cells,mats,body,bd,meta=reader.read_geometry(path)
        canonical_mats=np.array([[m['c'],m['phi'],m['psi'],m['gamma']] for m in record['source_case']['materials']])
        np.testing.assert_array_equal(mats,canonical_mats)
        assert meta['q']==8 and meta['stress_scale']==max(1.,float(mats[:,0].max()))
        expected_boundary=boundary_ref.canonical_boundary(xy,cells[:,:3],record)
        assert bd.keys()==expected_boundary.keys()
        for edge,want in expected_boundary.items():
            got=bd[edge];assert (got['kind'],got['rigid'])==want[:2]
            np.testing.assert_array_equal(got['loads'],np.array(want[2:]).reshape(2,2).T)
        # The coupled native body is authoritative input here. Do not call the
        # dry model loader or reconstruct a pore-pressure case as a dry solve.
        mesh=core.Mesh(xy,cells[:,:3],mats,cells[:,3],cells[:,4],body,bd,meta['stress_scale'],'D08A original prescribed pressure snapshot').validate()
        area,first=ref.triangle_moments(xy,cells[:,:3]);assert np.all(area>0)
        np.testing.assert_allclose(area.sum(),158.,rtol=0,atol=1e-9)
        if path.stem=='base_raw':
            total=np.column_stack((np.zeros(len(cells)),-mats[cells[:,3],3]))
            np.testing.assert_array_equal(body.sum(axis=2),total)
        else:
            total=ref.exact_hydrostatic_body(mesh,3.,9.81).sum(axis=2)
            np.testing.assert_allclose(body.sum(axis=2),total,rtol=0,atol=1e-12)
            gotforce=area@body.sum(axis=2);gotmoment=first.T@body.sum(axis=2)
            np.testing.assert_allclose(gotforce,[0.,-2571.4],rtol=0,atol=1e-8)
            if force is None:force=gotforce;moment=gotmoment
            np.testing.assert_allclose(gotforce,force,rtol=0,atol=1e-8)
            np.testing.assert_allclose(gotmoment,moment,rtol=0,atol=1e-7)
            if path.stem=='common_raw':
                np.testing.assert_allclose(xy,expected[0],rtol=0,atol=1e-12)
                np.testing.assert_array_equal(cells[:,3],expected[2])
                diagonals=alternative_diagonals(xy,cells[:,:3],expected[1],expected[3]['parent_element_ids'])
                partition=ref.verify_partition(base[0],base[1][:,:3],base[1][:,3],xy,cells[:,:3],cells[:,3],expected[3])
            if path.stem.endswith('reference') or path.stem.endswith('idempotent'):
                np.testing.assert_array_equal(body[:,:,0],np.zeros_like(body[:,:,0]))
            else:np.testing.assert_array_equal(body[:,:,1],np.zeros_like(body[:,:,1]))
        rows[path.name]={'nodes':len(xy),'elements':len(cells),'area':float(area.sum()),'force':(area@body.sum(axis=2)).tolist(),'first_force_moments':(first.T@body.sum(axis=2)).tolist(),'canonical_material_body_boundary':True}
    assert len(rows)==8
    common=reader.read_geometry(folder/'common_reference.bin')
    idem=reader.read_geometry(folder/'common_idempotent.bin')
    for i in [0,1,2,3]:np.testing.assert_array_equal(common[i],idem[i])
    result={'status':'PASS','scope':'original physical prescribed pressure, same exported geometry, solver0','case':'D08A','level':3.,'gamma_water':9.81,'external_water_force_added':False,'initial_clipping_connectivity_equal':not diagonals,'admissible_alternative_diagonal_quality_ties':diagonals,'original_parent_partition_audit':partition,'idempotent':True,'stages':rows}
    (folder/'INDEPENDENT_MICRO_AUDIT.json').write_text(json.dumps(result,indent=2)+'\n',encoding='utf8')
    return result

if __name__=='__main__':print(json.dumps(audit(Path(sys.argv[1])),indent=2))
