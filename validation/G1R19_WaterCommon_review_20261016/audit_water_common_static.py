"""Read-only independent overlay review checks; no Excel or numerical solves."""
import collections, hashlib, importlib.util, json, re, unittest
from pathlib import Path
import numpy as np
from scipy.spatial import Delaunay

HERE = Path(__file__).resolve().parent
REPO = HERE.parents[1]
ROOT = REPO/'versions/G1R19_WaterCommon_20261016'
spec = importlib.util.spec_from_file_location('watercommon_translation', ROOT/'tests/test_water_common_geometry.py')
translation = importlib.util.module_from_spec(spec); spec.loader.exec_module(translation)


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def measures(xy, tri):
    p = xy[tri]
    a, b = p[:, 1]-p[:, 0], p[:, 2]-p[:, 0]
    area = (a[:, 0]*b[:, 1]-a[:, 1]*b[:, 0])/2
    return area, area[:, None]*p.mean(axis=1)


def wet_area(points, level):
    """Independent direct clipped-polygon area, without child triangulation."""
    clipped = []
    for previous, current in zip(np.roll(points, 1, axis=0), points):
        before, after = previous[1] <= level, current[1] <= level
        if before != after:
            fraction = (level-previous[1])/(current[1]-previous[1])
            clipped.append(previous+fraction*(current-previous))
        if after:
            clipped.append(current)
    if len(clipped) < 3:
        return 0.
    # Translate first to avoid cancellation of global xy products.
    p = np.asarray(clipped)-clipped[0]
    return float(abs(np.sum(p[:, 0]*np.roll(p[:, 1], -1)-p[:, 1]*np.roll(p[:, 0], -1)))/2)


def main():
    modules = {}
    procedure = re.compile(r'^(?:(?:Public|Private|Friend) )?(?:Static )?(Sub|Function) (\w+)\b.*?^End \1[ \t]*$', re.M|re.S)
    manifest = json.loads((ROOT/'SOURCE_MANIFEST.json').read_text(encoding='utf8'))
    for entry in manifest['modules']:
        name = entry['module']; oldpath = REPO/entry['base']; newpath = ROOT/entry['source']
        old = oldpath.read_text(encoding='utf8'); new = newpath.read_text(encoding='utf8')
        imported = ROOT/entry['imported']
        assert sha(oldpath) == entry['base_sha256'] and sha(newpath) == entry['source_sha256'] and sha(imported) == entry['import_sha256']
        assert new == imported.read_text(encoding='cp932')
        before = {m[2]:m[0] for m in procedure.finditer(old)}
        after = {m[2]:m[0] for m in procedure.finditer(new)}
        changed = [key for key in before.keys() & after.keys() if before[key] != after[key]]
        assert not before.keys()-after.keys()
        expected = {'RPX_Mesh':{'RPX_AssignRegions'},'RPX_Water':{'RPX_WaterApply'},'RPX_Guard':{'RPX_InputBegin','RPX_InputEnd','RPX_AssertInputs'}}
        assert set(changed) == expected[name]
        modules[name] = dict(source_sha256=sha(newpath), normalized_changed_existing_functions=changed,
                             added_functions=sorted(after.keys()-before.keys()), import_decoded_equal=True)
    rng = np.random.default_rng(20261016)
    worst_area = worst_moment = worst_wet = 0.
    total_parents = total_children = 0
    for trial in range(100):
        xy = rng.random((30, 2))*[20, 10]
        tri = Delaunay(xy).simplices.copy()
        a, _ = measures(xy, tri)
        tri[a < 0] = tri[a < 0, ::-1]
        ids = rng.integers(0, 3, len(tri))
        level = float(rng.uniform(1, 9))
        count = collections.Counter(tuple(sorted((int(a),int(b)))) for t in tri for a,b in zip(t,np.roll(t,-1)))
        boundary = {key:('fixed' if j%2 else 'load', -1, 1., 2., 3., 4.) for j,(key,n) in enumerate(count.items()) if n == 1}
        frozen = [a.copy() for a in (xy,tri,ids)]
        got = translation.translated_split(xy,tri,ids,level,dict(boundary),boundary)
        for old, source in zip(frozen,(xy,tri,ids)):
            np.testing.assert_array_equal(old,source)
        np.testing.assert_array_equal(got[0][:len(xy)],xy)
        np.testing.assert_array_equal(got[2],ids[got[3]])
        oldarea, oldmoment = measures(xy,tri)
        area, moment = measures(*got[:2])
        assert np.all(area > 0)
        errors = [abs(float(area.sum()-oldarea.sum())), float(np.max(abs(moment.sum(axis=0)-oldmoment.sum(axis=0))))]
        worst_area=max(worst_area, errors[0]); worst_moment=max(worst_moment, errors[1])
        assert errors[0] < 1e-10 and errors[1] < 1e-9
        points=got[0][got[1]]
        assert not np.any((points[:,:,1].min(axis=1) < level)&(points[:,:,1].max(axis=1)>level))
        actual_wet=float(area[points[:,:,1].max(axis=1)<=level].sum())
        expected_wet=sum(wet_area(p,level) for p in xy[tri])
        worst_wet=max(worst_wet,abs(actual_wet-expected_wet))
        assert abs(actual_wet-expected_wet)<1e-10
        children=collections.Counter(tuple(sorted((int(a),int(b)))) for t in got[1] for a,b in zip(t,np.roll(t,-1)))
        assert max(children.values()) == 2
        assert {edge for edge,n in children.items() if n==1} == got[6].keys()
        for edge,n in children.items():
            if all(got[0][node,1] == level for node in edge):
                assert edge in got[5]
                if n==2: assert edge not in got[6]
        total_parents+=len(tri);total_children+=len(got[1])
    # Run only tests that do not write the author's frozen fixture/result.
    names=[name for name in unittest.defaultTestLoader.getTestCaseNames(translation.WaterCommonTests)
           if name != 'test_actual_d08_partition_and_common_load']
    suite=unittest.TestSuite(translation.WaterCommonTests(name) for name in names)
    result=unittest.TextTestRunner(verbosity=1).run(suite)
    assert result.wasSuccessful()
    output=dict(status='STATIC_PASS_EXCEL_RUNTIME_PENDING', scope='independent source review and scalar geometry checks; no VBA execution or solver calls',
                modules=modules, independent_random_meshes=100, parent_triangles=total_parents, child_triangles=total_children,
                worst_total_area_error=worst_area,worst_total_first_moment_error=worst_moment,
                worst_independent_clipped_wet_area_error=worst_wet,
                nonwriting_author_tests=result.testsRun,
                author_fixture_or_source_modified=False,
                original_blocker='Standard input snapshot excluded seep input/result sheets. Isolated overlay now adds eligible exact water-source begin/assert/end hooks; actual VBA mutation tests remain pending.')
    (HERE/'STATIC_REVIEW_CHECKS.json').write_text(json.dumps(output,ensure_ascii=False,indent=2)+'\n',encoding='utf8')
    print(json.dumps({k:v for k,v in output.items() if k!='modules'},ensure_ascii=False))


if __name__=='__main__':
    main()
