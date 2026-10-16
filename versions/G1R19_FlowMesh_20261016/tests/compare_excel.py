"""Compare a solver-free Excel candidate against immutable Python expectations."""
from pathlib import Path
import argparse,json
import reference_fixtures as ref
np=ref.np

def compare(fixture_root,output_root):
    manifest=json.loads((fixture_root/'FIXTURES.json').read_text(encoding='utf8'));rows=[]
    for entry in manifest['fixtures']:
        input_path=fixture_root/entry['input'];expected_path=fixture_root/entry['expected_tri']
        if ref.sha(input_path)!=entry['input_sha256'] or ref.sha(expected_path)!=entry['expected_tri_sha256']:
            raise ValueError('Immutable fixture bytes changed')
        xy,tri,mat,velocity,guards,passes=ref.read_input(input_path)
        actual,info=ref.read_output(output_root/(entry['name']+'.bin'))
        with np.load(expected_path,allow_pickle=False) as data:expected=data['tri'].copy()
        exact=np.array_equal(actual,expected)
        # Always independently verify material areas, protected edges and CCW.
        def edges(cells):return {tuple(sorted((int(a),int(b)))) for t in cells for a,b in zip(t,np.roll(t,-1))}
        def area(cells):
            p=xy[cells];u=p[:,1]-p[:,0];v=p[:,2]-p[:,0]
            return (u[:,0]*v[:,1]-u[:,1]*v[:,0])/2
        before,after=area(tri),area(actual)
        if not np.all(after>5e-13) or not set(map(tuple,guards)).issubset(edges(actual)):
            raise ValueError('VBA candidate orientation/protected edge failed')
        for material in np.unique(mat):np.testing.assert_allclose(before[mat==material].sum(),after[mat==material].sum(),rtol=1e-12,atol=1e-10)
        if info['certificate'] or info['nodes']!=len(xy) or info['elements']!=len(tri):raise ValueError('VBA candidate is not geometry-only')
        row=dict(name=entry['name'],exact_python_connectivity=exact,python_flips=entry['flips'],vba_flips=info['flips'],
                 vba_seconds=info['seconds'],independent_geometry_pass=True,certificate=False,solver_calls=0,
                 output_sha256=ref.sha(output_root/(entry['name']+'.bin')))
        rows.append(row)
        if not exact:raise ValueError('VBA/Python connectivity differs; preserve output and inspect numeric tie/ordering: '+entry['name'])
    return dict(status='EXCEL_PYTHON_EXACT_GEOMETRY_PASS',solver_calls=0,fixtures=rows)

if __name__=='__main__':
    p=argparse.ArgumentParser(description=__doc__);p.add_argument('--fixtures',type=Path,required=True);p.add_argument('--outputs',type=Path,required=True)
    a=p.parse_args();result=compare(a.fixtures,a.outputs)
    (a.outputs/'PYTHON_COMPARISON.json').write_text(json.dumps(result,indent=2)+'\n',encoding='utf8');print(json.dumps(result))
