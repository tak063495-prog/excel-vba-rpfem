"""Check immutable rollup and strict root identity using a real clay witness pair."""
from __future__ import annotations
import copy
import json
from pathlib import Path
import shutil
import tempfile

import report_accuracy as report
from experiment_core import write


def main():
    results=report.driver.REPO/'validation/accuracy_20261013/results'
    source=results/'clay_combined/P08A'
    if not (source/'summary.json').is_file():
        raise ValueError('The completed P08A combined certificate is required')
    original=json.loads((source/'summary.json').read_text(encoding='utf8'))
    rec=json.loads((report.driver.CATALOG/'P08A.json').read_text(encoding='utf8'))
    mesh=report.driver.load_saved_mesh(source/'best_lower_mesh.npz',rec)
    metadata=original['lower'];fs=metadata['fs']
    assert report.root_is_complete(metadata,mesh,fs,'lower','P08A')
    bad=[]
    for key,value in [('case','P09A'),('mode','upper'),('q',4),('mesh_id','different'),
                      ('physical_model_id','different'),('root_complete',False)]:
        d=copy.deepcopy(metadata);d[key]=value
        assert not report.root_is_complete(d,mesh,fs,'lower','P08A');bad.append('outer_'+key)
    for key,value in [('case','P09A'),('mode','upper'),('q',4),('policy','different'),
                      ('mesh_id','different'),('physical_model_id','different'),
                      ('root_complete',False),('bracket_signs',[-1,1]),
                      ('endpoint',fs+1e-6),('bracket',[fs,fs+.01]),('bracket',[fs,float('nan')])]:
        d=copy.deepcopy(metadata);d['audit']['discrete_root_search'][key]=value
        assert not report.root_is_complete(d,mesh,fs,'lower','P08A');bad.append('root_'+key)
    with tempfile.TemporaryDirectory(prefix='report_checks_',dir=results) as tmp:
        workspace=Path(tmp).resolve()
        assert workspace.is_relative_to(results.resolve())
        stable=workspace/'sources/P08A';stable.mkdir(parents=True)
        for name in ['summary.json','best_lower_mesh.npz','best_lower_field.npz',
                     'best_upper_mesh.npz','best_upper_field.npz']:
            shutil.copyfile(source/name,stable/name)
        corrupt=workspace/'sources/corrupt';shutil.copytree(stable,corrupt)
        d=copy.deepcopy(original);d['lower']['fs']+=.1;d['upper']['fs']+=.1
        write(corrupt/'summary.json',d)
        out=workspace/'sources/rollup'
        payload=report.run([workspace/'sources'],out)
        cases={d['case']:d for d in payload['cases']}
        assert len(cases)==26 and cases['P08A']['target_met'] and cases['P08A']['root_complete']
        assert len(payload['candidate_audits'])==2
        assert len(payload['skipped_or_rejected'])==2
        assert all('inconsistent' in d['reason'] for d in payload['skipped_or_rejected'])
        assert not payload['all_26_targets_revalidated']
        assert all(not d['target_met'] for c,d in cases.items() if c!='P08A')
        assert len(report.source_files([workspace/'sources'],out))==2
        for mode in ['lower','upper']:
            metadata=cases['P08A'][mode]
            assert metadata['source_hashes'] and metadata['root_complete']
            snapshot=out/'snapshots/P08A'
            for name in [metadata['mesh_path'],metadata['field_path'],f'best_{mode}_certificate.json']:
                assert (snapshot/name).is_file()
            ownmesh=report.driver.load_saved_mesh(snapshot/metadata['mesh_path'],rec)
            assert report.root_is_complete(metadata,ownmesh,metadata['fs'],mode,'P08A')
        tracked={};probe=workspace/'mutable.txt';probe.write_bytes(b'first')
        report.read_tracked(probe,tracked);assert report.unchanged(tracked)
        probe.write_bytes(b'second');assert not report.unchanged(tracked)
        try:report.inside_file(stable,'../outside.npz')
        except ValueError:pass
        else:raise AssertionError('Path escape was accepted')
        result={'pass':True,'actual_same_model_physical_pair':'P08A',
                'reaudited_target_width':cases['P08A']['relative_width_symmetric'],
                'strict_root_rejections':bad,'corrupt_metadata_rejected':2,
                'all_26_table_rows':len(cases),'immutable_snapshot_checks':True,
                'mutable_hash_change_rejected':True,'source_output_recursive_scan_excluded':True}
    write(results/'report_checks.json',result)
    print(json.dumps(result,ensure_ascii=False,indent=2),flush=True)


if __name__=='__main__':main()
