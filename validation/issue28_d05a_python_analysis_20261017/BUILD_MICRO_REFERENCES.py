"""Capture original-host coefficient bytes, bound to frozen producer digests.

This creates no solution. It refuses overwrite and refuses a host assembly
that does not match the original six-run metadata exactly.
"""
import json
from pathlib import Path
import sys
sys.dont_write_bytecode=True
from run_fixed_point_experiment import ROOT,core,frozen_hashes,load_model,sha,utc,write
from portable_program_audit import pack_program,unpack_program,program_digest

destination=ROOT/'program_refs'
if destination.exists():raise FileExistsError(destination)
destination.mkdir()
before=frozen_hashes()
manifest={'schema':'issue28_micro_original_program_references_v1','generated_utc':utc(),
    'scope':'Original-host coefficient snapshots only; each raw program digest equals the frozen experiment metadata. No optimizer call.',
    'optimizer_calls':0,'source_data_manifest_sha256':sha((ROOT/'SOURCE_DATA_MANIFEST.json').read_bytes()),
    'producer_script_sha256':sha((ROOT/'run_fixed_point_experiment.py').read_bytes()),
    'comparator_sha256':sha((ROOT/'portable_program_audit.py').read_bytes()),
    'builder_sha256':sha(Path(__file__).read_bytes()),'programs':{}}
for case in ('P02A','D05A'):
    mesh,_=load_model(case);pr=core.Prepared(mesh,1.,'upper',8,reduction='none')
    digest=program_digest(pr.original);records=[]
    for meta_path in sorted((ROOT/'results/fixed_point_02').glob(case+'*/result.json')):
        meta=json.loads(meta_path.read_text(encoding='utf-8'))
        assert meta['original_program']['sha256']==digest,case+' current host differs from original producer'
        assert meta['source_data_hashes_before']==before==meta['source_data_hashes_after']
        records.append({'path':meta_path.relative_to(ROOT).as_posix(),'sha256':sha(meta_path.read_bytes()),'recorded_original_program_sha256':digest})
    assert len(records)==(2 if case=='P02A'else 4)
    arrays=pack_program(pr.original);target=destination/(case+'_original_program.npz')
    core.np.savez_compressed(target,**arrays)
    with core.np.load(target,allow_pickle=False)as data:
        restored=unpack_program({name:data[name]for name in data.files})
        assert program_digest(restored)==digest
    manifest['programs'][case]={'path':target.relative_to(ROOT).as_posix(),'bytes':target.stat().st_size,
        'sha256':sha(target.read_bytes()),'original_program_sha256':digest,
        'array_layout':{name:{'shape':array.shape,'dtype':str(array.dtype)}for name,array in arrays.items()},
        'frozen_producer_records':records}
assert before==frozen_hashes()
write(destination/'PROGRAM_REFERENCE_MANIFEST.json',manifest)
print(json.dumps({'status':'PASS','programs':len(manifest['programs']),'optimizer_calls':0,'manifest_sha256':sha((destination/'PROGRAM_REFERENCE_MANIFEST.json').read_bytes())}))
