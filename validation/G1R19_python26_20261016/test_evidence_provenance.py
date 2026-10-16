"""Synthetic/temp-only provenance rejection tests; no actual run artifacts."""
import copy
import hashlib
import json
from pathlib import Path
import tempfile
import unittest
from evidence_provenance import verify_source_provenance


def digest(raw):return hashlib.sha256(raw).hexdigest()
def canonical(value):return json.dumps(value,sort_keys=True,separators=(',',':'),ensure_ascii=False).encode()
def write(path,value):
    path.parent.mkdir(parents=True,exist_ok=True)
    path.write_text(json.dumps(value,indent=2,ensure_ascii=False)+'\n',encoding='utf8')


class ProvenanceTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory();self.addCleanup(self.temp.cleanup)
        self.repo=Path(self.temp.name);self.case='D01A'
        self.producer=self.repo/'validation/prototype';self.python=self.producer/'python'
        self.folder=self.producer/'results/fresh/cases'/self.case
        self.folder.mkdir(parents=True)
        self.input=self.repo/'docs/benchmarks/G1R17_native26_20261013/data/inputs'/f'{self.case}.json'
        self.record={'case':self.case,'source_case':{'id':self.case,'policy':'equal_reduced_phi','pore':False,'gamma':20.},'water':{'coupling':0}}
        write(self.input,self.record)
        for name in ('accuracy_driver.py','experiment_core.py','benchmark_models.py','vendor/independent_audit.py'):
            path=self.python/name;path.parent.mkdir(parents=True,exist_ok=True);path.write_text('# immutable numerical implementation\n')
        (self.producer/'run_case.py').write_text('# immutable producer wrapper\n')
        geometry=self.repo/'geometry.bin';geometry.write_bytes(b'fixed initial geometry')
        entry={'case':self.case,'policy':'equal_reduced_phi',
               'input':{'path':self.input.relative_to(self.repo).as_posix(),'sha256_lf':digest(self.input.read_bytes().replace(b'\r\n',b'\n')),'sha256_canonical_json':digest(canonical(self.record))},
               'initial_mesh':{'path':'geometry.bin','sha256':digest(geometry.read_bytes())}}
        self.manifest={'cases':[entry]+[{'case':'OTHER'+str(i)} for i in range(25)]}
        write(self.producer/'INPUTS_26.json',self.manifest)
        self.model=digest(json.dumps(self.record['source_case'],sort_keys=True).encode())
        self.metadata={}
        for mode,fs,bracket in (('lower',1.,[1.,1.00005]),('upper',1.01,[1.00995,1.01])):
            mesh=digest(mode.encode())
            root={'case':self.case,'mode':mode,'q':8,'policy':'equal_reduced_phi','mesh_id':mesh,'physical_model_id':self.model,
                  'bracket':bracket,'bracket_signs':[1,-1],'endpoint':fs,'root_complete':True}
            meta={'case':self.case,'mode':mode,'q':8,'fs':fs,'mesh_id':mesh,'physical_model_id':self.model,'root_complete':True,
                  'audit':{'discrete_root_search':root},'mesh_path':f'best_{mode}_mesh.npz','field_path':f'best_{mode}_field.npz'}
            self.metadata[mode]=meta
            write(self.folder/f'best_{mode}_certificate.json',meta)
            for suffix in ('mesh','field'):(self.folder/f'best_{mode}_{suffix}.npz').write_bytes((mode+suffix).encode())
        self.summary={'case':self.case,'status':'PASS','source_unchanged':True,'environment':{},**copy.deepcopy(self.metadata)}
        self.refreeze()
        self.audit={'case':self.case,'sides':{}}
        for mode,meta in self.metadata.items():
            files=[self.input,self.folder/f'best_{mode}_certificate.json',self.folder/f'best_{mode}_mesh.npz',self.folder/f'best_{mode}_field.npz']
            self.audit['sides'][mode]={**copy.deepcopy(meta),'file_sha256':{str(p):digest(p.read_bytes()) for p in files}}

    def refreeze(self):
        files=list(self.python.rglob('*.py'))+[self.producer/'run_case.py',self.producer/'INPUTS_26.json']
        self.sources={p.relative_to(self.repo).as_posix():digest(p.read_bytes()) for p in files}
        self.summary['environment']['source_sha256']=dict(self.sources)
        write(self.folder/'source_manifest.json',{'source_sha256':self.sources})
        write(self.folder/'summary.json',self.summary)

    def verify(self,audit=True):
        return verify_source_provenance(self.folder/'summary.json',self.case,repo=self.repo,audit=self.audit if audit else None)

    def test_valid_and_budget_are_separate(self):
        result=self.verify();self.assertTrue(result['source_provenance_verified']);self.assertTrue(result['input_provenance_verified'])
        self.assertFalse(result['cold_execution_status_verified']);self.assertFalse(result['explicit_producer_end_source_map'])
        self.summary.update(status='BUDGET_EXHAUSTED',target_met=False,budget_expired=True,worker_exit_code=None)
        write(self.folder/'summary.json',self.summary)
        self.assertTrue(self.verify()['source_provenance_verified'])

    def test_changed_actual_source_rejected_despite_true_flag(self):
        (self.python/'experiment_core.py').write_text('# altered after producer finished\n')
        with self.assertRaisesRegex(ValueError,'producer source changed'):self.verify()

    def test_start_summary_and_explicit_end_maps_must_match(self):
        self.summary['environment']['source_sha256_at_checkpoint']=dict(self.sources)
        write(self.folder/'summary.json',self.summary)
        self.assertTrue(self.verify()['explicit_producer_end_source_map'])
        first=next(iter(self.sources));self.summary['environment']['source_sha256_at_checkpoint'][first]='0'*64
        write(self.folder/'summary.json',self.summary)
        with self.assertRaisesRegex(ValueError,'Start/end'):self.verify()

    def test_self_consistent_omitted_code_and_new_code_rejected(self):
        key=next(k for k in self.sources if k.endswith('independent_audit.py'))
        self.sources.pop(key);self.summary['environment']['source_sha256']=dict(self.sources)
        write(self.folder/'source_manifest.json',{'source_sha256':self.sources});write(self.folder/'summary.json',self.summary)
        with self.assertRaisesRegex(ValueError,'omits current'):self.verify()
        self.refreeze();(self.python/'new_algorithm.py').write_text('# untracked implementation\n')
        with self.assertRaisesRegex(ValueError,'omits current'):self.verify()

    def test_case_record_and_manifest_binding(self):
        self.summary['case']='D02A';write(self.folder/'summary.json',self.summary)
        with self.assertRaisesRegex(ValueError,'Summary case'):self.verify()
        self.summary['case']=self.case;write(self.folder/'summary.json',self.summary)
        self.manifest['cases'][0]['input']['path']=str(self.input.parent/'D02A.json')
        write(self.producer/'INPUTS_26.json',self.manifest);self.refreeze()
        with self.assertRaisesRegex(ValueError,'another canonical input'):self.verify()

    def test_canonical_input_and_initial_geometry_changes_rejected(self):
        self.record['source_case']['gamma']=19.;write(self.input,self.record)
        with self.assertRaisesRegex(ValueError,'Canonical input differs'):self.verify()
        self.record['source_case']['gamma']=20.;write(self.input,self.record)
        (self.repo/'geometry.bin').write_bytes(b'changed geometry')
        with self.assertRaisesRegex(ValueError,'initial catalog mesh changed'):self.verify()

    def test_canonical_model_and_strict_roots_rejected(self):
        path=self.folder/'best_lower_certificate.json';original=copy.deepcopy(self.metadata['lower'])
        for mutation in ('model','q','case','policy','signs','endpoint','width'):
            meta=copy.deepcopy(original)
            if mutation=='model':meta['physical_model_id']='0'*64
            elif mutation in ('q','case'):meta[mutation]=7 if mutation=='q' else 'D02A'
            elif mutation=='policy':meta['audit']['discrete_root_search']['policy']='wrong'
            elif mutation=='signs':meta['audit']['discrete_root_search']['bracket_signs']=[-1,1]
            elif mutation=='endpoint':meta['audit']['discrete_root_search']['endpoint']=1.00001
            else:meta['audit']['discrete_root_search']['bracket'][1]=1.01
            write(path,meta)
            with self.subTest(mutation=mutation),self.assertRaises(ValueError):self.verify()
        write(path,original)

    def test_summary_and_audit_cannot_bind_other_adopted_fields(self):
        self.summary['lower']['fs']=.5;write(self.folder/'summary.json',self.summary)
        with self.assertRaisesRegex(ValueError,'Summary and adopted certificate'):self.verify()
        self.summary['lower']['fs']=1.;write(self.folder/'summary.json',self.summary)
        self.audit['sides']['lower']['mesh_id']='1'*64
        with self.assertRaisesRegex(ValueError,'Audit and producer'):self.verify()

    def test_independent_audit_must_track_own_evidence_and_input(self):
        tracked=self.audit['sides']['lower']['file_sha256'];tracked.pop(str(self.input))
        with self.assertRaisesRegex(ValueError,'omitted canonical input'):self.verify()

    def test_external_paths_and_missing_source_manifest_rejected(self):
        self.summary['environment']['source_sha256']['../outside.py']='0'*64
        write(self.folder/'summary.json',self.summary)
        with self.assertRaisesRegex(ValueError,'outside'):self.verify()
        self.summary['environment']['source_sha256']=dict(self.sources);write(self.folder/'summary.json',self.summary)
        (self.folder/'source_manifest.json').unlink()
        with self.assertRaises(OSError):self.verify()

    def test_wrapper_extension_and_explicit_checkpoint_formats(self):
        numerical={k:v for k,v in self.sources.items() if '/python/' in k}
        self.summary['environment']['source_sha256']=dict(numerical)
        self.summary['environment']['source_sha256_at_checkpoint']=dict(numerical)
        self.summary['wrapper_source_sha256']=dict(self.sources)
        write(self.folder/'source_manifest.json',{'source_sha256':numerical});write(self.folder/'summary.json',self.summary)
        self.assertTrue(self.verify()['explicit_producer_end_source_map'])


if __name__=='__main__':unittest.main()
