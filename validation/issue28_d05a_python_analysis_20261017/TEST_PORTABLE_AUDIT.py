"""Meaningful tamper/structure/roundoff controls; no optimization calls."""
from __future__ import annotations
import argparse
import copy
import json
from pathlib import Path
import sys
from unittest import mock
sys.dont_write_bytecode=True
from run_fixed_point_experiment import ROOT,core,sha,write
from portable_program_audit import compare_program,float_roundoff_bound,pack_program,program_digest
from VERIFY_MICRO import load_references,bind_recorded_program_digest,validate_run_summary,validate_run_entry

def main():
    parser=argparse.ArgumentParser();parser.add_argument('--output',type=Path,required=True);args=parser.parse_args()
    if args.output.exists():raise FileExistsError(args.output)
    refs,_=load_references();reference=refs['P02A'][0];records=[]
    def accepted(name,call):
        call();records.append({'name':name,'expected':'ACCEPT','status':'PASS'})
    def rejected(name,call):
        try:call()
        except (AssertionError,ValueError):records.append({'name':name,'expected':'REJECT','status':'PASS'});return
        raise AssertionError('Negative control was accepted: '+name)
    accepted('identical_original_program',lambda:compare_program(reference,reference))
    near=copy.deepcopy(reference);near[-1][0]=core.np.nextafter(near[-1][0],core.np.inf)
    accepted('one_ulp_float_coefficient_with_structure_unchanged',lambda:compare_program(near,reference))
    changed=copy.deepcopy(reference);changed[-1][0]+=2*float_roundoff_bound(reference[-1])
    rejected('float_coefficient_exceeds_explicit_32eps_bound',lambda:compare_program(changed,reference))
    structure=copy.deepcopy(reference);structure[2].indices[0]=(structure[2].indices[0]+1)%structure[2].shape[0]
    rejected('changed_CSC_row_index',lambda:compare_program(structure,reference))
    dtype=list(copy.deepcopy(reference));dtype[0]=dtype[0].astype(core.np.float32)
    rejected('float32_dtype_substitution',lambda:compare_program(tuple(dtype),reference))
    nonfinite=copy.deepcopy(reference);nonfinite[2].data[0]=core.np.nan
    rejected('NaN_coefficient',lambda:compare_program(nonfinite,reference))
    dims=list(copy.deepcopy(reference));dims[4]=dims[4].copy();dims[4][0]=2;dims[4][1]=4
    rejected('changed_integer_cone_dimensions',lambda:compare_program(tuple(dims),reference))
    accepted('frozen_record_digest_binds_exact_reference',lambda:bind_recorded_program_digest(program_digest(reference),near,reference))
    accepted('fresh_current_host_digest_binds_exact_current_program',lambda:bind_recorded_program_digest(program_digest(near),near,reference))
    rejected('unknown_third_host_digest_without_coefficients',lambda:bind_recorded_program_digest('0'*64,near,reference))
    summary=json.loads((ROOT/'results/fixed_point_02/summary.json').read_text(encoding='utf-8'))
    accepted('six_unique_expected_route_names',lambda:validate_run_summary(summary))
    missing=copy.deepcopy(summary);missing['runs'].pop()
    rejected('missing_route_even_if_optimizer_count_claims_six',lambda:validate_run_summary(missing))
    duplicate=copy.deepcopy(summary);duplicate['runs'][-1]=copy.deepcopy(duplicate['runs'][0])
    rejected('duplicate_route_replaces_required_case',lambda:validate_run_summary(duplicate))
    header=copy.deepcopy(summary);header['q']=4
    rejected('summary_quadrature_mismatch',lambda:validate_run_summary(header))
    entry=summary['runs'][0];meta=json.loads((ROOT/'results/fixed_point_02'/entry['name']/'result.json').read_text(encoding='utf-8'))
    meta['route']='affine'
    rejected('entry_case_route_metadata_disagrees',lambda:validate_run_entry(entry,meta))
    target=ROOT/refs['P02A'][1]['path'];original_sha=sha(target.read_bytes());original_read=Path.read_bytes
    def tampered_read(path):
        raw=original_read(path)
        return raw[:-1]+bytes([raw[-1]^1])if path.resolve()==target.resolve()else raw
    with mock.patch.object(Path,'read_bytes',tampered_read):
        rejected('reference_NPZ_byte_tamper_fails_actual_loader_SHA',load_references)
    assert sha(target.read_bytes())==original_sha
    write(args.output,{'status':'PASS','optimizer_calls':0,'tests':records,'test_count':len(records),
        'original_reference_file_unchanged':True,'auditor_sha256':sha(Path(__file__).read_bytes()),
        'comparator_sha256':sha((ROOT/'portable_program_audit.py').read_bytes())})
    print({'status':'PASS','tests':len(records),'optimizer_calls':0})
if __name__=='__main__':main()
