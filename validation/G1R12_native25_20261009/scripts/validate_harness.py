"""Fast native smoke for the evidence adapter; not one of the full 25 runs."""
import json, sys
import run_native25 as runner
from prepare_inputs import ROOT

def run():
    root=ROOT/'harness_validation';assert not root.exists()
    for sub in ['inputs','cases','results']:(root/sub).mkdir(parents=True,exist_ok=True)
    rec=json.loads((ROOT/'inputs/P08A.json').read_text(encoding='utf-8'))
    rec['settings'].update({'TARGET':32,'CYCLES':1,'QUADRATURE':2})
    rec['scope_notes']=['Evidence-adapter native smoke only; does not count as a catalog/full-profile execution.']
    (root/'inputs/P08A.json').write_text(json.dumps(rec,ensure_ascii=False,indent=2),encoding='utf-8')
    runner.ROOT=root
    result=runner.worker('P08A')
    assert result['fs_obtained'],result
    import check_saved_output as saved
    saved.ROOT=root
    saved.case_folder=lambda case:root/'cases'/case
    saved.evidence_root=lambda case:root
    out=saved.check('P08A');assert out['pass']
    print('HARNESS NATIVE EXPORT / INDEPENDENT AUDIT / SAVED OUTPUT PASS; not counted in 25-case ledger',flush=True)

if __name__=='__main__':run()
