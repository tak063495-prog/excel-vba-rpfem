"""Run source/provenance regression controls; this never starts Excel or a solver."""
from pathlib import Path
import argparse,json,subprocess,sys,time

ROOT=Path(__file__).resolve().parents[1]
REPO=ROOT.parents[1]
SUITES=[
 ('fix1_ui','versions/G1R20_Fix1/tests','UI_test_contract.py'),
 ('fix1_water','versions/G1R20_Fix1/tests','test_contract_water.py'),
 ('fix1_scope','versions/G1R20_Fix1/tests','test_release_scope.py'),
 ('dual','versions/G1R19_DualRepair_20261016/tests','test_contract.py'),
 ('terminal','versions/G1R19_TerminalRoot_20261016/tests','test_contract.py'),
 ('upper','versions/G1R19_UpperRecovery_20261016/tests','test_contract.py'),
 ('water','versions/G1R19_WaterCommon_20261016/tests','test_*.py'),
 ('state','versions/G1R19_StateReset_20261016/tests','test_contract.py'),
 ('flow','versions/G1R19_FlowMesh_20261016/tests','test_contract.py'),
 ('collector','validation/G1R19_python26_20261016','test_collector.py'),
 ('provenance','validation/G1R19_python26_20261016','test_evidence_provenance.py'),
]

def main():
    parser=argparse.ArgumentParser();parser.add_argument('--output',type=Path,required=True)
    args=parser.parse_args();out=args.output.resolve()
    if out.exists():raise FileExistsError('Use a new result file')
    rows=[]
    for label,path,pattern in SUITES:
        start=time.perf_counter()
        command=[sys.executable,'-X','utf8','-m','unittest','discover','-s',path,'-p',pattern]
        p=subprocess.run(command,cwd=REPO,capture_output=True,text=True,encoding='utf8')
        rows.append(dict(suite=label,command=command,exit_code=p.returncode,
                         seconds=time.perf_counter()-start,stdout=p.stdout,stderr=p.stderr))
        print(label+': '+('PASS' if p.returncode==0 else 'FAIL'),flush=True)
    passed=all(row['exit_code']==0 for row in rows)
    out.parent.mkdir(parents=True,exist_ok=True)
    out.write_text(json.dumps(dict(status='PASS' if passed else 'FAIL',solver_calls=0,excel_runs=0,
                                  scope='Portable source and rejection contracts; not fresh numerical solves',suites=rows),
                              indent=2,ensure_ascii=False,allow_nan=False)+'\n',encoding='utf8')
    if not passed:raise SystemExit(1)

if __name__=='__main__':main()
