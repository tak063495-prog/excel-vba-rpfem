"""Serial inherited UI and source-provenance regressions on final Fix4 sources."""
from pathlib import Path
import argparse, hashlib, json, subprocess, sys, time
ROOT=Path(__file__).resolve().parents[1]
PLAN=[('presets','UI_run_excel.py',[]),('setup_control','run_ui_setup_control_excel.py',[]),('seep_ui','run_seep_ui_fix3_excel.py',[]),('water_source','run_water_source_fix3_excel.py',[]),('water_restart','run_water_source_process_restart.py',[])]
PLAN += [('saved_'+scenario,'run_saved_formula_excel.py',['--scenario',scenario]) for scenario in ('edit','missingproof','openchanged')]

def selected_plan(only=None):
    """Select named groups while preserving their established serial order."""
    requested=set(only) if only is not None else {label for label,_,_ in PLAN}
    unknown=requested-{label for label,_,_ in PLAN}
    if unknown:raise ValueError('Unknown inherited groups '+repr(sorted(unknown)))
    return [item for item in PLAN if item[0] in requested]

def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
def exited(record):
    stages=record.get('sessions',record.get('stages',[]))
    return bool(record.get('owned_excel_exited') or (stages and all(s.get('owned_excel_exited') for s in stages)))
def main():
    p=argparse.ArgumentParser();p.add_argument('--execute',action='store_true');p.add_argument('--output',type=Path,required=True)
    p.add_argument('--only',nargs='+',choices=[label for label,_,_ in PLAN],help='Run only these groups in the standard serial order; omitted groups are recorded as skipped')
    a=p.parse_args()
    if not a.execute:raise SystemExit('Requires serial owned Excel --execute')
    out=a.output.resolve();out.mkdir(parents=True,exist_ok=False)
    plan=selected_plan(a.only)
    sources=[*sorted((ROOT/'src/vba').glob('*.bas')),ROOT/'src/workbook/ThisWorkbook.code.txt']
    before={str(s):sha(s) for s in sources}
    rec=dict(status='RUNNING',scope='Inherited positive tests; synthetic display sentinels are not numerical certificates',source_sha256=before,sessions=[],runs=[],requested_groups=a.only,selected_groups=[label for label,_,_ in plan],skipped_groups=[label for label,_,_ in PLAN if label not in {item[0] for item in plan}],full_suite=len(plan)==len(PLAN))
    def save():(out/'summary.json').write_text(json.dumps(rec,ensure_ascii=False,indent=2)+'\n',encoding='utf8')
    save()
    for label,script,extra in plan:
        folder=out/label
        cmd=[sys.executable,'-X','utf8',str(ROOT/'tests'/script),'--execute','--output',str(folder),*extra]
        start=time.perf_counter();result=subprocess.run(cmd,capture_output=True,text=True,encoding='utf8')
        (out/(label+'.log')).write_text(result.stdout+result.stderr,encoding='utf8')
        row=dict(name=label,command=cmd,exit_code=result.returncode,seconds=time.perf_counter()-start)
        summary=folder/'summary.json'
        if summary.exists():
            value=json.loads(summary.read_text(encoding='utf8'));row.update(record=str(summary),record_sha256=sha(summary),status=value.get('status',value.get('state')),owned_excel_exited=exited(value))
        rec['runs'].append(row);rec['sessions'].append(row);save()
        print(label+': '+str(row),flush=True)
        if result.returncode or row.get('status')!='PASS' or not row.get('owned_excel_exited'):
            rec['status']='FAIL';save();raise SystemExit(1)
    rec['sources_unchanged']=all(sha(Path(s))==h for s,h in before.items())
    rec['status']='PASS' if rec['sources_unchanged'] else 'FAIL';save()
    if rec['status']!='PASS':raise SystemExit(1)
if __name__=='__main__':main()
