from pathlib import Path
import json,collections,time,sys
from datetime import datetime,timezone
ROOT=Path(__file__).resolve().parent
from case_locations import case_folder
from grade_results import grade
def status():
    rows=[]
    for inp in sorted((ROOT/'inputs').glob('*.json')):
        path=case_folder(inp.stem)/'status.json'
        if not path.exists():continue
        try:r=grade(json.loads(path.read_text(encoding='utf-8')))
        except (FileNotFoundError,json.JSONDecodeError):continue
        row={k:r[k] for k in ['case','status','excel_pid','lower_fs','upper_fs','solve_seconds','seconds','error'] if k in r}
        if r['status']=='RUNNING':
            from datetime import datetime
            row['elapsed_s']=round((datetime.now(timezone.utc)-datetime.fromisoformat(r['run_started_utc'])).total_seconds())
            stage=path.parent/'stage.txt'
            if stage.exists():
                with stage.open('rb') as f:
                    f.seek(max(0,stage.stat().st_size-2097152));raw=f.read()
                lines=raw.decode('cp932',errors='replace').splitlines()
                row['last_stage']=lines[-1][:160] if lines else ''
                trials=[(i,line) for i,line in enumerate(lines) if line.startswith(('# lower F=', '# upper F='))]
                if trials:
                    trial_index,trial_line=trials[-1]
                    row['native_stage_trial']=trial_line[:160]
                    lines=lines[trial_index:]
                iterations=[line for line in lines if line.startswith('# HSD ') and ' residual=' in line]
                if iterations:row['last_iteration']=iterations[-1][:160]
            logs=sorted((path.parent/'RPFEM_logs').glob('*_analysis.tsv'))
            if logs:
                log=logs[-1]
                with log.open('rb') as f:
                    start=max(2,log.stat().st_size-1048576);start-=start%2;f.seek(start);raw=f.read()
                events=[]
                for line in raw.decode('utf-16-le',errors='replace').splitlines():
                    cols=line.split('\t')
                    if len(cols)>=11 and cols[6] in ['solve_start','solve_end','candidate_begin','run_end','outcome']:
                        events.append({'solve':cols[2],'candidate':cols[3],'side':cols[4],'fs':cols[5],'event':cols[6],'detail':cols[10][:180]})
                row['recent_events']=events[-3:]
        rows.append(row)
    counts=dict(collections.Counter(r['status'] for r in rows))
    compact=[]
    for r in rows:
        if 'recent_events' in r:
            events=r.pop('recent_events')
            if events:r['last_solve_event']=events[-1]
        compact.append(r)
    if '--compact' in sys.argv:
        print(json.dumps(counts,ensure_ascii=False))
        for r in compact:
            if '--running-only' in sys.argv and r['status']!='RUNNING':continue
            ev=r.get('last_solve_event',{})
            print(r['case'],r['status'],str(round(r.get('elapsed_s',r.get('seconds',0))/60,1))+'min',
                  ev.get('candidate',''),r.get('native_stage_trial',ev.get('side','')+' Fs='+str(ev.get('fs',''))),
                  r.get('last_stage',r.get('error',''))[:95],r.get('last_iteration',''))
    else:print(json.dumps({'counts':counts,'rows':compact},ensure_ascii=False,indent=2))
if __name__=='__main__':status()
