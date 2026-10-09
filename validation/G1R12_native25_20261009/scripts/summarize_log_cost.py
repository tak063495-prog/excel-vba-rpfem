"""Read completed native solve events; never infer whole-case success from them."""
from pathlib import Path
import csv, json, collections, math
from datetime import datetime, timezone
from prepare_inputs import ROOT
from case_locations import case_folder

def detail(s):
    return dict(x.split('=',1) for x in s.split(';') if '=' in x)

def collect():
    costs=[];cases=[]
    for p in sorted((ROOT/'inputs').glob('*.json')):
        case=p.stem;folder=case_folder(case)
        starts={};ends=[];events=collections.Counter();peak=0;hopt_run={};phases_run={}
        for log in sorted((folder/'RPFEM_logs').glob('*_analysis.tsv')):
            with log.open(encoding='utf-16',errors='replace',newline='') as f:
                for r in csv.DictReader(f,delimiter='\t'):
                    event=r.get('event','');events[event]+=1
                    try:peak=max(peak,int(r.get('private_bytes','0')))
                    except (ValueError,TypeError):pass
                    if event=='solve_start':starts[r['solve_id']]=r
                    elif event=='solve_end':ends.append(r)
                    elif event=='hopt0':
                        d=detail(r.get('detail',''))
                        if d.get('scope')=='run':
                            hopt_run[d['bucket']]={k:float(d[k]) for k in ['inclusive_seconds','exclusive_seconds','calls']}
                    elif event=='run_phase':
                        d=detail(r.get('detail',''))
                        phases_run[d['phase']]={'seconds':float(d['seconds']),'calls':int(d['calls'])}
        for e in ends:
            s=starts.get(e['solve_id'],{});sd=detail(s.get('detail',''));ed=detail(e.get('detail',''))
            row={'case':case,'solve_id':e['solve_id'],'candidate':e['candidate'],'side':e['mode'],'fs':e['strength_factor'],
                 'q':sd.get('subdivisions'),'elements':sd.get('elements'),'nodes':sd.get('nodes'),
                 'status':ed.get('status'),'seconds':float(ed['wall_seconds']) if ed.get('wall_seconds') else None,
                 'iterations':ed.get('last_iteration_index'),'objective':ed.get('objective'),
                 'primal':ed.get('primal'),'dual':ed.get('dual'),'gap':ed.get('gap'),'message':ed.get('message','')}
            costs.append(row)
        own=[r for r in costs if r['case']==case]
        byside={k:sum(r['seconds'] or 0 for r in own if r['side']==k) for k in ['upper','lower']}
        cases.append({'case':case,'finished_solves':len(own),'native_solve_seconds_by_side':byside,
                      'completed_solve_status_counts':dict(collections.Counter(r['status'] for r in own)),
                      'event_counts':dict(events),'peak_private_bytes_observed':peak,
                      'native_run_hopt_buckets':hopt_run,'native_run_phases':phases_run,
                      'profiler_note':'Different instrumentation views overlap; do not add phase and hopt times together. Empty run summaries indicate a still-running or unreported case.',
                      'scope':'completed solve events only; not a whole-case success judgment'})
    stamp=datetime.now(timezone.utc).isoformat()
    (ROOT/'results/solve_costs.json').write_text(json.dumps({'updated_utc':stamp,'cases':cases,'solves':costs},ensure_ascii=False,indent=2),encoding='utf-8')
    keys=['case','solve_id','candidate','side','fs','q','elements','nodes','status','seconds','iterations','objective','primal','dual','gap','message']
    with (ROOT/'results/solve_costs.csv').open('w',encoding='utf-8-sig',newline='') as f:
        w=csv.DictWriter(f,fieldnames=keys);w.writeheader();w.writerows(costs)
    print(json.dumps({'cases':len(cases),'completed_solve_events':len(costs),'path':str(ROOT/'results/solve_costs.json')},ensure_ascii=False))

if __name__=='__main__':collect()
