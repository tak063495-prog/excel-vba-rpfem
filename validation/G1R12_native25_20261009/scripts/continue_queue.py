"""Transfer not-started jobs to eight independent Excel workers.

The existing process pool has no live resizing API. Reserve only jobs that have
not entered Excel; its workers will return a preparation assertion rather than
start duplicate FEM runs. Keep all active native solves untouched. Real native
runs use a separate output root, overlaid by the evidence collector. Reservations
are scheduler control, never native analysis results or case failures.
"""
from pathlib import Path
import json, shutil, sys, concurrent.futures, multiprocessing, time
from datetime import datetime,timezone
from prepare_inputs import ROOT
import run_native25 as runner

QUEUE=ROOT/'queued_continuation'
ACTIVE={'P06A','P06N','P08A','P09A','D01A'}

def actual(case):
    runner.ROOT=QUEUE
    return runner.worker(case)

def reserve():
    assert not (ROOT/'queue_transfer.json').exists()
    for name in ['inputs','cases','results']:(QUEUE/name).mkdir(parents=True,exist_ok=True)
    ids=[p.stem for p in sorted((ROOT/'inputs').glob('*.json')) if p.stem not in ACTIVE]
    assert len(ids)==20
    rec={'created_utc':datetime.now(timezone.utc).isoformat(),'active_native_cases_untouched':sorted(ACTIVE),
         'workers':8,'output_root':'queued_continuation','cases':[],
         'scope':'Reserve only not-started queued jobs; existing Excel analyses remain running; no duplicate FEM executions.'}
    for case in ids:
        folder=ROOT/'cases'/case;folder.mkdir(exist_ok=True)
        assert not list((folder/'RPFEM_logs').glob('*_analysis.tsv')),(case,'native run already started')
        status=folder/'status.json'
        if status.exists():
            r=json.loads(status.read_text(encoding='utf-8'));assert r['status']=='NOT_RUN',r
        marker=folder/f'G1R12_{case}_Full1536_Q8.xlsm'
        with marker.open('xb') as f:f.write(b'SCHEDULER RESERVATION ONLY - actual workbook under queued_continuation/cases\r\n')
        shutil.copy2(ROOT/'inputs'/f'{case}.json',QUEUE/'inputs'/f'{case}.json')
        rec['cases'].append(case)
    runner.write(ROOT/'queue_transfer.json',rec)
    print('Reserved 20 not-started jobs; native solves untouched',flush=True)
    return ids

def run():
    ids=reserve()
    priority=['P10A','D02A','D03N']
    ordered=priority+[c for c in ids if c not in priority]
    with concurrent.futures.ProcessPoolExecutor(max_workers=8,mp_context=multiprocessing.get_context('spawn')) as pool:
        pending={pool.submit(actual,c):c for c in ordered}
        while pending:
            done,_=concurrent.futures.wait(pending,timeout=10,return_when=concurrent.futures.FIRST_COMPLETED)
            for future in done:
                case=pending.pop(future)
                try:future.result()
                except Exception as e:runner.write(QUEUE/'results'/f'{case}.json',{'case':case,'status':'WORKER_ERROR','error':repr(e)})
    print('CONTINUATION20_FINISHED',flush=True)

if __name__=='__main__':run()
