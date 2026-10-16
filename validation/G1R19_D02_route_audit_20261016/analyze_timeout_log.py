"""Read the published D02A timeout evidence; no Excel or numerical solve."""
from pathlib import Path
from datetime import datetime,timezone
import csv,collections,hashlib,json

ROOT=Path(__file__).resolve().parent
REPO=ROOT.parents[1]
SOURCE=REPO/'validation/G1R19_preset2pct_20261016/D02A'


def detail(row):return dict(x.split('=',1) for x in row['detail'].split(';') if '=' in x)
def sha(path):return hashlib.sha256(path.read_bytes()).hexdigest()


def main():
    tree=json.loads((ROOT/'github_main_tree_latest.json').read_text(encoding='utf8'))
    paths={item['path']:item for item in tree['tree']}
    files={}
    for path in sorted(SOURCE.iterdir()):
        if not path.is_file():continue
        data=path.read_bytes();key=path.relative_to(REPO).as_posix()
        blob=hashlib.sha1(b'blob '+str(len(data)).encode()+b'\0'+data).hexdigest()
        if key in paths and blob!=paths[key]['sha']:raise ValueError('Published Git blob differs: '+key)
        files[key]={'bytes':len(data),'sha256':hashlib.sha256(data).hexdigest(),'git_blob_sha1':blob}
    if (SOURCE/'analysis.tsv').read_bytes().decode('utf16')!=(SOURCE/'analysis_utf8.tsv').read_bytes().decode('utf8'):
        raise ValueError('UTF8 conversion differs from the original UTF16 log')
    rows=list(csv.DictReader((SOURCE/'analysis_utf8.tsv').open(encoding='utf8'),delimiter='\t'))
    ends=[row for row in rows if row['event']=='solve_end']
    phases=[]
    for phase in dict.fromkeys(row['candidate'] for row in ends):
        selected=[row for row in ends if row['candidate']==phase]
        evaluation=next(row for row in rows if row['event']=='adapt_eval' and row['candidate']==phase)
        settings=detail(evaluation)
        phases.append({'phase':phase,'start_seconds':float(evaluation['elapsed_s']),
                       'q':int(settings['quadrature']),'fs_tolerance':float(settings['fs_tolerance']),
                       'elements':int(settings['elements']),'solve_count':len(selected),
                       **{mode+'_calls':sum(row['mode']==mode for row in selected) for mode in ('lower','upper')},
                       **{mode+'_solve_wall_seconds':sum(float(detail(row)['wall_seconds']) for row in selected if row['mode']==mode)
                          for mode in ('lower','upper')}})
    direct=collections.Counter((row['mode'],detail(row)['accepted'],detail(row)['kind'])
                               for row in rows if row['event']=='direct_dll_end')
    case=json.loads((SOURCE/'results_D02A.json').read_text(encoding='utf8'))
    timeout=json.loads((SOURCE/'results_D02A_timeout.json').read_text(encoding='utf8'))
    cooperative=json.loads((SOURCE/'cooperative_stop.json').read_text(encoding='utf8'))
    worker=json.loads((SOURCE/'results_D02A_worker_error.json').read_text(encoding='utf8'))
    started=datetime.fromisoformat(case['started_utc']).timestamp()
    killed=datetime.fromisoformat(timeout['utc']).timestamp()
    result={'scope':'Published D02A TIMEOUT original log only; not the new overlay cold result',
        'github_commit':tree['sha'],'github_pr':'https://github.com/tak063495-prog/excel-vba-rpfem/pull/20',
        'files':files,'utf16_original_and_utf8_conversion_text_equal':True,
        'workbook_sha256':case['source_sha256'],'dll_sha256':case['dll_sha256'],'input_sha256':case['input_sha256'],
        'actual_settings':case['actual_settings'],'stored_json_status_is_stale':case['status']=='RUNNING',
        'actual_terminal_state':'TIMEOUT_THEN_OWNED_EXCEL_KILLED_AFTER_GRACE','worker_exit_code':worker['exit_code'],
        'configured_limit_seconds':timeout['timeout_s'],'cooperative_stop_after_worker_start_seconds':cooperative['time']-started,
        'forced_kill_after_worker_start_seconds':killed-started,'grace_after_cooperative_stop_seconds':killed-cooperative['time'],
        'log_rows':len(rows),'last_log_event':rows[-1],'run_end_present':any(row['event']=='run_end' for row in rows),
        'solve_start_count':sum(row['event']=='solve_start' for row in rows),'solve_end_count':len(ends),
        'status_counts':dict(collections.Counter(detail(row)['status'] for row in ends)),
        'event_counts':dict(collections.Counter(row['event'] for row in rows)),
        'phases':phases,'direct_routes':[{'mode':key[0],'accepted':key[1]=='True','kind':key[2],'count':count} for key,count in direct.items()],
        'adapt_cycles':[{'elapsed':float(row['elapsed_s']),**detail(row)} for row in rows if row['event']=='adapt_cycle'],
        'accuracy_rounds':[{'elapsed':float(row['elapsed_s']),**detail(row)} for row in rows if row['event']=='accuracy_round'],
        'initial_band':[{'elapsed':float(row['elapsed_s']),**detail(row)} for row in rows if row['event']=='adapt_band'],
        'last_five_solves':ends[-5:],
        'interpretation':'Lower about1.5 is already inside the independent fresh Python certified interval [1.4999745740800001,1.500142126501501]; do not treat its lack of increase as the accuracy defect. The weak upper and repeated near-critical lower searches dominate the remaining gap and cost.',
        'unverified':'No log events after6052.925531s. The precise unresponsive instruction and reason cooperative cancellation failed are not established by this evidence.'}
    (ROOT/'D02_LOG_STATS.json').write_text(json.dumps(result,ensure_ascii=False,indent=2,allow_nan=False)+'\n',encoding='utf8')
    with (ROOT/'D02_PHASE_TIMINGS.csv').open('w',encoding='utf-8-sig',newline='') as stream:
        writer=csv.DictWriter(stream,fieldnames=list(phases[0]));writer.writeheader();writer.writerows(phases)
    print(json.dumps({'published_commit':result['github_commit'],'solves':len(ends),'phases':phases,
                      'cooperative_s':result['cooperative_stop_after_worker_start_seconds'],
                      'killed_s':result['forced_kill_after_worker_start_seconds']},ensure_ascii=False))


if __name__=='__main__':main()
