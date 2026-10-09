from pathlib import Path
import csv,json
ROOT=Path(__file__).resolve().parents[1]
for case in ['D07A','D03N']:
    folder=ROOT/'native_retry'/case;st=folder/'status.json'
    if not st.exists():continue
    r=json.loads(st.read_text(encoding='utf-8'))
    if r['status']!='RUNNING':
        print(case,r['status'],r.get('lower_fs'),r.get('upper_fs'));continue
    logs=sorted((folder/'RPFEM_logs').glob('*.tsv'))
    if logs:
        rows=list(csv.reader(logs[-1].open(encoding='utf-16'),delimiter='\t'))
        starts=[row for row in rows if len(row)>10 and row[6]=='solve_start']
        if starts:print(case,'solve',starts[-1][2],starts[-1][3:6])
    p=folder/'stage.txt'
    if p.exists():print(p.read_text(encoding='cp932',errors='replace').splitlines()[-1])
