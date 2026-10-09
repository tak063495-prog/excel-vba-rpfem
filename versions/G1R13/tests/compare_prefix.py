from pathlib import Path
import csv,json
ROOT=Path(__file__).resolve().parents[1]
old=Path('C:/Users/link_/Desktop/RPFEM_G1R12_Native25_20261009/queued_continuation/cases/D03N/RPFEM_logs/20261009_081113_23396_analysis.tsv')
new=next((ROOT/'native_retry/D03N/RPFEM_logs').glob('*.tsv'))
def starts(p):
    return [r[2:7] for r in csv.reader(p.open(encoding='utf-16'),delimiter='\t') if len(r)>10 and r[6]=='solve_start']
a,b=starts(old),starts(new);same=0
for x,y in zip(a,b):
    if x!=y:break
    same+=1
rec={'scope':'ordered native solve-start ID/stage/mode/Fs labels only; not numerical iterate/field bit equality','current_starts':len(b),'matching_prefix':same,'first_difference':{'old':a[same],'new':b[same]} if same<min(len(a),len(b)) else None}
(ROOT/'results/native_solve_prefix.json').write_text(json.dumps(rec,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps(rec,ensure_ascii=False))
