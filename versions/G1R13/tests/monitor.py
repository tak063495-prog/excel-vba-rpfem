from pathlib import Path
import csv,json,sys
ROOT=Path(__file__).resolve().parents[1]
for folder in [ROOT/'single_point/D03N',ROOT/'native_retry/D03N',ROOT/'native_retry/D07A']:
    st=folder/'status.json';r=json.loads(st.read_text(encoding='utf-8')) if st.exists() else {}
    print(folder.relative_to(ROOT),r.get('status'),r.get('excel_pid'),r.get('lower_fs'),r.get('upper_fs'),r.get('error',''))
    logs=sorted((folder/'RPFEM_logs').glob('*.tsv'))
    if not logs:continue
    try:
        rows=list(csv.reader(logs[-1].open(encoding='utf-16'),delimiter='\t'))
        for row in [v for v in rows if len(v)>10 and v[6] in ['solve_end','fs_numerical_unknown','load_step_mesh_repair','upper_work_normalize','p4_final_bracket','fatal']][-3:]:print('  ',row[2:7],row[-1][:450])
        starts=[v for v in rows if len(v)>10 and v[6]=='solve_start'];print('  latest solve',starts[-1][2:7] if starts else None)
    except Exception as exc:print(type(exc).__name__)
    stage=folder/'stage.txt'
    if stage.exists():print('  ',stage.read_text(encoding='cp932',errors='replace').splitlines()[-1])
