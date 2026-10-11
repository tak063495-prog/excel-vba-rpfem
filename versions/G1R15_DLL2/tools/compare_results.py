from pathlib import Path
import csv,json,hashlib,xml.etree.ElementTree as ET,zipfile,struct
ROOT=Path(__file__).resolve().parents[1]
NS={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
def saved_cells(path):
    with zipfile.ZipFile(path) as z:
        shared=[]
        if 'xl/sharedStrings.xml' in z.namelist():
            shared=[''.join(e.itertext()) for e in ET.fromstring(z.read('xl/sharedStrings.xml')).findall('s:si',NS)]
        rels={e.attrib['Id']:e.attrib['Target'] for e in ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
        result={}
        for sheet in ET.fromstring(z.read('xl/workbook.xml')).findall('s:sheets/s:sheet',NS):
            rid=sheet.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']; target=rels[rid]
            target=target.lstrip('/') if target.startswith('/') else 'xl/'+target
            cells={}
            for cell in ET.fromstring(z.read(target)).findall('.//s:sheetData/s:row/s:c',NS):
                v=cell.find('s:v',NS); typ=cell.get('t')
                if typ=='inlineStr': value=''.join(cell.find('s:is',NS).itertext())
                elif v is None or v.text is None:continue
                elif typ=='s':value=shared[int(v.text)]
                elif typ in ('str','e'):value=v.text
                else:value=float(v.text)
                cells[cell.attrib['r']]=value
            result[sheet.attrib['name']]=cells
        return result
def read_log(path):return list(csv.DictReader(path.open(encoding='utf-16'),delimiter='\t'))
def parts(text):return dict(s.split('=',1) for s in text.split(';') if '=' in s)
original=read_log(ROOT.parent/'G1R15/results/RPFEM_logs/20261011_211428_26004_analysis.tsv')
previous_folder=ROOT.parent/'G1R15_DLL/results/dll_run3_final'
previous=read_log(next((previous_folder/'RPFEM_logs').glob('*analysis.tsv')))
previous_time=float(next(r['elapsed_s'] for r in previous if r['event']=='run_end'))
previous_book=saved_cells(previous_folder/'RPFEM_G1R15_DLL_Levee.xlsm')
reference_time=float(next(r['elapsed_s'] for r in original if r['event']=='run_end'))
reference_starts=[(r['candidate'],r['mode'],r['strength_factor'],r['detail']) for r in original if r['event']=='solve_start']
reference_phases={parts(r['detail'])['phase']:float(parts(r['detail'])['seconds']) for r in original if r['event']=='run_phase'}
old_book=Path('C:/Users/link_/Desktop/G1R15_work_20261011_1250/iso_L_cpl_eq_opt_2115.xlsm')
old=saved_cells(old_book) if old_book.exists() else None
report={'reference_seconds':reference_time,'reference_kind':'existing VBA run on the same PC; not rerun at user request',
        'previous_dll_seconds':previous_time,
        'source_commit':'d6c4d0903582a325d7c0368cf22dbfb80ee64db6','runs':[]}
for folder in sorted((ROOT/'results').glob('dll_run*')):
    result=json.loads((folder/'status.json').read_text(encoding='utf-8'))
    if result.get('status')!='completed':continue
    logs=list((folder/'RPFEM_logs').glob('*analysis.tsv')); assert len(logs)==1
    rows=read_log(logs[0]); terminal=next(r for r in rows if r['event']=='run_end')
    duration=float(terminal['elapsed_s'])
    starts=[(r['candidate'],r['mode'],r['strength_factor'],r['detail']) for r in rows if r['event']=='solve_start']
    phases={parts(r['detail'])['phase']:float(parts(r['detail'])['seconds']) for r in rows if r['event']=='run_phase'}
    current=saved_cells(folder/'RPFEM_G1R15_DLL_Levee.xlsm')
    comparisons={}
    if old:
        for name in ['要素定義','材料データ','設定','浸透入力']:
            # All numeric/material input cells; generated display/status cells in
            # hydraulic input are outputs and are listed separately if different.
            a=old.get(name,{}); b=current.get(name,{})
            differences=[(k,a.get(k),b.get(k)) for k in set(a)|set(b) if a.get(k)!=b.get(k)]
            comparisons[name]=differences
        for name in ['節点データ','要素データ','応力','速度','応力結果','速度結果']:
            if name in old and name in current:comparisons[name]=old[name]==current[name]
    audit_rows=[r for r in rows if r['event']=='audit']
    reference_audits=[r['detail'] for r in original if r['event']=='audit']
    signature_keys=['status','last_iteration_index','objective','primal','dual','gap','message']
    def signatures(records):
        return [tuple(parts(r['detail']).get(k) for k in signature_keys) for r in records if r['event']=='solve_end']
    audit_rejections=[r for r in rows if 'audit' in r['event'] and ('status=FAIL' in r['detail'] or 'accepted=0' in r['detail'])]
    item={'tag':folder.name,'analysis_seconds':duration,'com_wall_seconds':result['analysis_seconds'],
          'speedup_vs_previous_dll':previous_time/duration,
          'reduction_percent_vs_previous_dll':100*(1-duration/previous_time),
          'speedup':reference_time/duration,'reduction_percent':100*(1-duration/reference_time),
          'lower':result['lower'],'upper':result['upper'],
          'gap_percent':200*(result['upper']-result['lower'])/(result['upper']+result['lower']),
          'solve_count':len(starts),'solve_start_labels_and_mesh_exact':starts==reference_starts,
          'terminal':terminal['detail'],'native_calls':result['native_end'],
          'audit_rows':len(audit_rows),'audit_rejections':audit_rejections,
          'audit_details_exact_at_logged_precision':[r['detail'] for r in audit_rows]==reference_audits,
          'solve_end_status_iteration_objective_residual_gap_exact_at_logged_precision':signatures(rows)==signatures(original),
          'reference_phases':reference_phases,'dll_phases':phases,'saved_book_comparisons':comparisons,
          'previous_dll_phases':{parts(r['detail'])['phase']:float(parts(r['detail'])['seconds']) for r in previous if r['event']=='run_phase'},
          'previous_dll_audits_exact_at_logged_precision':[r['detail'] for r in audit_rows]==[r['detail'] for r in previous if r['event']=='audit'],
          'previous_dll_solve_ends_exact_at_logged_precision':signatures(rows)==signatures(previous),
          'previous_dll_saved_inputs_and_mesh_exact':{name:current[name]==previous_book[name] for name in ['要素定義','材料データ','設定','浸透入力','節点データ','要素データ']},
          'log_sha256':hashlib.sha256(logs[0].read_bytes()).hexdigest()}
    assert starts==reference_starts, 'Solver sequence or mesh changed'
    assert 'status=OPTIMAL' in terminal['detail'], 'Not an optimal completion'
    assert not audit_rejections, 'Audit rejection in benchmark'
    report['runs'].append(item)
(ROOT/'results/comparison.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps([{k:r[k] for k in ['tag','analysis_seconds','speedup','reduction_percent','lower','upper','gap_percent','solve_count','solve_start_labels_and_mesh_exact','native_calls']} for r in report['runs']],ensure_ascii=False,indent=2))
