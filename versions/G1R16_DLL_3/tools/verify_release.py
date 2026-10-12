import csv,difflib,gc,json,re,shutil,struct,zipfile,xml.etree.ElementTree as ET
from pathlib import Path
from oletools.olevba import VBA_Parser
from prepare import ROOT,REPO,SOURCE,sha,write
def canonical(text):
    text='\n'.join(s for s in text.splitlines() if not s.startswith('Attribute VB_'))
    result=[]
    for m in re.finditer(r'"(?:""|[^"])*"|\x27[^\n]*|[A-Za-z_][A-Za-z_0-9]*|(?:\d+\.\d*|\.\d+|\d+)(?:[eE][-+]?\d+)?[#@!%&]?|[^\s]',text):
        t=m.group()
        if t.startswith("'"):continue
        if t.startswith('"'):result.append(t);continue
        if re.fullmatch(r'(?:\d+\.\d*|\.\d+|\d+)(?:[eE][-+]?\d+)?[#@!%&]?',t) and ('.' in t or 'e' in t.lower() or t.endswith('#')):
            try:t='DOUBLE:'+struct.pack('<d',float(t.rstrip('#'))).hex()
            except ValueError:pass
        result.append(t.lower())
    return result
def macros(path):
    parser=VBA_Parser(str(path))
    try:return {name.rsplit('.',1)[0]:text for _,_,name,text in parser.extract_macros()}
    finally:parser.close()
def saved_cells(path):
    ns={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    with zipfile.ZipFile(path) as z:
        shared=[''.join(e.itertext()) for e in ET.fromstring(z.read('xl/sharedStrings.xml')).findall('s:si',ns)] if 'xl/sharedStrings.xml' in z.namelist() else []
        rels={e.attrib['Id']:e.attrib['Target'] for e in ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))};result={}
        for sheet in ET.fromstring(z.read('xl/workbook.xml')).findall('s:sheets/s:sheet',ns):
            target=rels[sheet.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']]
            target=target.lstrip('/') if target.startswith('/') else 'xl/'+target;cells={}
            for cell in ET.fromstring(z.read(target)).findall('.//s:sheetData/s:row/s:c',ns):
                v=cell.find('s:v',ns);typ=cell.get('t')
                if typ=='inlineStr':value=''.join(cell.find('s:is',ns).itertext())
                elif v is None or v.text is None:continue
                elif typ=='s':value=shared[int(v.text)]
                elif typ in ('str','e'):value=v.text
                else:value=float(v.text)
                cells[cell.attrib['r']]=value
            result[sheet.attrib['name']]=cells
        return result
def log(path):
    data=path.read_bytes();enc='utf-16' if data[:2] in [b'\xff\xfe',b'\xfe\xff'] else 'utf-8-sig'
    return list(csv.DictReader(data.decode(enc).splitlines(),delimiter='\t'))
def parts(text):return dict(x.split('=',1) for x in text.split(';') if '=' in x)
def main():
    folder=ROOT/'results/levee_dll3_final';status=json.loads((folder/'status.json').read_text(encoding='utf-8'));assert status['status']=='completed'
    assert status['dll_sha256']==sha(ROOT/'delivery/rpfem_native_x64.dll')
    dst=ROOT/'delivery/RPFEM_20261012_G1R16_DLL_3_Levee_Checked.xlsm';shutil.copyfile(folder/'DLL_3_Levee_Result.xlsm',dst)
    baseline=json.loads((ROOT/'results/baseline.json').read_text(encoding='utf-8'))
    assert all(sha(REPO/p)==h for p,h in baseline['protected'].items())
    old=macros(SOURCE);new=macros(dst);sources={p.stem:p.read_text(encoding='utf-8') for p in (ROOT/'src/vba').iterdir() if p.is_file()}
    assert set(new)==set(sources) and set(new)-set(old)=={'RPX_NativeBulk','RPX_NativeLU'}
    assert all(canonical(v)==canonical(new[n]) for n,v in sources.items())
    changed=[n for n in old if canonical(old[n])!=canonical(new[n])]
    assert set(changed)=={'RPX_Conic','RPX_HSDLu','RPX_Lower','RPX_LowerCache','RPX_Native','RPX_Solve','RPX_Upper'}
    assert all(canonical(old[n])==canonical(new[n]) for n in old if n.startswith('Sheet') or n=='ThisWorkbook')
    imports=json.loads((ROOT/'results/import_modules.json').read_text(encoding='utf-8'))['modules']
    for name in imports:
        data=(ROOT/'import'/f'{name}.bas').read_bytes();assert data.decode('cp932').replace('\r\n','\n')==sources[name];assert b'\r\n' in data
    diff=[]
    for name in sorted(changed+['RPX_NativeBulk','RPX_NativeLU']):
        a=(ROOT/'baseline_sources'/f'{name}.bas').read_text(encoding='utf-8').splitlines(keepends=True) if name in old else []
        diff.extend(difflib.unified_diff(a,sources[name].splitlines(keepends=True),fromfile='G1R16/'+name+'.bas' if a else '/dev/null',tofile='G1R16_DLL_3/'+name+'.bas'))
    (ROOT/'changes.diff').write_text(''.join(diff),encoding='utf-8')
    write(ROOT/'results/delivery_verification.json',{'protected_original_files':len(baseline['protected']),'protected_hashes_match':True,'original_modules':len(old),'delivery_modules':len(new),'changed_existing_modules':sorted(changed),'added_modules':['RPX_NativeBulk','RPX_NativeLU'],'unchanged_existing_modules':len(old)-len(changed),'document_modules_unchanged':True,'sources_match_saved_workbook':True,'comparison':'case-insensitive VBA identifiers; exact string literals; IEEE754-equivalent VBE numeric formatting; excludes comments/attributes','cp932_crlf_strict_roundtrip':True,'xlsm_sha256':sha(dst),'dll_sha256':sha(ROOT/'delivery/rpfem_native_x64.dll')})
    previous=log(REPO/'versions/G1R16/results/logs/20261012_095430_23840_analysis.tsv')
    reference=ROOT/'results/reference';reference.mkdir(parents=True,exist_ok=True)
    if not (reference/'dll2_saved_cells.json').exists():
        refbook=Path('C:/Users/link_/Desktop/G1R16_work_20261012/out/lev_DLL2.xlsm')
        refmacros=macros(refbook);assert set(refmacros)==set(old) and all(canonical(refmacros[n])==canonical(old[n]) for n in old)
        write(reference/'dll2_saved_cells.json',{'source_path':str(refbook),'source_sha256':sha(refbook),'same_macros_as_source':True,'cells':saved_cells(refbook)})
        reflog=Path('C:/Users/link_/Desktop/G1R16_work_20261012/out/RPFEM_logs/20261012_094013_24884_analysis.tsv')
        shutil.copyfile(reflog,reference/'vba_analysis.tsv')
        write(reference/'metadata.json',{'vba_log_source':str(reflog),'vba_log_sha256':sha(reflog),'dll2_log_sha256':sha(REPO/'versions/G1R16/results/logs/20261012_095430_23840_analysis.tsv')})
    fallback=log(reference/'vba_analysis.tsv')
    current=log(next((folder/'RPFEM_logs').glob('*analysis.tsv')))
    def starts(rows):return [(r['candidate'],r['mode'],r['strength_factor'],r['detail']) for r in rows if r['event']=='solve_start']
    keys=['status','last_iteration_index','objective','primal','dual','gap','message']
    def ends(rows):return [tuple(parts(r['detail']).get(k) for k in keys) for r in rows if r['event']=='solve_end']
    def audits(rows):return [r['detail'] for r in rows if r['event']=='audit']
    source_cells=saved_cells(SOURCE);before=json.loads((reference/'dll2_saved_cells.json').read_text(encoding='utf-8'))['cells'];after=saved_cells(dst)
    exact_inputs={n:before[n]==after[n] for n in ['要素定義','材料データ','設定','浸透入力','節点データ','要素データ']}
    assert all(source_cells[n]==after[n] for n in ['要素定義','材料データ','設定','浸透入力'])
    assert all(exact_inputs.values()),exact_inputs
    assert starts(current)==starts(previous) and ends(current)==ends(previous) and audits(current)==audits(previous)
    assert starts(current)==starts(fallback) and ends(current)==ends(fallback) and audits(current)==audits(fallback)
    fs_exact=all(struct.pack('<d',before['解析結果'][p])==struct.pack('<d',after['解析結果'][p]) for p in ['B4','C4']);assert fs_exact
    def duration(rows):return float(next(r['elapsed_s'] for r in rows if r['event']=='run_end'))
    seconds=duration(current);previous_seconds=duration(previous)
    write(ROOT/'results/comparison.json',{'native_seconds':seconds,'com_wall_seconds':status['analysis_seconds'],'previous_DLL2_seconds':previous_seconds,'VBA_existing_seconds':duration(fallback),'VBA_time_source':'results/reference/vba_analysis.tsv; actual existing G1R16 VBA-fallback run, not timed again','reduction_percent_vs_DLL2':100*(1-seconds/previous_seconds),'reduction_percent_vs_existing_VBA':100*(1-seconds/duration(fallback)),'lower':status['lower'],'upper':status['upper'],'Fs_bitwise_equal_saved_baseline':fs_exact,'solve_count':len(starts(current)),'solve_sequence_exact':True,'solve_results_exact_at_logged_precision':True,'VBA_solve_sequence_results_and_audits_exact':True,'audit_rows':len(audits(current)),'all_audit_details_exact':True,'same_saved_inputs_and_mesh':exact_inputs,'previous_phases':{parts(r['detail'])['phase']:float(parts(r['detail'])['seconds']) for r in previous if r['event']=='run_phase'},'native_phases':{parts(r['detail'])['phase']:float(parts(r['detail'])['seconds']) for r in current if r['event']=='run_phase'},'dll_sha256':status['dll_sha256'],'baseline_report_discrepancy':'G1R16 REPORT.md upper Fs 1.2344080288 is a transcription error: execution log and actual result workbook show 1.2343080287831916. GitHub delivery source retains G1R15 cached results (1.2244477358145585/1.2343080357010867) and its mesh; reference uses the actual G1R16 run. Original delivery and report preserved.'})
    print('RELEASE source/invariants/precision verification PASS',len(new),'modules;',len(starts(current)),'solves;',len(audits(current)),'audits;',seconds,'seconds',flush=True)
if __name__=='__main__':main()
