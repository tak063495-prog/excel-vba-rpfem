from pathlib import Path
import json,hashlib,shutil,re
REPO=Path(__file__).resolve().parents[3]
ROOT=Path(__file__).resolve().parents[1]
SOURCE=REPO/'versions/G1R20_Fix8/delivery/RPFEM_20261019_G1R20_Fix8.xlsm'
DLL=SOURCE.parent/'rpfem_native_x64.dll'
SOURCE_SHA=(ROOT/'SOURCE_SHA256.txt').read_text().strip()  # Fix5 delivery workbook built on Winmax2
DLL_SHA='66da1a21be43cd2b3a3bfbaa9d13aad5fb1fd1960e1bf2507cc4e1cfa9ebe70b'
PRESET_MACRO='RPX_Speed2pctPreset'
CASE_IDS=[c['id'] for c in json.loads((REPO/'versions/G1R13/validation_runtime/data/cases.json').read_text(encoding='utf8'))]
# Fix5 operation test (user decision 2026-10-18 13:30 JST): Fix4 weak fast cases only; no P06N/P06A.
# Fix8 (2026-10-19): forced thin-layer smoke first (P02A_THIN: P02A input + extra_settings), then the
# Fix7 regression cases and two cheap extra cases. No P06N/P06A (user runs those on another PC).
ORDER=['P10A']
def base_case(case):return case.split('_')[0]
assert {base_case(c) for c in ORDER}<=set(CASE_IDS)
# Keys owned by the product preset; the G1R17 input values (TARGET 1536, CYCLES 3, GAP 0.01) must NOT be written.
# Case-specific keys (policy, seismic, diagnostics) are still taken from each input.
CASE_KEYS=['GRAVITY','TEST_INJECTION','HSD_CAPTURE','DIAGNOSTIC_MODE','ANALYSIS_POLICY','SEISMIC_KH','SEISMIC_DIR']
# Expected effective values after RPX_Speed2pctPreset (RPX_AccuracyPreset, then overrides), G1R20 Fix4 RPX_UI.bas.
PRESET_EXPECTED={'MODE':'Fs','TARGET':1000,'QUADRATURE':8,'FS_TOL':0.0001,'GAP':0.02,'ADAPT':1,'CYCLES':8,
 'RAW_LOWER_AUDIT':1,'C0_AUTO':0,'ADAPT_ACCURACY_CAP':24000,'ADAPT_ACCURACY_ROUNDS':8,'ADAPT_LOWER_TARGET':4000,
 'ADAPT_LOWER_FAN_ROUNDS':1,'ADAPT_CERT_QUADRATURE':8,'ACCURACY_DLL_RECOVERY':1,'ACCURACY_DLL_SECONDS':120,
 'ACCURACY_DLL_DUAL_SECONDS':20,'ACCURACY_DLL_CONE_SCALING':1,'HSD_SOLVE_SECONDS':300,'DIRECT_DLL':1,
 'ADAPT_PRECISION_FROM_START':1,'ADAPT_C0_WEDGE_AUTO':1,'ADAPT_BALANCED_LOWER_TARGET':0,'ADAPT_UPPER_FLOW_PASSES':0,
 'FS_TERMINAL_SCHEDULE':0,'ACCURACY_DLL_UPPER_RECENTER':0,'ACCURACY_DLL_DUAL_LINEAR_REPAIR':0,
 'ACCURACY_DLL_DUAL_REGULARIZATION':0,'DIRECT_DLL_SECONDS':5}
def preset_from_source():
    text=(REPO/'versions/G1R20_Fix8/src/vba/RPX_UI.bas').read_text(encoding='utf-8')
    def body(name):
        m=re.search(r'Public Sub '+name+r'\(\)(.*?)End Sub',text,re.S);return m.group(1)
    vals={}
    for name in ['RPX_AccuracyPreset','RPX_Speed2pctPreset']:
        for k,v in re.findall(r'AccuracySet "([A-Z0-9_]+)", ("[^"]*"|[-0-9.Ee]+)',body(name)):
            vals[k]=v.strip('"') if v.startswith('"') else float(v)
    return vals
def same(a,b):
    if isinstance(b,str):return str(a)==b
    try:return abs(float(a)-float(b))<=1e-12*max(1.,abs(float(b)))
    except (TypeError,ValueError):return False
src=preset_from_source()
assert set(src)==set(PRESET_EXPECTED) and all(same(src[k],v) for k,v in PRESET_EXPECTED.items()),('PRESET_TABLE_MISMATCH',src)
# Timeouts: 3x best prior analysis time, min 900 s, P06 <= 8 h. Sources:
# G1R17 DLL5 native26 (Triela 5900X, docs/benchmarks/G1R17_native26_20261013/data/RESULTS.csv);
# G1R19 native26 / #11 10-case rerun / preset2pct (Rico 5600X) where available.
G1R17={'D01A':1863,'D02A':2057,'D03N':2035,'D04A':322,'D05A':364,'D06A':321,'D07A':266,'D08A':348,'D09A':323,'P01A':265,'P01N':263,'P02A':334,'P02N':321,'P03A':253,'P03N':254,'P04A':336,'P04N':310,'P05A':284,'P05N':553,'P06A':6050,'P06N':6057,'P07A':548,'P07N':747,'P08A':1375,'P09A':1127,'P10A':990}
G1R19={'D01A':1787,'D02A':2020,'D03N':2183,'P06A':4851,'P06N':4491,'D08A':105,'P04A':209,'P04N':246,'P02N':205,'D07A':153,'P08A':951,'P09A':972,'P07A':993,'P07N':1015,'P05A':588}
# Revision 2 (2026-10-17 11:2x JST, user decision): speed measurement of all 26 must not be truncated.
# Limit = 3x G1R17 DLL5 time on the same machine (Triela 5900X), min 1200 s, P06 <= 8 h. G1R19 (Rico) times no longer used.
TIMEOUT_REVISION='fix8_optest_1'
# Fix4 Triela times: D05A failed 11 s (G1R17 364 s), P02A 234 s, P05A >=1200 s (limit hit), P10A 2262 s.
# Winmax2 (Ryzen 7 7840U) limits are generous: no speed claims against Triela.
FIX5_LIMITS={'P02A_THIN':5400,'D05A':3600,'P02A':3600,'P05A':7200,'P10A':10800,'P04A':3600,'D07A':3600}
def limit(case):
    return {'best_prior_seconds':G1R17[base_case(case)],'timeout_seconds':FIX5_LIMITS[case],'timeout_basis':'Fix5 operation test on Winmax2, generous fixed limit','timeout_revision':TIMEOUT_REVISION}
TIMEOUTS={c:limit(c) for c in ORDER}
OVERALL_SECONDS=8*3600
STOP_GRACE_SECONDS=600
PREPARING_LIMIT_SECONDS=1800
def sha(path):return hashlib.sha256(path.read_bytes()).hexdigest()
def write(path,data):
    path.parent.mkdir(parents=True,exist_ok=True)
    temp=path.with_suffix(path.suffix+'.tmp')
    temp.write_text(json.dumps(data,ensure_ascii=False,indent=2,default=str),encoding='utf8',newline='\n');temp.replace(path)
def prepare():
    assert sha(SOURCE)==SOURCE_SHA and sha(DLL)==DLL_SHA
    records=[]
    for case in ORDER:
        p=ROOT/'inputs'/f'{case}.json';r=json.loads(p.read_text(encoding='utf8'))
        assert r['case']==case and all(k in r['settings'] for k in CASE_KEYS)
        assert not r['water']['coupling']
        records.append({'case':case,'input_sha256':sha(p),'pore_pressure':bool(r['water']['coupling']),'area':r['source_case']['area'],'weight':r['total_weight'],**TIMEOUTS[case]})
    write(ROOT/'case_manifest.json',{'source_commit':(ROOT/'SOURCE_COMMIT.txt').read_text().strip(),'source':str(SOURCE.relative_to(REPO)),'source_sha256':SOURCE_SHA,'dll_sha256':DLL_SHA,
        'preset_macro':PRESET_MACRO,'preset_expected':PRESET_EXPECTED,'case_keys_from_input':CASE_KEYS,'run_order':ORDER,'cases':records,
        'overall_limit_seconds':OVERALL_SECONDS,'stop_grace_seconds':STOP_GRACE_SECONDS,
        'inputs':'docs/benchmarks/G1R17_native26_20261013/data/inputs (same geometry/material/water as previous 26-case runs); preset-owned settings replaced by product preset',
        'scope':'G1R20 Fix8 candidate operation test (2% standard preset, DLL) on Winmax2: forced thin-layer smoke P02A_THIN, regression P10A/D05A/P02A/P05A, extra P04A/D07A'})
    shutil.copyfile(DLL,ROOT/'workbooks/rpfem_native_x64.dll')
    print('PASS G1R20 Fix8 source/DLL hashes, preset table, inputs',flush=True)
if __name__=='__main__':prepare()
