from pathlib import Path
import json,hashlib,re,struct,difflib
from oletools.olevba import VBA_Parser
ROOT=Path(__file__).resolve().parents[1]
def macros(path):
    parser=VBA_Parser(str(path))
    try:return {name.rsplit('.',1)[0]:text for _,_,name,text in parser.extract_macros()}
    finally:parser.close()
def canonical(text):
    text='\n'.join(s for s in text.splitlines() if not s.startswith('Attribute VB_'))
    tokens=[]
    for match in re.finditer(r'"(?:""|[^"])*"|\x27[^\n]*|[A-Za-z_][A-Za-z_0-9]*|(?:\d+\.\d*|\.\d+|\d+)(?:[eE][-+]?\d+)?[#@!%&]?|[^\s]',text):
        t=match.group()
        if t.startswith("'"):continue
        if t.startswith('"'):tokens.append(t);continue
        if re.fullmatch(r'(?:\d+\.\d*|\.\d+|\d+)(?:[eE][-+]?\d+)?[#@!%&]?',t) and ('.' in t or 'e' in t.lower() or t.endswith('#')):
            try:t='DOUBLE:'+struct.pack('<d',float(t.rstrip('#'))).hex()
            except ValueError:pass
        tokens.append(t.lower())
    return tokens
protected=json.loads((ROOT/'results/protected_g1r15_hashes.json').read_text(encoding='utf-8'))
broken=[p for p,h in protected.items() if hashlib.sha256((ROOT.parent.parent/p).read_bytes()).hexdigest()!=h]
assert not broken,broken
original=macros(ROOT.parent/'G1R15/examples/RPFEM_20261011_G1R15_Levee_Example.xlsm')
delivered=macros(ROOT/'delivery/RPFEM_20261012_G1R15_DLL2_Levee_Checked.xlsm')
assert set(delivered)-set(original)=={'RPX_Native'}
assert not set(original)-set(delivered)
changed=[n for n in original if canonical(original[n])!=canonical(delivered[n])]
expected=['RPX_LDL','RPX_SchurLDL','RPX_SeepUI','RPX_Order','RPX_ColumnOrder']
assert set(changed)==set(expected),changed
for name in expected+['RPX_Native']:
    expected_source=(ROOT/'src/vba'/f'{name}.bas').read_text(encoding='utf-8')
    assert canonical(expected_source)==canonical(delivered[name]),name
    data=(ROOT/'import'/f'{name}.bas').read_bytes()
    assert data.decode('cp932').replace('\r\n','\n')==expected_source
    assert b'\r\n' in data
diff=[]
for name in expected:
    a=(ROOT/'baseline_sources'/f'{name}.bas').read_text(encoding='utf-8').splitlines(keepends=True)
    b=(ROOT/'src/vba'/f'{name}.bas').read_text(encoding='utf-8').splitlines(keepends=True)
    diff.extend(difflib.unified_diff(a,b,fromfile='G1R15/'+name+'.bas',tofile='G1R15_DLL/'+name+'.bas'))
new=(ROOT/'src/vba/RPX_Native.bas').read_text(encoding='utf-8').splitlines(keepends=True)
diff.extend(difflib.unified_diff([],new,fromfile='/dev/null',tofile='G1R15_DLL/RPX_Native.bas'))
(ROOT/'changes.diff').write_text(''.join(diff),encoding='utf-8')
result={'protected_g1r15_files':len(protected),'all_protected_hashes_match':True,
        'original_modules':len(original),'delivered_modules':len(delivered),
        'added_modules':['RPX_Native'],'changed_modules':changed,
        'unchanged_modules':len(original)-len(changed),
        'comparison':'VBA identifiers case-insensitive, string literals exact, equivalent VBE decimal/scientific Double literals normalized; comments/attributes excluded',
        'all_document_and_class_modules_unchanged':all(n in expected or canonical(original[n])==canonical(delivered[n]) for n in original if not n.startswith('RPX_')),
        'cp932_crlf_lossless':True,
        'xlsm_sha256':hashlib.sha256((ROOT/'delivery/RPFEM_20261012_G1R15_DLL2_Levee_Checked.xlsm').read_bytes()).hexdigest(),
        'dll_sha256':hashlib.sha256((ROOT/'delivery/rpfem_native_x64.dll').read_bytes()).hexdigest()}
(ROOT/'results/delivery_verification.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps(result,ensure_ascii=False,indent=2))
