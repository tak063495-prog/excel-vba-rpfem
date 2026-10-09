from pathlib import Path
import json,difflib,hashlib
ROOT=Path(__file__).resolve().parents[1]
changed=[];unchanged=[];diff=[]
for p in sorted((ROOT/'src/vba').glob('*.bas')):
    bp=ROOT/'before/vba'/p.name;old=bp.read_text(encoding='utf-8') if bp.exists() else '';new=p.read_text(encoding='utf-8')
    if old==new:unchanged.append(p.name)
    else:
        changed.append(p.name);diff.extend(difflib.unified_diff(old.splitlines(True),new.splitlines(True),fromfile='G1R12/'+p.name,tofile='G1R13/'+p.name))
    data=('\r\n'.join(new.splitlines())+'\r\n').encode('cp932',errors='strict')
    assert data.decode('cp932').splitlines()==new.splitlines();(ROOT/'import'/p.name).write_bytes(data)
assert changed==['RPX_Audit.bas','RPX_Fs.bas','RPX_HSDCandidate.bas','RPX_LoadStepMesh.bas','RPX_Mesh.bas','RPX_WorkPrecision.bas'],changed
assert len(unchanged)==42
for name in ['RPX_LDL.bas','RPX_Solve.bas','RPX_NT.bas','RPX_NTPrecision.bas','RPX_HSDPrecision.bas','RPX_Upper.bas','RPX_Lower.bas','RPX_Certificate.bas','RPX_Definition.bas']:
    assert name in unchanged,name
s=(ROOT/'src/vba/RPX_WorkPrecision.bas').read_text(encoding='utf-8')
assert 'If XQ(i) <> 0# Then Exit Function' in s and 'If XC(i).offset(j) <> 0# Then Exit Function' in s
assert 'trial = candidate' in s and s.index('If coneError > 0.0000001 Then Exit Function')<s.index('candidate = trial')
assert 'candidate.field = candidate.x' in (ROOT/'src/vba/RPX_HSDCandidate.bas').read_text(encoding='utf-8')
for token in ['UPPER_PHYSICAL_AUDIT_FAILED','LOWER_PHYSICAL_AUDIT_FAILED','ANALYSIS_CANCELLED']:
    assert token not in (ROOT/'src/vba/RPX_Fs.bas').read_text(encoding='utf-8').split('Private Function FsNumericRecoverable')[1].split('End Function')[0]
(ROOT/'changes.diff').write_text(''.join(diff),encoding='utf-8')
record={'pass':True,'changed_and_added':changed,'unchanged_existing_standard_modules':unchanged,'numeric_kernels_constitutive_and_tolerances_unchanged':True,'import_encoding':'CP932 strict lossless / CRLF','scope':'static/source assertions; no Excel execution claim'}
(ROOT/'results/source_regression.json').write_text(json.dumps(record,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps({'pass':True,'unchanged':42,'changed':changed}))
