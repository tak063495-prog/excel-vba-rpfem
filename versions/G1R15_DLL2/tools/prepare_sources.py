raise SystemExit('DLL2: use prepare_imports.py. This older generator resets edited sources and must not be used.')
from pathlib import Path
import shutil, json, hashlib, re
ROOT=Path(__file__).resolve().parents[1]
target=ROOT/'src/vba'; target.mkdir(parents=True,exist_ok=True)
base=ROOT/'baseline_sources'
for p in base.glob('*.bas'): shutil.copyfile(p,target/p.name)
for name,prefix in [('RPX_LDL','RPX'),('RPX_SchurLDL','RPXC')]:
    p=target/(name+'.bas'); text=p.read_text(encoding='utf-8')
    # Local flags preserve the original Boolean pivot-failure contract.
    anchor='    Dim useCache As Boolean\n'
    assert text.count(anchor)==1
    text=text.replace(anchor,anchor+'    Dim nativeAccepted As Boolean\n',1)
    anchor='    For k = 0 To orderN - 1\n        If k Mod 4096 = 0 Then\n'
    assert text.count(anchor)>=2
    hook=f'    If RPX_NativeNumeric(orderN, ap, ai, av, colPtr, rowInd, reachPtr, reachNode, writeAt, factorVal, diagonal, scatter, {prefix}_PivotMinimum, {prefix}_PivotFailure, nativeAccepted) Then\n        {prefix}_Numeric = nativeAccepted: Exit Function\n    End If\n'
    text=text.replace(anchor,hook+anchor,1)
    start=text.index(f'Public Sub {prefix}_Backsolve(')
    anchor='    ReDim answer(0 To orderN - 1)\n'
    at=text.index(anchor,start)+len(anchor)
    text=text[:at]+'    If RPX_NativeBacksolve(orderN, colPtr, rowInd, factorVal, diagonal, rhs, answer) Then Exit Sub\n'+text[at:]
    if prefix=='RPX':
        start=text.index('Public Function RPX_DDNumeric(')
        at=text.index('    Dim bh As Double, bl As Double\n',start)+len('    Dim bh As Double, bl As Double\n')
        text=text[:at]+'    Dim nativeAccepted As Boolean\n'+text[at:]
        anchor='    RPX_PivotMinimum = 1E+300: RPX_PivotFailure = -1\n'
        at=text.index(anchor,start)+len(anchor)
        text=text[:at]+('    If RPX_NativeDDNumeric(orderN, ap, ai, av, colPtr, rowInd, reachPtr, reachNode, writeAt, ddFactor, ddDiagonal, ddScatter, RPX_PivotMinimum, RPX_PivotFailure, nativeAccepted) Then\n'
                       '        ddReady = nativeAccepted: RPX_DDNumeric = nativeAccepted: Exit Function\n    End If\n')+text[at:]
        start=text.index('Public Sub RPX_DDBacksolve('); at=text.index('    ReDim answer(0 To orderN - 1)\n',start)+len('    ReDim answer(0 To orderN - 1)\n')
        text=text[:at]+'    If RPX_NativeDDBacksolve(orderN, colPtr, rowInd, ddFactor, ddDiagonal, rhs, answer) Then Exit Sub\n'+text[at:]
    p.write_text(text,encoding='utf-8')
# Existing G1R15 preview locals px/py shadow its coordinate helper functions.
# Only helper names/calls change; local point coordinates and all equations stay.
p=target/'RPX_SeepUI.bas'; text=p.read_text(encoding='utf-8')
for old,new in [('px','PreviewX'),('py','PreviewY')]:
    text=re.sub(r'\b'+old+r'\(',new+'(',text,flags=re.I)
    start=text.index('Private Function '+new+'('); end=text.index('End Function',start)
    text=text[:start]+re.sub(r'\b'+old+r'\s*=',new+' =',text[start:end],flags=re.I)+text[end:]
p.write_text(text,encoding='utf-8')
imports=ROOT/'import'; imports.mkdir(exist_ok=True)
for name in ['RPX_Native','RPX_LDL','RPX_SchurLDL','RPX_SeepUI']:
    text=(target/(name+'.bas')).read_text(encoding='utf-8')
    data=text.replace('\n','\r\n').encode('cp932',errors='strict')
    assert data.decode('cp932').replace('\r\n','\n')==text
    (imports/(name+'.bas')).write_bytes(data)
hashes={str(p.relative_to(ROOT.parent.parent)):hashlib.sha256(p.read_bytes()).hexdigest() for p in (ROOT.parent/'G1R15').rglob('*') if p.is_file()}
(ROOT/'results/protected_g1r15_hashes.json').write_text(json.dumps(hashes,indent=2),encoding='utf-8')
print('Sources prepared; CP932 round-trip exact; G1R15 hashes saved:',len(hashes))
