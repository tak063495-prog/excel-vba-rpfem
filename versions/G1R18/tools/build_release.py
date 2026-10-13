"""Generate G1R18 full VBA modules, preserving the archived RayGate candidate."""
from pathlib import Path
import hashlib
import json
import shutil
import subprocess
import sys
from accuracy_fs_connection import connect

ROOT = Path(__file__).resolve().parents[1]
BASE = ROOT.parent / 'G1R17'
RAY = ROOT.parent / 'G1R18_RayGate'


def write(name, source):
    source = source.replace('\r\n', '\n').replace('\r', '\n')
    for directory, encoding in [('src/vba', 'utf-8'), ('import', 'cp932')]:
        path = ROOT / directory / (name + '.bas')
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(source.replace('\n', '\r\n').encode(encoding))


def main():
    geometry = ROOT / 'tools/accuracy_mesh_base.bas.txt'
    if not geometry.is_file():
        raise RuntimeError('Retained baseline accuracy geometry template is required')
    for path in sorted((BASE / 'src/vba').glob('*.bas')):
        write(path.stem, path.read_text(encoding='utf-8'))
    write('RPX_AccuracyMesh', geometry.read_text(encoding='utf-8'))
    subprocess.run([sys.executable, str(ROOT / 'tools/build_adapt_candidate.py')], check=True)
    subprocess.run([sys.executable, str(RAY / 'tools/build_ray_candidate.py')], check=True)
    for path in sorted((RAY / 'src/vba').glob('*.bas')):
        write(path.stem, path.read_text(encoding='utf-8'))
    adapt = (ROOT / 'src/vba/RPX_Adapt.bas').read_text(encoding='utf-8')
    old = '    If adaptAccuracyCap < adaptFinalTarget Or adaptAccuracyCap > 50000 Then'
    if adapt.count(old) != 1:
        raise ValueError('Expected accuracy cap guard')
    write('RPX_Adapt', adapt.replace(old, '    If adaptAccuracyCap = 0 Then adaptAccuracyCap = adaptFinalTarget\n'+old, 1))
    # Issue #8 hydraulics lives on the newer branch, not the retained G1R17 tree.
    for name in ['RPX_Seep', 'RPX_SeepUI']:
        write(name, (ROOT / f'tools/{name}_issue8.bas.txt').read_text(encoding='utf-8'))
    write('RPX_AccuracyNative', (ROOT / 'tools/accuracy_native.bas.txt').read_text(encoding='utf-8'))
    conic = (ROOT / 'src/vba/RPX_Conic.bas').read_text(encoding='utf-8')
    conic = conic.replace('    cost() As Double\nEnd Type', '    cost() As Double\n    normalizedCost() As Double\nEnd Type', 1)
    conic = conic.replace('    loadPhase As String\n    values() As Double', '    loadPhase As String\n    strengthFactor As Double\n    values() As Double', 1)
    conic = conic.replace('    XOriginalCompiled.loadPhase = RPX_LoadPhase', '    XOriginalCompiled.loadPhase = RPX_LoadPhase: XOriginalCompiled.strengthFactor = RStrengthFactor', 1)
    old_norm = '    For i = 0 To XN - 1: XQ(i) = XQ(i) / XQScale: Next i'
    if conic.count(old_norm) != 1:
        raise ValueError('Expected original objective normalization')
    conic = conic.replace(old_norm, old_norm+'\n    XOriginalCompiled.normalizedCost = XQ', 1)
    write('RPX_Conic', conic)
    cache = (ROOT / 'src/vba/RPX_LowerCache.bas').read_text(encoding='utf-8')
    old = '8# * (XN + XM + 1# + XEPtr(XM) + 1#)'
    if cache.count(old) != 1:
        raise ValueError('Expected original snapshot cache accounting')
    write('RPX_LowerCache', cache.replace(old, '8# * (2# * XN + XM + 1# + XEPtr(XM) + 1#)', 1))
    robust = (BASE / 'src/vba/RPX_Robust.bas').read_text(encoding='utf-8')
    robust = robust.replace('    Dim errorNumber As Long, errorSource As String, errorText As String, errorLine As Long', '    Dim errorNumber As Long, errorSource As String, errorText As String, errorLine As Long\n    Dim original As RPX_OriginalCompiledData', 1)
    old = 'Call RPX_SchurRelease:     RPX_HLURelease: RPX_ResetPreparedState'
    if robust.count(old) != 1:
        raise ValueError('Expected numeric-only robust reset')
    robust = robust.replace(old, '''    original = XOriginalCompiled
Call RPX_SchurRelease:     RPX_HLURelease: RPX_ResetPreparedState
    XOriginalCompiled = original''', 1)
    write('RPX_Robust', robust)
    lower = (BASE / 'src/vba/RPX_Lower.bas').read_text(encoding='utf-8')
    old = 'Public Sub RPX_RefreshLower()'
    if lower.count(old) != 1:
        raise ValueError('Expected original lower refresh')
    lower = lower.replace(old, old+'\n    If XOriginalCompiled.ready And Not XOriginalCompiled.affine Then XOriginalCompiled.strengthFactor = RStrengthFactor', 1)
    write('RPX_Lower', lower)
    fs_path = ROOT / 'src/vba/RPX_Fs.bas'
    write('RPX_Fs', connect(fs_path.read_text(encoding='utf-8')))
    # Prepared upper strength refresh also updates the literal objective snapshot.
    upper = (BASE / 'src/vba/RPX_Upper.bas').read_text(encoding='utf-8')
    old = '    For Each item In zeroPhiCosts: XQ(CLng(item)) = XQ(CLng(item)) * oldFactor / RStrengthFactor: Next item'
    if upper.count(old) != 1:
        raise ValueError('Expected the retained upper refresh implementation')
    new = old + '''
    If XOriginalCompiled.ready And Not XOriginalCompiled.affine Then
        For i = 0 To XN - 1: XOriginalCompiled.cost(i) = XQ(i) * XQScale: Next i
        XOriginalCompiled.normalizedCost = XQ
        XOriginalCompiled.strengthFactor = RStrengthFactor
    End If'''
    write('RPX_Upper', upper.replace(old, new, 1))
    ui = (BASE / 'src/vba/RPX_UI.bas').read_text(encoding='utf-8')
    old = '    RPX_EnsureUpperPilotSetting\n'
    if ui.count(old) < 1:
        raise ValueError('Expected settings initialization')
    ui = ui.replace(old, old+'    RPX_EnsureG1R18Settings\n', 1)
    ui += (ROOT / 'tools/accuracy_ui.bas.txt').read_text(encoding='utf-8')
    write('RPX_UI', ui)
    manifest = {'release': 'G1R18', 'base': 'G1R17', 'ray_candidate': 'G1R18_RayGate', 'modules': {}}
    for path in sorted((ROOT / 'src/vba').glob('*.bas')):
        utf = path.read_text(encoding='utf-8').replace('\r\n', '\n')
        imp = ROOT / 'import' / path.name
        if imp.read_bytes().decode('cp932').replace('\r\n', '\n') != utf:
            raise ValueError(f'Import mismatch: {path.name}')
        manifest['modules'][path.name] = {'utf8_sha256': hashlib.sha256(path.read_bytes()).hexdigest(),
                                        'cp932_sha256': hashlib.sha256(imp.read_bytes()).hexdigest()}
    (ROOT / 'SOURCE_MANIFEST.json').write_text(json.dumps(manifest, indent=2) + '\n', encoding='utf-8')
    print(f'Generated {len(manifest["modules"])} G1R18 full modules and CP932 imports.')


if __name__ == '__main__':
    main()
