from pathlib import Path
import json,hashlib
R=Path(__file__).resolve().parents[1]
changed=[p.name for p in (R/'src').glob('*.bas') if p.read_bytes()!=(R/'original'/p.name).read_bytes()]
assert sorted(changed)==['RPX_Adapt.bas','RPX_Fs.bas','RPX_Refine.bas']
def proc(s,header,end):
    a=s.index(header);return s[a:s.index(end,a)+len(end)]
old=(R/'original/RPX_Adapt.bas').read_text(encoding='utf-8');new=(R/'src/RPX_Adapt.bas').read_text(encoding='utf-8')
for name in ['ComputeShearWeight','PlaceBandCore','RPX_BuildBandMesh','RestoreFinalBounds']:
    assert proc(old,'Private Sub '+name+'(','End Sub')==proc(new,'Private Sub '+name+'(','End Sub'),name
oldfs=(R/'original/RPX_Fs.bas').read_text(encoding='utf-8');newfs=(R/'src/RPX_Fs.bas').read_text(encoding='utf-8')
for name in ['FsValue','TryFsValue','TryFsTrialPoint','BracketUsable','SaveBracket','UpperSeedUsable','RPX_Bound']:
    kind='Sub' if name=='SaveBracket' else 'Function'
    a=('Public' if name=='RPX_Bound' else 'Private')+' '+kind+' '+name+'('
    assert proc(oldfs,a,'End '+kind)==proc(newfs,a,'End '+kind),name
assert oldfs[oldfs.index('    For i = 1 To 40'):oldfs.index('    If i > 40')]==newfs[newfs.index('    For i = 1 To 40'):newfs.index('    If i > 40')]
for name in changed:
    text=(R/'src'/name).read_text(encoding='utf-8');raw=(R/'import_cp932'/name).read_bytes()
    assert raw.decode('cp932').splitlines()==text.splitlines()
    assert b'\n' not in raw.replace(b'\r\n',b'')
identity=json.loads((R/'source_identity.json').read_text());assert hashlib.sha256(Path(identity['path']).read_bytes()).hexdigest()==identity['sha256']
out={'changed_modules':changed,'legacy_band_and_weight_sources_unchanged':True,'physical_assembly_Davis_audit_NT_HSD_LDL_order_precision_kernel_sources_unchanged':True,'Fs_trial_classification_secant_termination_target_tolerances_unchanged':True,'strict_cp932_crlf_roundtrip':'PASS','source_sha_intact':True}
(R/'results/static_invariants.json').write_text(json.dumps(out,indent=2),encoding='utf-8');print(out)
