"""Build a fresh TEST-ONLY 63-module source set; never opens Excel.

Traps stop before conic compilation / seep solving. Private cache sentinels are
accessible only through functions appended to that module in this test copy.
Do not use this instrumented source to produce a numerical certificate.
"""
from pathlib import Path
import argparse
import hashlib
import importlib.util
import json

ROOT = Path(__file__).resolve().parents[1]
REPO = ROOT.parents[1]
spec = importlib.util.spec_from_file_location('state_reset_overlay', ROOT/'make_overlay_manifest.py')
overlay = importlib.util.module_from_spec(spec)
spec.loader.exec_module(overlay)

CONIC = '''
' TEST-ONLY: appended in a disposable microtest book, never product source.
Public Function SRProbeConicEntry() As String
    Dim candidate As RPX_HsdCandidateState, accepted As Boolean
    Dim oldReady As Boolean, oldModel As String, oldReason As String, message As String
    oldReady = originalGuideReady: oldModel = originalGuideModel: oldReason = OriginalFarkasReason
    On Error GoTo Failed
    originalGuideReady = True: originalGuideModel = "TEST_ONLY_READY_SENTINEL"
    candidate.kind = "TEST_ONLY_REJECTED_SCOPE"
    accepted = RPX_OriginalFarkasGate(candidate)
    If accepted Or originalGuideReady Or Len(originalGuideModel) <> 0 Then err.Raise 5, , "GUIDE_ENTRY_DID_NOT_CLEAR_SENTINEL"
    SRProbeConicEntry = "PASS conic_entry;ready_sentinel_cleared=True;solver_calls=0"
    GoTo Cleanup
Failed:
    message = err.description: SRProbeConicEntry = "FAIL conic_entry " & message
Cleanup:
    originalGuideReady = oldReady: originalGuideModel = oldModel: OriginalFarkasReason = oldReason
End Function
'''

SOLVE = '''
' TEST-ONLY private state access; not a mesh/field/root witness.
Public Sub SRProbePrimeSolveSentinels()
    SRProbeGeneration = hStructGen
    originalKReady = True: warmValid = True: hCoefReady = True: hEvalReady = True
    ReDim originalK(0 To 0): originalK(0) = 123#
    ReDim warmXX(0 To 0): ReDim warmXY(0 To 0): ReDim warmSlack(0 To 0): ReDim warmDual(0 To 0)
    ReDim xx(0 To 0): ReDim xy(0 To 0): ReDim XSlack(0 To 0): ReDim XDual(0 To 0)
    ReDim dx(0 To 0): ReDim dy(0 To 0): ReDim ds(0 To 0): ReDim dz(0 To 0)
End Sub
Private Function SRProbeAbsent(ByRef values() As Double) As Boolean
    Dim high As Long
    On Error Resume Next
    high = UBound(values): SRProbeAbsent = (err.number <> 0): err.Clear
    On Error GoTo 0
End Function
Public Function SRProbeSolveSentinelsCleared() As Boolean
    If originalKReady Or warmValid Or hCoefReady Or hEvalReady Then Exit Function
    If hStructGen <= SRProbeGeneration Then Exit Function
    If Not SRProbeAbsent(originalK) Or Not SRProbeAbsent(warmXX) Or Not SRProbeAbsent(warmXY) Then Exit Function
    If Not SRProbeAbsent(warmSlack) Or Not SRProbeAbsent(warmDual) Then Exit Function
    If Not SRProbeAbsent(xx) Or Not SRProbeAbsent(xy) Or Not SRProbeAbsent(XSlack) Or Not SRProbeAbsent(XDual) Then Exit Function
    If Not SRProbeAbsent(dx) Or Not SRProbeAbsent(dy) Or Not SRProbeAbsent(ds) Or Not SRProbeAbsent(dz) Then Exit Function
    SRProbeSolveSentinelsCleared = True
End Function
Public Function SRProbeKrylovEntry() As String
    Dim bdd(0 To 1) As RPX_DD, answer(0 To 1) As RPX_DD, best As Double, accepted As Boolean, used As Boolean
    Dim oldN As Long, oldM As Long, oldB As Long, oldS As Long, oldCancel As Boolean
    Dim number As Long, message As String
    If Not RPX_NativeAvailable() Then SRProbeKrylovEntry = "FAIL krylov_entry PREREQUISITE_NATIVE_DLL": Exit Function
    oldN = XN: oldM = XM: oldB = XB: oldS = XS: oldCancel = XCancel
    XN = 1: XM = 1: XB = 1: XS = 1: XCancel = True
    SRProbeKrylovTrapActive = True: SRProbeKrylovTrapHits = 0
    On Error Resume Next
    used = HNativeKrylov(bdd, answer, best, accepted)
    number = err.number: message = err.description: err.Clear
    On Error GoTo 0
    SRProbeKrylovTrapActive = False
    XN = oldN: XM = oldM: XB = oldB: XS = oldS: XCancel = oldCancel
    If number = 18 And InStr(message, "ANALYSIS_CANCELLED") > 0 And SRProbeKrylovTrapHits = 0 Then
        SRProbeKrylovEntry = "PASS krylov_entry;cancel_before_native_begin=True;solver_calls=0"
    Else
        SRProbeKrylovEntry = "FAIL krylov_entry;error=" & number & ";message=" & message
    End If
End Function
'''

PROBE = '''Attribute VB_Name = "StateResetMicroProbe"
Option Explicit
' TEST-ONLY disposable book. The sentinels are not physical proof state.
Public SRProbeGeneration As Long
Public SRProbeCompileTrapActive As Boolean, SRProbeSeepTrapActive As Boolean
Public SRProbeKrylovTrapActive As Boolean, SRProbeKrylovTrapHits As Long
Public SRProbeCompileHits As Long, SRProbeSeepHits As Long
Public SRProbeCompileCleared As Boolean, SRProbeSeepDefaultsOK As Boolean
Public Sub SRProbeKrylovAfterEntryTrap()
    If Not SRProbeKrylovTrapActive Then Exit Sub
    SRProbeKrylovTrapActive = False: SRProbeKrylovTrapHits = SRProbeKrylovTrapHits + 1
    err.Raise 18, "StateResetMicroProbe", "TEST_ONLY_KRYLOV_ENTRY_GUARD_WAS_SKIPPED"
End Sub
Public Sub SRProbeCompileTrap()
    If Not SRProbeCompileTrapActive Then Exit Sub
    SRProbeCompileTrapActive = False: SRProbeCompileHits = SRProbeCompileHits + 1
    SRProbeCompileCleared = SRProbeSolveSentinelsCleared()
    err.Raise 18, "StateResetMicroProbe", "TEST_ONLY_STOP_BEFORE_COMPILE"
End Sub
Public Sub SRProbeSeepTrap()
    If Not SRProbeSeepTrapActive Then Exit Sub
    SRProbeSeepTrapActive = False: SRProbeSeepHits = SRProbeSeepHits + 1
    SRProbeSeepDefaultsOK = (SKrMin = 0.0001 And SBeta = 20# And SRelax = 0.25 And STol = 0.000001 And SMaxIt = 300 And SGammaW = 9.81 And SResTol = 0.000001)
    err.Raise 18, "StateResetMicroProbe", "TEST_ONLY_STOP_BEFORE_SEEP_SOLVE"
End Sub
Public Function SRProbeDirectEntry(ByVal mode As String, ByVal factor As Double) As String
    Dim number As Long, message As String, value As Double
    On Error GoTo Failed
    If RPX_Busy Or rn < 3 Or re < 1 Or RR <> 0 Or (mode <> "lower" And mode <> "upper") Then err.Raise 5, , "DIRECT_ENTRY_REQUIRES_IDLE_SOIL_MESH"
    If Not RPX_DirectEnabled() Then err.Raise 5, , "DIRECT_ENTRY_REQUIRES_Q8_FS_DIRECT_PROFILE"
    XCancel = False: RProgramKind = ""
    Call SRProbePrimeSolveSentinels
    SRProbeCompileHits = 0: SRProbeCompileCleared = False: SRProbeCompileTrapActive = True
    On Error Resume Next
    value = RPX_Bound(mode, factor)
    number = err.number: message = err.description: err.Clear
    On Error GoTo Failed
    If number <> 18 Or InStr(message, "TEST_ONLY_STOP_BEFORE_COMPILE") = 0 Then err.Raise 5, , "DIRECT_ENTRY_DID_NOT_HIT_COMPILE_TRAP " & number & " " & message
    If SRProbeCompileHits <> 1 Or Not SRProbeCompileCleared Then err.Raise 5, , "DIRECT_ENTRY_DID_NOT_RESET_PRIVATE_SENTINELS"
    SRProbeDirectEntry = "PASS direct_entry_" & mode & ";private_sentinels_cleared=True;compile_calls=0;solver_calls=0"
    GoTo Cleanup
Failed:
    message = err.description: SRProbeDirectEntry = "FAIL direct_entry_" & mode & " " & message
Cleanup:
    On Error Resume Next
    SRProbeCompileTrapActive = False: Call RPX_ResetPreparedState
    RProgramKind = "": XCancel = False: RPX_ResultCurrent = False: RPX_MeshCurrent = False
    On Error GoTo 0
End Function
Public Function SRProbeSeepEntry() As String
    Dim result As String, number As Long, message As String
    SKrMin = -777#: SBeta = -777#: SRelax = -777#: STol = -777#: SMaxIt = -777: SGammaW = -777#: SResTol = -777#
    SRProbeSeepHits = 0: SRProbeSeepDefaultsOK = False: SRProbeSeepTrapActive = True
    On Error Resume Next
    result = RPX_SeepTestRelax(0.25)
    number = err.number: message = err.description: err.Clear
    On Error GoTo 0
    SRProbeSeepTrapActive = False
    If number = 18 And InStr(message, "TEST_ONLY_STOP_BEFORE_SEEP_SOLVE") > 0 And SRProbeSeepHits = 1 And SRProbeSeepDefaultsOK Then
        SRProbeSeepEntry = "PASS seep_entry;defaults_sentinels_replaced=True;seep_solver_calls=0"
    Else
        SRProbeSeepEntry = "FAIL seep_entry;error=" & number & ";message=" & message & ";defaults=" & SRProbeSeepDefaultsOK
    End If
    Call RPX_SeepDefaults
    RPX_ResultCurrent = False: RPX_MeshCurrent = False
End Function
'''


def replace_once(text, before, after):
    if text.count(before) != 1:
        raise ValueError('Test instrumentation anchor is not unique: '+before)
    return text.replace(before, after, 1)


def replace_in_procedure(text, declaration, before, after):
    if text.count(declaration) != 1:
        raise ValueError('Test procedure anchor is not unique: '+declaration)
    first = text.index(declaration)
    last = text.index('\nEnd Function', first)
    return text[:first]+replace_once(text[first:last],before,after)+text[last:]


def generate(output, negative_control=False):
    output = Path(output).resolve()
    if output.exists():
        raise FileExistsError('Refusing to replace probe source directory: '+str(output))
    # Validate all four product changes before making any test-only copies.
    overlay.manifest_data()
    files = {p.name:p for p in (REPO/'versions/G1R19/src/vba').glob('*.bas')}
    files.update({p.name:p for p in (REPO/'versions/G1R19_2pct/src/vba').glob('*.bas')})
    if not negative_control:
        files.update({p.name:p for p in (ROOT/'src/vba').glob('*.bas')})
    if len(files) != 62:
        raise ValueError('Expected exactly 62 numerical modules')
    texts = {name:overlay.normalize(path.read_text(encoding='utf8')) for name,path in files.items()}
    texts['RPX_Conic.bas'] = replace_once(texts['RPX_Conic.bas'],
        'Public Sub RPX_Compile()\n', 'Public Sub RPX_Compile()\n    Call SRProbeCompileTrap\n') + CONIC
    # Stop a negative control before HEnsureDD too: that helper can itself
    # reject the deliberately synthetic dimensions before later native setup.
    entry='    '+('' if negative_control else 'Call ')+'HContext: HEnsureDD XN + XM\n'
    texts['RPX_Solve.bas'] = replace_in_procedure(texts['RPX_Solve.bas'],
        'Private Function HNativeKrylov(',entry,
        entry.replace(': HEnsureDD', ': Call SRProbeKrylovAfterEntryTrap: HEnsureDD')) + SOLVE
    texts['RPX_Seep.bas'] = replace_once(texts['RPX_Seep.bas'],
        'Public Sub RPX_SeepSolve()\n', 'Public Sub RPX_SeepSolve()\n    Call SRProbeSeepTrap\n')
    texts['StateResetMicroProbe.bas'] = PROBE
    records = {}
    for name, text in texts.items():
        for directory, encoding in [('src/vba','utf8'), ('import','cp932')]:
            path = output/directory/name
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_bytes(text.replace('\n','\r\n').encode(encoding))
        records[name] = dict(source_sha256=overlay.sha(output/'src/vba'/name),
                             import_sha256=overlay.sha(output/'import'/name),
                             original_source=(str(files[name].relative_to(REPO)) if name in files else None),
                             original_sha256=(overlay.sha(files[name]) if name in files else None))
    data = dict(schema='rpfem.statereset.test_only.v1', test_only=True,
                negative_control=bool(negative_control), numerical_evidence_allowed=False,
                status='SOURCE_ONLY; EXCEL_NOT_RUN', modules=len(records),
                generator_sha256=overlay.sha(Path(__file__)),
                overlay_manifest_sha256=overlay.sha(ROOT/'SOURCE_MANIFEST.json'),
                module_manifest=records)
    (output/'TEST_ONLY_MANIFEST.json').write_text(json.dumps(data, indent=2)+'\n', encoding='utf8')
    return {k:data[k] for k in ['test_only','negative_control','numerical_evidence_allowed','modules','status']}


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output', required=True, type=Path)
    parser.add_argument('--negative-control', action='store_true')
    args = parser.parse_args()
    print(json.dumps(generate(args.output, args.negative_control)))
