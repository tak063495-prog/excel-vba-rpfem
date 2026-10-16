"""Build/check a scheduling-only sibling of the frozen Combined10 sources."""
from pathlib import Path
import argparse,hashlib,json

ROOT=Path(__file__).resolve().parent
REPO=ROOT.parents[1]
BASE=REPO/'versions/G1R19_Combined_20261016/src/vba'

REQUEST_DECLARATIONS='''Private upperFirstReady As Boolean, upperFirstModel As String, upperFirstPolicy As String, upperFirstPhase As String
Private upperFirstFactor As Double, upperFirstEpoch As Long, upperFirstQ As Long
Private upperFirstSolve As Long, upperFirstCandidate As Long
'''
REQUEST_FUNCTIONS='''
' One-shot scheduling request only; never a primal or dual certificate.
Public Sub RPX_AccuracyUpperDualFirstReset()
    upperFirstReady = False: upperFirstModel = "": upperFirstPolicy = "": upperFirstPhase = ""
End Sub
Public Sub RPX_AccuracyUpperDualFirstArm(ByVal factor As Double)
    Call RPX_AccuracyUpperDualFirstReset
    If XCancel Then err.Raise 18, "RPX_AccuracyNative", "ANALYSIS_CANCELLED"
    If Not RPX_SchurFinite(factor) Or factor <= 0# Then err.Raise 5, "RPX_AccuracyNative", "INVALID_UPPER_DUAL_FIRST_FACTOR"
    If Not RPX_DirectEnabled() Or RSubdivisions <> 8 Or RPX_LoadPhase <> "reference_total" Then Exit Sub
    upperFirstFactor = factor: upperFirstEpoch = RPX_InputEpoch: upperFirstQ = RSubdivisions
    upperFirstModel = RPX_ExactModel(): upperFirstPolicy = RPX_AnalysisPolicy(): upperFirstPhase = RPX_LoadPhase
    upperFirstSolve = P4Solve: upperFirstCandidate = RPX_HSDCandidateId: upperFirstReady = True
End Sub
Private Function ConsumeUpperDualFirst(ByVal factor As Double, ByVal mode As String) As Boolean
    If Not upperFirstReady Then GoTo Consumed
    If mode <> "upper" Or Not RPX_DirectEnabled() Or XCancel Then GoTo Consumed
    If Not RPX_SchurFinite(factor) Or factor <> upperFirstFactor Then GoTo Consumed
    If upperFirstEpoch <> RPX_InputEpoch Or upperFirstQ <> RSubdivisions Or RSubdivisions <> 8 Then GoTo Consumed
    If upperFirstModel <> RPX_ExactModel() Or upperFirstPolicy <> RPX_AnalysisPolicy() Or upperFirstPhase <> RPX_LoadPhase Then GoTo Consumed
    If upperFirstSolve <> P4Solve Or upperFirstCandidate <> RPX_HSDCandidateId Then GoTo Consumed
    ConsumeUpperDualFirst = True
Consumed:
    Call RPX_AccuracyUpperDualFirstReset
End Function
'''

TERMINAL_ENABLED='''
Private Function TerminalScheduleEnabled() As Boolean
    Dim enabled As Double
    enabled = CDbl(RPX_Setting("FS_TERMINAL_SCHEDULE", 0))
    If Not RPX_SchurFinite(enabled) Or (enabled <> 0# And enabled <> 1#) Then err.Raise 5, "RPX_Fs", "INVALID_FS_TERMINAL_SCHEDULE"
    TerminalScheduleEnabled = (enabled = 1# And RSubdivisions = 8 And RPX_AccuracyEnabled())
End Function
'''

LOWER_TERMINAL='''        terminalActive = (terminalEnabled And width <= 2.5 * RFsTolerance * valueScale)
        If terminalActive And mode = "lower" Then
            ' Fixed original-span fractions, at most two full-gated trials.
            ' UNKNOWN never changes a signed endpoint or supplies a field.
            terminalA = a: terminalB = b: terminalProgress = False
            x = terminalA + 0.3 * (terminalB - terminalA)
            If x > a And x < b Then
                If TryFsTrialPoint(x, mode, target, trialA) Then
                    If trialA.valid Then
                        If trialA.residual >= 0# Then
                            a = x: fa = trialA.residual: fieldA = xx: objectiveA = trialA.objective: statusA = trialA.status
                        Else
                            b = x: fb = trialA.residual: fieldB = xx: objectiveB = trialA.objective: statusB = trialA.status
                        End If
                        terminalProgress = True
                    End If
                End If
            End If
            valueScale = (a + b) / 2#: If valueScale < 1# Then valueScale = 1#
            If b - a > RFsTolerance * valueScale Then
                x = terminalA + 0.7 * (terminalB - terminalA)
                If x > a And x < b Then
                    If TryFsTrialPoint(x, mode, target, trialB) Then
                        If trialB.valid Then
                            If trialB.residual >= 0# Then
                                a = x: fa = trialB.residual: fieldA = xx: objectiveA = trialB.objective: statusA = trialB.status
                            Else
                                b = x: fb = trialB.residual: fieldB = xx: objectiveB = trialB.objective: statusB = trialB.status
                            End If
                            terminalProgress = True
                        End If
                    End If
                End If
            End If
            RPX_DiagEvent "fs_terminal_lower", "a=" & a & ";b=" & b & ";progress=" & terminalProgress & ";source_a=" & terminalA & ";source_b=" & terminalB & ";max_trials=2;unknown_is_not_a_sign=1"
            If Not terminalProgress Then
                RFsSearchIncomplete = True: RFsSearchNote = "FS_TERMINAL_BOTH_UNKNOWN": Exit For
            End If
            GoTo TerminalIterationComplete
        End If
'''

UPPER_FIRST='''    If useDualFirst Then
        ' A separately declared helper budget can precede the ordinary LP.
        ' Each budget remains explicit; no helper result bypasses original audit.
        dualFirstAttempted = True
        dualFirstSeconds = CDbl(RPX_Setting("ACCURACY_DLL_DUAL_SECONDS", 20))
        If Not RPX_SchurFinite(dualFirstSeconds) Or dualFirstSeconds <= 0# Or dualFirstSeconds > 600# Then err.Raise 5, "RPX_AccuracyNative", "INVALID_UPPER_DUAL_FIRST_SECONDS"
        RPX_AssertInputs "direct_upper_dual_first"
        RPX_TestHook "direct_upper_dual_first"
        If RPX_AccuracyRecoverDual(factor, mode, dualFirstSeconds) Then
            RPX_AssertInputs "direct_upper_dual_first_accepted"
            If Not CurrentProgram(original) Then err.Raise 5, "RPX_AccuracyNative", "UPPER_DUAL_FIRST_CONTEXT_CHANGED"
            kind = "original_dual_helper": accepted = True: helperUsed = True
            GoTo Restore
        End If
        Call RPX_AccuracyGuideReset
        RPX_AssertInputs "direct_upper_dual_first_fallback"
        If XCancel Then err.Raise 18, "RPX_AccuracyNative", "ANALYSIS_CANCELLED"
        If Not CurrentProgram(original) Then err.Raise 5, "RPX_AccuracyNative", "UPPER_DUAL_FIRST_CONTEXT_CHANGED"
        RPX_DiagEvent "direct_upper_dual_first", "accepted=False;fallback=original_lp;helper_seconds=" & dualFirstSeconds & ";lp_seconds=" & seconds & ";guide_is_not_field=1"
    End If
'''


def norm(text):return text.replace('\r\n','\n').replace('\r','\n')
def replace(text,before,after):
    if text.count(before)!=1:raise ValueError('Expected one anchor: '+before[:100])
    return text.replace(before,after,1)
def sha(path):return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def native():
    text=norm((BASE/'RPX_AccuracyNative.bas').read_text(encoding='utf8'))
    text=replace(text,'Public Function RPX_AccuracyEnabled() As Boolean\n',REQUEST_DECLARATIONS+'Public Function RPX_AccuracyEnabled() As Boolean\n')
    text+=REQUEST_FUNCTIONS
    text=replace(text,'    Dim started As Double, remaining As Double\n',
                 '    Dim started As Double, remaining As Double\n    Dim useDualFirst As Boolean, dualFirstAttempted As Boolean, dualFirstSeconds As Double\n')
    text=replace(text,'    AccuracyDirectLastKind = "": AccuracyDirectLastAccepted = False\n',
                 '    useDualFirst = ConsumeUpperDualFirst(factor, mode)\n    AccuracyDirectLastKind = "": AccuracyDirectLastAccepted = False\n')
    text=replace(text,'    started = RPX_DiagClock()\n    RPX_TestHook "direct_dll_begin"\n',
                 UPPER_FIRST+'    started = RPX_DiagClock()\n    RPX_TestHook "direct_dll_begin"\n')
    text=replace(text,'    If Not accepted Then\n        remaining = seconds - (RPX_DiagClock() - started)\n',
                 '    If Not accepted And Not dualFirstAttempted Then\n        remaining = seconds - (RPX_DiagClock() - started)\n')
    return text


def fs():
    text=norm((BASE/'RPX_Fs.bas').read_text(encoding='utf8'))
    text=replace(text,'Private Function TryFsValue(ByVal factor As Double, ByVal mode As String, ByVal target As Double, ByRef value As Double) As Boolean\n',
                 'Private Function TryFsValue(ByVal factor As Double, ByVal mode As String, ByVal target As Double, ByRef value As Double, Optional ByVal preferUpperDual As Boolean = False) As Boolean\n')
    text=replace(text,'    RPX_TestHook "fs_value"\n    value = FsValue(factor, mode, target): TryFsValue = True: Exit Function\n',
                 '    Call RPX_AccuracyUpperDualFirstReset\n    If preferUpperDual And mode = "upper" Then Call RPX_AccuracyUpperDualFirstArm(factor)\n    RPX_TestHook "fs_value"\n    value = FsValue(factor, mode, target)\n    Call RPX_AccuracyUpperDualFirstReset\n    TryFsValue = True: Exit Function\n')
    # Only the first Failed handler is TryFsValue; other trial handlers stay exact.
    first=text.index('Private Function TryFsValue(');last=text.index('\nEnd Function',first)
    part=text[first:last]
    part=replace(part,'    RPX_ErrorEvidence number, source, message, errorLine, "fs_trial"\n',
                 '    Call RPX_AccuracyUpperDualFirstReset\n    RPX_ErrorEvidence number, source, message, errorLine, "fs_trial"\n')
    text=text[:first]+part+text[last:]
    text=replace(text,'Public Sub RPX_FsBracketReset()\n','Public Sub RPX_FsBracketReset()\n    Call RPX_AccuracyUpperDualFirstReset\n')
    text=replace(text,'    Dim limitTrial As FsTrialPoint\n',
                 '    Dim limitTrial As FsTrialPoint\n    Dim terminalEnabled As Boolean, terminalActive As Boolean, terminalProgress As Boolean\n    Dim terminalA As Double, terminalB As Double\n')
    text=replace(text,'    If RFsTolerance <= 0# Then RFsTolerance = 0.0001\n',
                 '    If RFsTolerance <= 0# Then RFsTolerance = 0.0001\n    terminalEnabled = TerminalScheduleEnabled()\n')
    text=replace(text,'        If width <= RFsTolerance * valueScale Then Exit For\n        probe = False\n',
                 '        If width <= RFsTolerance * valueScale Then Exit For\n'+LOWER_TERMINAL+'        probe = False\n')
    # Existing three attempts remain three attempts; only terminal upper calls
    # arm the context-bound one-shot scheduling request.
    if text.count('TryFsValue(x, mode, target, fx)')!=3:raise ValueError('Unexpected ordinary trial count')
    text=text.replace('TryFsValue(x, mode, target, fx)','TryFsValue(x, mode, target, fx, terminalActive)')
    text=replace(text,'    Next i\n    If i > 40 Then RFsSearchIncomplete = True:',
                 'TerminalIterationComplete:\n    Next i\n    If i > 40 Then RFsSearchIncomplete = True:')
    text+=TERMINAL_ENABLED
    return text


def expected():return {'RPX_Fs.bas':fs(),'RPX_AccuracyNative.bas':native()}


def main():
    parser=argparse.ArgumentParser(description=__doc__);parser.add_argument('--initialize',action='store_true');args=parser.parse_args()
    modules=[]
    for name,text in expected().items():
        source=ROOT/'src/vba'/name;imp=ROOT/'import'/name
        if args.initialize:
            for p,encoding in [(source,'utf8'),(imp,'cp932')]:
                p.parent.mkdir(parents=True,exist_ok=True)
                if not p.exists():p.write_bytes(text.replace('\n','\r\n').encode(encoding))
        if norm(source.read_text(encoding='utf8'))!=text:raise ValueError('Candidate differs from controlled scheduling patch: '+name)
        if source.read_bytes().decode('utf8')!=imp.read_bytes().decode('cp932'):raise ValueError('Source/import differs: '+name)
        modules.append(dict(name=name,baseline=(BASE/name).relative_to(REPO).as_posix(),baseline_sha256=sha(BASE/name),
                            source='src/vba/'+name,source_sha256=sha(source),imported='import/'+name,import_sha256=sha(imp)))
    tracked=['README_JA.md','SETTINGS_CONTRACT.json','tests/test_contract.py','tests/MICROTEST_JA.md','build_overlay.py']
    manifest=dict(schema='rpfem.terminal_root.scheduling.v1',status='ISOLATED_SCHEDULING_PROTOTYPE; EXCEL_RUNTIME_PENDING',
                  base='Frozen Combined10',production_artifacts_modified=False,solver_calls=0,modules=modules,
                  unchanged=['Original primal/dual physical equations and gates','ActualRootBracket','TryFsTrialPoint','SaveBracket',
                             'Original dual feasibility program/separation threshold','Forty root iterations','Physical field/Fs unchanged'],
                  profile=json.loads((ROOT/'SETTINGS_CONTRACT.json').read_text()),
                  files={p:sha(ROOT/p) for p in tracked})
    (ROOT/'SOURCE_MANIFEST.json').write_text(json.dumps(manifest,indent=2,ensure_ascii=False)+'\n',encoding='utf8')
    print(json.dumps(dict(status=manifest['status'],manifest_sha256=sha(ROOT/'SOURCE_MANIFEST.json'),modules=modules)))

if __name__=='__main__':main()
