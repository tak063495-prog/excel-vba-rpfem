from pathlib import Path
import difflib,json
R=Path(__file__).resolve().parent
pre=json.loads((R/'results/staged_pretest.json').read_text())
assert any(x['accepted_witness'] for x in pre if x['mesh']=='staged_1000' and x['mode']=='upper')
assert any(x['accepted_witness'] for x in pre if x['mesh']=='staged_1000' and x['mode']=='lower')
assert (R/'results/python_bracket_pretest.json').exists()

p=R/'src/RPX_Refine.bas';s=(R/'original/RPX_Refine.bas').read_text(encoding='utf-8')
s=s.replace('ByVal fraction As Double)\n','ByVal fraction As Double, Optional ByVal forceLegacy As Boolean = False)\n',1)
s=s.replace('red = RPX_RobustMeshPolicy()','red = RPX_RobustMeshPolicy() And Not forceLegacy',1)
s=s.replace('Case "ROBUST_REFINE":','Case "ROBUST_REFINE", "LEGACY_THEN_REFINE":',1)
p.write_text(s,encoding='utf-8')

p=R/'src/RPX_Adapt.bas';s=(R/'original/RPX_Adapt.bas').read_text(encoding='utf-8')
s=s.replace('ByVal maxElements As Long) As Boolean','ByVal maxElements As Long, Optional ByVal forceLegacy As Boolean = False) As Boolean',1)
s=s.replace('robustRefine = RPX_RobustMeshPolicy()','robustRefine = RPX_RobustMeshPolicy() And Not forceLegacy',1)
s=s.replace('        RPX_RefineAttempt score, fraction','''        If forceLegacy Then
            RPX_TestHook "refine"
            RPX_RefineMesh score, fraction, True
        Else
            RPX_RefineAttempt score, fraction
        End If''',1)
s=s.replace('Dim robustRefine As Boolean, robustExhausted As Boolean','Dim robustRefine As Boolean, robustExhausted As Boolean, stagedRefine As Boolean, stagedCap As Long',1)
at=s.index('    robustRefine = RPX_RobustMeshPolicy()',s.index('Public Sub RPX_RunAdapt('))
end=at+len('    robustRefine = RPX_RobustMeshPolicy()')
s=s[:at]+'''    stagedRefine = (UCase$(Trim$(CStr(RPX_Setting("ADAPT_MESH_POLICY", "LEGACY")))) = "LEGACY_THEN_REFINE" And RR = 0)
    robustRefine = RPX_RobustMeshPolicy() And Not stagedRefine
    If stagedRefine Then
        stagedCap = CLng(RPX_Setting("ADAPT_REFINE_CAP", 1000))
        If stagedCap < 16 Or stagedCap > 50000 Then err.Raise 5, "RPX_Adapt", "INVALID_ADAPT_REFINE_CAP"
    End If'''+s[end:]
s=s.replace('";effective_red_green=" & robustRefine & ";budget=" & DTarget','";effective_red_green=" & robustRefine & ";staged=" & stagedRefine & ";staged_cap=" & stagedCap & ";budget=" & DTarget',1)
# Legacy band and legacy centroid fan are kept unchanged before the staged addition.
s=s.replace('If TryShearRefine(baseline, maxElements) Then','If TryShearRefine(baseline, maxElements, stagedRefine) Then',1)
s=s.replace('            If robustRefine Then On Error GoTo ShearRefineFailed','            If robustRefine Or stagedRefine Then On Error GoTo ShearRefineFailed',1)
s=s.replace('            If robustRefine Then meshStage = "evaluate"','            If robustRefine Or stagedRefine Then meshStage = "evaluate"',1)
idx=s.index('AfterCycles:',s.index('Public Sub RPX_RunAdapt('))
block='''    If stagedRefine And cycles > 0 And currentPairValid And RCurrent.pairGap > targetGap Then
        CloneState baseline, RCurrent
        maxElements = stagedCap: If maxElements > adaptFinalTarget Then maxElements = adaptFinalTarget
        If maxElements > baseline.elements Then
            RPX_DiagEvent "adapt_action", "action=legacy_then_red_green;baseline_elements=" & baseline.elements & ";max_elements=" & maxElements
            On Error GoTo StagedRefineFailed
            RPX_DiagBegin dpMesh
            If TryShearRefine(baseline, maxElements) Then
                RPX_DiagEnd dpMesh: meshStage = "evaluate"
                Call Evaluate(mode, "staged_refine")
                If Not AcceptCandidate(baseline, RCurrent) Then
                    RPX_DiagEvent "adapt_rollback", "action=staged_refine;candidate_gap=" & RCurrent.pairGap
                    RPX_RestoreState baseline: CloneState RCurrent, baseline
                End If
                RPX_DiagEvent "adapt_staged_end", "elements=" & re & ";lower=" & RCurrent.pairLower & ";upper=" & RCurrent.pairUpper & ";best_lower=" & RBestLower.value & ";best_upper=" & RBestUpper.value
            Else
                RPX_DiagEnd dpMesh
            End If
            On Error GoTo 0
        End If
    End If
'''
s=s[:idx]+block+s[idx:]
for label in ['FinalCandidateFailed:','ShearRefineFailed:','CycleRemeshFailed:']:
    a=s.index(label);b=s.index('    RPX_',a)
    part=s[a:b].replace('If robustRefine And InStr','If (robustRefine Or stagedRefine) And InStr')
    s=s[:a]+part+s[b:]
# Failures do not enter another hidden refinement loop. Keep existing witnesses.
idx=s.index('RefineRestoreFailed:',s.index('Public Sub RPX_RunAdapt('))
s=s[:idx]+'''StagedRefineFailed:
    savedErrNo = err.number: savedErrText = err.description: savedErrSource = err.source: savedErrLine = Erl
    On Error GoTo 0
    RPX_ErrorEvidence savedErrNo, savedErrSource, savedErrText, savedErrLine, meshStage
    If InStr(1, savedErrText, "HSD_STABLE_UNAVAILABLE_MEMORY", vbBinaryCompare) = 1 Then err.Raise savedErrNo, savedErrSource, savedErrText
    If Not RPX_Recoverable(savedErrNo, savedErrText, meshStage, True) Then err.Raise savedErrNo, savedErrSource, savedErrText
    On Error GoTo RefineRestoreFailed
    RPX_DiagEnd dpMesh
    RPX_RestoreState baseline: CloneState RCurrent, baseline
    currentPairValid = True: adaptStatus = "staged_candidate_rejected"
    RPX_DiagEvent "adapt_staged_rejected", "old_witnesses_kept=1;number=" & savedErrNo & ";message=" & savedErrText
    GoTo AfterCycles
'''+s[idx:]
# Always pass the just-audited upper endpoint, including lower-history startup.
s=s.replace('RPX_FsEndpoint("lower", adaptHintLower, AdaptFsHintRatio)','RPX_FsEndpoint("lower", adaptHintLower, AdaptFsHintRatio, upper)',1)
p.write_text(s,encoding='utf-8')

p=R/'src/RPX_Fs.bas';s=(R/'original/RPX_Fs.bas').read_text(encoding='utf-8')
helper='''Private Function UpperLimitUsable(ByVal upperLimit As Double, ByVal meshKey As String) As Boolean
    Dim policy As String, reason As String, usedMaterials As Object, e As Long
    policy = UCase$(Trim$(CStr(RPX_Setting("FS_BRACKET_POLICY", "LEGACY"))))
    If policy = "LEGACY" Then Exit Function
    If policy <> "AUDITED_UPPER_LIMIT" Then err.Raise 5, "RPX_Fs", "INVALID_FS_BRACKET_POLICY"
    If RR > 0 Or Not RFsActive Or upperLimit <= 0# Or upperLimit > 10000# Then Exit Function
    If Not RPX_SchurFinite(upperLimit) Then Exit Function
    ' Preserve the legacy stages (including their exact fields used for remeshing).
    If UCase$(Trim$(CStr(RPX_Setting("ADAPT_MESH_POLICY", "LEGACY")))) = "LEGACY_THEN_REFINE" And re <= 800 Then Exit Function
    Set usedMaterials = RPX_Row()
    For e = 0 To re - 1
        If RRigidId(e) < 0 Then usedMaterials(CStr(RMatId(e))) = True
    Next e
    If usedMaterials.count < 2 Then Exit Function
    If Not BracketUsable(upperBracket, "upper", 0.999999, meshKey, True, reason) Then Exit Function
    If upperBracket.b <> upperLimit Or upperBracket.statusA <> "OPTIMAL" Or upperBracket.statusB <> "OPTIMAL" Then Exit Function
    If RPX_AuditRejected Or RFsSearchIncomplete Or Not FsArrayReady(upperFieldB) Then Exit Function
    UpperLimitUsable = True
End Function
'''
idx=s.index('Public Function RPX_FsEndpoint(');s=s[:idx]+helper+s[idx:]
s=s.replace('Dim seedCalls As Long, expansionCount As Long, callsBefore As Long','''Dim seedCalls As Long, expansionCount As Long, callsBefore As Long
    Dim limitActive As Boolean, upperLimit As Double, legacyB As Double, nextB As Double
    Dim limitTrial As FsTrialPoint''',1)
s=s.replace('    meshKey = FsMeshKey()','''    meshKey = FsMeshKey()
    If mode = "lower" Then
        limitActive = UpperLimitUsable(upperSeed, meshKey)
        If limitActive Then upperLimit = upperSeed
    End If''',1)
target='''            fa = FsValue(a, mode, target): fieldA = xx: objectiveA = XObjective: statusA = XStatus
            fb = FsValue(b, mode, target): fieldB = xx: objectiveB = XObjective: statusB = XStatus'''
replacement='''            legacyB = b
            If limitActive Then
                If b > upperLimit Then b = upperLimit
                If a >= b Then a = b / 2#
            End If
            fa = FsValue(a, mode, target): fieldA = xx: objectiveA = XObjective: statusA = XStatus
            If limitActive And b = upperLimit Then
                If TryFsTrialPoint(b, mode, target, limitTrial) Then
                    fb = limitTrial.residual
                Else
                    If InStr(1, limitTrial.errorText, "HSD_STABLE_UNAVAILABLE_MEMORY", vbBinaryCompare) = 1 Then err.Raise limitTrial.errorNumber, limitTrial.errorSource, limitTrial.errorText
                    limitActive = False: b = legacyB
                    RPX_DiagEvent "fs_upper_limit_fallback", "reason=numerical_unknown;mode=lower;legacy_b=" & b
                    fb = FsValue(b, mode, target)
                End If
            Else
                fb = FsValue(b, mode, target)
            End If
            fieldB = xx: objectiveB = XObjective: statusB = XStatus'''
# This exact block occurs in the generic startup only (not the quadrature retry).
idx=s.index('        If seedAttempted And trialA.valid And trialB.valid Then')
assert s[idx:].count(target)==1;s=s[:idx]+s[idx:].replace(target,replacement,1)
old='''        a = b: fa = fb: fieldA = fieldB: objectiveA = objectiveB: statusA = statusB: b = b * 2#
        If b > 10000# Then err.Raise 5, , "FS_UPPER_BRACKET_NOT_FOUND"
        fb = FsValue(b, mode, target): fieldB = xx: objectiveB = XObjective: statusB = XStatus'''
new='''        a = b: fa = fb: fieldA = fieldB: objectiveA = objectiveB: statusA = statusB
        nextB = b * 2#
        If limitActive Then
            If b >= upperLimit Then
                limitActive = False
                RPX_DiagEvent "fs_upper_limit_fallback", "reason=actual_lower_sign_positive;mode=lower;at=" & b
            ElseIf nextB > upperLimit Then
                nextB = upperLimit
            End If
        End If
        If nextB > 10000# Then err.Raise 5, , "FS_UPPER_BRACKET_NOT_FOUND"
        If limitActive And nextB = upperLimit Then
            If TryFsTrialPoint(nextB, mode, target, limitTrial) Then
                fb = limitTrial.residual
            Else
                If InStr(1, limitTrial.errorText, "HSD_STABLE_UNAVAILABLE_MEMORY", vbBinaryCompare) = 1 Then err.Raise limitTrial.errorNumber, limitTrial.errorSource, limitTrial.errorText
                limitActive = False: nextB = b * 2#
                If nextB > 10000# Then err.Raise 5, , "FS_UPPER_BRACKET_NOT_FOUND"
                RPX_DiagEvent "fs_upper_limit_fallback", "reason=numerical_unknown;mode=lower;legacy_b=" & nextB
                fb = FsValue(nextB, mode, target)
            End If
        Else
            fb = FsValue(nextB, mode, target)
        End If
        b = nextB: fieldB = xx: objectiveB = XObjective: statusB = XStatus'''
assert old in s;s=s.replace(old,new,1)
s=s.replace('        detail = "lower_hint_source=" & seedSource','        RPX_DiagEvent "fs_upper_limit", "active=" & limitActive & ";hint=" & upperLimit & ";actual_lower_sign_checks=1"\n        detail = "lower_hint_source=" & seedSource',1)
s=s.replace('    RLowerFs = RPX_FsEndpoint("lower"): ReDim RLowerField(0 To UBound(xx))','''    If UpperLimitUsable(RUpperFs, FsMeshKey()) Then
        RLowerFs = RPX_FsEndpoint("lower", 0#, 0.1, RUpperFs)
    Else
        RLowerFs = RPX_FsEndpoint("lower")
    End If
    ReDim RLowerField(0 To UBound(xx))''',1)
p.write_text(s,encoding='utf-8')

patch=''
for name in ['RPX_Refine','RPX_Adapt','RPX_Fs']:
    a=(R/'original'/(name+'.bas')).read_text(encoding='utf-8');b=(R/'src'/(name+'.bas')).read_text(encoding='utf-8')
    patch+=''.join(difflib.unified_diff(a.splitlines(True),b.splitlines(True),fromfile='original/'+name+'.bas',tofile='src/'+name+'.bas'))
    for d,t in [('import_cp932',b),('rollback_cp932',a)]:
        (R/d/(name+'.bas')).write_bytes(('\r\n'.join(t.splitlines())+'\r\n').encode('cp932',errors='strict'))
(R/'changes.patch').write_text(patch,encoding='utf-8')
print('Prepared staged legacy geometry + bounded additional refinement; sign-checked upper hint; unchanged kernel/tolerances')
