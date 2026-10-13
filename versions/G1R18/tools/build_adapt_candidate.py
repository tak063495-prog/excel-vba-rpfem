"""Build the isolated G1R18 adaptive candidate from the retained G1R17 source."""
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
BASE = ROOT.parent / "G1R17/src/vba/RPX_Adapt.bas"


def replace_once(source, old, new):
    if source.count(old) != 1:
        raise ValueError(f"Expected one occurrence of {old[:100]!r}")
    return source.replace(old, new, 1)


def main():
    source = BASE.read_text(encoding="utf-8")
    source = replace_once(source, "    originalModel As String\nEnd Type", "    originalModel As String\n    lowerStressStart() As Long\n    lowerLoadVariable As Long\n    lowerReactions As Object\nEnd Type")
    source = replace_once(source, "    pending.originalModel = RPX_ExactModel(): pending.complete = True", "    pending.originalModel = RPX_ExactModel(): pending.complete = True\n    CaptureWitnessLowerLayout pending")
    source = replace_once(source, "        pending.phase = src.phase: pending.gravity = src.gravity: pending.factor = src.factor: pending.originalModel = src.originalModel", "        pending.phase = src.phase: pending.gravity = src.gravity: pending.factor = src.factor: pending.originalModel = src.originalModel\n        pending.lowerStressStart = src.lowerStressStart: pending.lowerLoadVariable = src.lowerLoadVariable\n        Set pending.lowerReactions = CopyWitnessReactions(src.lowerReactions)")
    source = replace_once(source, "    RMatId = state.materialId: RRigidId = state.rigidId: RBody0 = state.body0: RBody1 = state.body1", "    RMatId = state.materialId: RRigidId = state.rigidId: RBody0 = state.body0: RBody1 = state.body1\n    RStressStart = state.lowerStressStart: RLoadVariable = state.lowerLoadVariable\n    Set RReactions = CopyWitnessReactions(state.lowerReactions)")
    source = replace_once(source, "Private Const AdaptCoarseTol As Double = 0.005", "Private Const AdaptCoarseTol As Double = 0.001")
    source = replace_once(source, "Private adaptFinalTarget As Long", "Private adaptFinalTarget As Long\nPrivate adaptAccuracyCap As Long, adaptAccuracyRounds As Long")
    source = replace_once(source, "Optional ByVal forceLegacy As Boolean = False) As Boolean\n    Dim fraction As Double, score() As Double, w() As Double, vel() As Double", "Optional ByVal forceLegacy As Boolean = False, Optional ByVal useIndicator As Boolean = False) As Boolean\n    Dim fraction As Double, score() As Double, w() As Double, vel() As Double")
    source = replace_once(source, "        If total > 0# Then\n            For e = 0 To re - 1\n                If robustRefine Then score(e) = w(e) Else score(e) = w(e) * RArea(e)", "        If useIndicator And VectorReady(baseline.indicator, re - 1) Then\n            For e = 0 To re - 1: score(e) = baseline.indicator(e) * RArea(e): Next e\n        ElseIf total > 0# Then\n            For e = 0 To re - 1\n                If robustRefine Then score(e) = w(e) Else score(e) = w(e) * RArea(e)")
    source = replace_once(source, "    adaptFinalTarget = DTarget\n    startTarget", "    adaptFinalTarget = DTarget\n    adaptAccuracyCap = CLng(RPX_Setting(\"ADAPT_ACCURACY_CAP\", adaptFinalTarget))\n    adaptAccuracyRounds = CLng(RPX_Setting(\"ADAPT_ACCURACY_ROUNDS\", 4))\n    If adaptAccuracyCap < adaptFinalTarget Or adaptAccuracyCap > 50000 Then err.Raise 5, \"RPX_Adapt\", \"INVALID_ADAPT_ACCURACY_CAP\"\n    If adaptAccuracyRounds < 0 Or adaptAccuracyRounds > 20 Then err.Raise 5, \"RPX_Adapt\", \"INVALID_ADAPT_ACCURACY_ROUNDS\"\n    startTarget")
    source = replace_once(source, "    Call RestoreFinalBounds(mode)\n    RPX_TestHook \"final_accept\"", "    Call RestoreFinalBounds(mode)\n    Call ContinueAccuracy(mode, targetGap)\n    Call RestoreFinalBounds(mode)\n    Call ContinueLowerAngularFan(mode)\n    Call RecheckIndependentQuadrature(mode)\n    RPX_TestHook \"final_accept\"")
    helper = (ROOT / "tools/accuracy_continuation.bas.txt").read_text(encoding="utf-8")
    helper += "\n" + (ROOT / "tools/lower_angular_fan.bas.txt").read_text(encoding="utf-8")
    helper += "\n" + (ROOT / "tools/independent_quadrature.bas.txt").read_text(encoding="utf-8")
    helper += "\n" + (ROOT / "tools/witness_lower_layout.bas.txt").read_text(encoding="utf-8")
    source = replace_once(source, "Public Sub RPX_RunAdapt(ByVal mode As String)", helper + "\nPublic Sub RPX_RunAdapt(ByVal mode As String)")
    for folder, encoding in [("src/vba", "utf-8"), ("import", "cp932")]:
        path = ROOT / folder / "RPX_Adapt.bas"
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(source.replace("\n", "\r\n").encode(encoding))
    geometry = (ROOT / "src/vba/RPX_AccuracyMesh.bas").read_text(encoding="utf-8")
    (ROOT / "import/RPX_AccuracyMesh.bas").write_bytes(geometry.replace("\n", "\r\n").encode("cp932"))
    guard = (BASE.parent / "RPX_Guard.bas").read_text(encoding="utf-8")
    guard = replace_once(guard, '    If mode = "Fs" Then\n        \' Endpoint code certifies', '''    If mode = "Fs" Then
        If Not RPX_PrecisionQuadratureReady() Then Exit Function
        ' Endpoint code certifies''')
    guard += '''
' A precision profile requires actual certificates at its requested quadrature.
Private Function RPX_PrecisionQuadratureReady() As Boolean
    Dim q As Long
    q = CLng(RPX_Setting("ADAPT_CERT_QUADRATURE", 0))
    If q <> 0 And q <> 2 And q <> 4 And q <> 8 Then err.Raise 5, "RPX_Guard", "INVALID_ADAPT_CERT_QUADRATURE"
    If q = 0 Then RPX_PrecisionQuadratureReady = True: Exit Function
    If Not RPX_CertCurrent("lower") Or Not RPX_CertCurrent("upper") Then Exit Function
    If RPX_CertLower.q <> q Or RPX_CertUpper.q <> q Then Exit Function
    If Not RPX_CertLower.rootBracketValid Or Not RPX_CertUpper.rootBracketValid Then Exit Function
    RPX_PrecisionQuadratureReady = True
End Function
'''
    for folder, encoding in [("src/vba", "utf-8"), ("import", "cp932")]:
        (ROOT / folder / "RPX_Guard.bas").write_bytes(guard.replace("\n", "\r\n").encode(encoding))
    print("Created G1R18 RPX_Adapt source and CP932 import; G1R17 is retained.")


if __name__ == "__main__":
    main()
