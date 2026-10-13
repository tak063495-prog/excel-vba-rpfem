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
    source = replace_once(source, "Private Const AdaptCoarseTol As Double = 0.005", "Private Const AdaptCoarseTol As Double = 0.001")
    source = replace_once(source, "Private adaptFinalTarget As Long", "Private adaptFinalTarget As Long\nPrivate adaptAccuracyCap As Long, adaptAccuracyRounds As Long")
    source = replace_once(source, "Optional ByVal forceLegacy As Boolean = False) As Boolean\n    Dim fraction As Double, score() As Double, w() As Double, vel() As Double", "Optional ByVal forceLegacy As Boolean = False, Optional ByVal useIndicator As Boolean = False) As Boolean\n    Dim fraction As Double, score() As Double, w() As Double, vel() As Double")
    source = replace_once(source, "        If total > 0# Then\n            For e = 0 To re - 1\n                If robustRefine Then score(e) = w(e) Else score(e) = w(e) * RArea(e)", "        If useIndicator And VectorReady(baseline.indicator, re - 1) Then\n            For e = 0 To re - 1: score(e) = baseline.indicator(e) * RArea(e): Next e\n        ElseIf total > 0# Then\n            For e = 0 To re - 1\n                If robustRefine Then score(e) = w(e) Else score(e) = w(e) * RArea(e)")
    source = replace_once(source, "    adaptFinalTarget = DTarget\n    startTarget", "    adaptFinalTarget = DTarget\n    adaptAccuracyCap = CLng(RPX_Setting(\"ADAPT_ACCURACY_CAP\", adaptFinalTarget))\n    adaptAccuracyRounds = CLng(RPX_Setting(\"ADAPT_ACCURACY_ROUNDS\", 4))\n    If adaptAccuracyCap < adaptFinalTarget Or adaptAccuracyCap > 50000 Then err.Raise 5, \"RPX_Adapt\", \"INVALID_ADAPT_ACCURACY_CAP\"\n    If adaptAccuracyRounds < 0 Or adaptAccuracyRounds > 20 Then err.Raise 5, \"RPX_Adapt\", \"INVALID_ADAPT_ACCURACY_ROUNDS\"\n    startTarget")
    source = replace_once(source, "    Call RestoreFinalBounds(mode)\n    RPX_TestHook \"final_accept\"", "    Call RestoreFinalBounds(mode)\n    Call ContinueAccuracy(mode, targetGap)\n    Call RestoreFinalBounds(mode)\n    RPX_TestHook \"final_accept\"")
    helper = (ROOT / "tools/accuracy_continuation.bas.txt").read_text(encoding="utf-8")
    source = replace_once(source, "Public Sub RPX_RunAdapt(ByVal mode As String)", helper + "\nPublic Sub RPX_RunAdapt(ByVal mode As String)")
    for folder, encoding in [("src/vba", "utf-8"), ("import", "cp932")]:
        path = ROOT / folder / "RPX_Adapt.bas"
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(source.replace("\n", "\r\n").encode(encoding))
    print("Created G1R18 RPX_Adapt source and CP932 import; G1R17 is retained.")


if __name__ == "__main__":
    main()
