from pathlib import Path
import difflib

ROOT = Path(__file__).resolve().parents[1]
def edit(name, old, new):
    path = ROOT / 'src' / name
    text = path.read_text(encoding='utf-8')
    assert text.count(old) == 1, name
    path.write_text(text.replace(old, new), encoding='utf-8', newline='\n')

edit('RPX_Control.bas', '''        Else
            lower = RLowerLoad: upper = RUpperLoad
            RPX_BearingMeasure CStr(RPX_Setting("BEARING_FACES", ""))
        End If
        Call RPX_WriteMesh: Call RPX_Draw''', '''        Else
            lower = RLowerLoad: upper = RUpperLoad
        End If
        ' G1R12/#1: write geometry and velocity from the same accepted upper witness.
        Call RPX_PrepareUpperOutput(True)
        If mode = "極限支持力" Then RPX_BearingMeasure CStr(RPX_Setting("BEARING_FACES", ""))
        Call RPX_WriteMesh: Call RPX_Draw''')

edit('RPX_Output.bas', 'Private Sub RPX_WriteFieldsCore(ByVal adaptive As Boolean)', '''' G1R12/#1: restore before mesh tables/plots, not after writing the upper field.
Public Sub RPX_PrepareUpperOutput(ByVal adaptive As Boolean)
    RPX_AssertInputs "prepare_upper_output"
    If adaptive Then
        RPX_RestoreWitnessState RBestUpper
        RUpperField = RBestUpper.field
    End If
    RPX_AssertUpperOutput adaptive
End Sub

Public Sub RPX_AssertUpperOutput(ByVal adaptive As Boolean)
    Dim i As Long
    RPX_AssertInputs "assert_upper_output"
    If adaptive Then
        If Not RBestUpper.complete Then err.Raise 5, "RPX_Output", "UPPER_OUTPUT_OWNER_INCOMPLETE"
        If RBestUpper.inputEpoch <> RPX_InputEpoch Or RBestUpper.policy <> RPX_AnalysisPolicy() Then err.Raise 5, "RPX_Output", "UPPER_OUTPUT_INPUT_MISMATCH"
        If RBestUpper.originalModel <> RPX_ExactModel() Then err.Raise 5, "RPX_Output", "UPPER_OUTPUT_MODEL_MISMATCH"
        If RBestUpper.q <> RSubdivisions Or RBestUpper.factor <> RStrengthFactor Or RBestUpper.stressScale <> RStressScale Then err.Raise 5, "RPX_Output", "UPPER_OUTPUT_CONTEXT_MISMATCH"
    End If
    If LBound(RUpperField) <> 0 Or UBound(RUpperField) <> 12 * re + 3 * RR - 1 Then err.Raise 5, "RPX_Output", "UPPER_OUTPUT_FIELD_SHAPE"
    If adaptive Then
        If LBound(RBestUpper.field) <> 0 Or UBound(RBestUpper.field) <> UBound(RUpperField) Then err.Raise 5, "RPX_Output", "UPPER_OUTPUT_OWNER_SHAPE"
        For i = 0 To UBound(RUpperField)
            If RUpperField(i) <> RBestUpper.field(i) Then err.Raise 5, "RPX_Output", "UPPER_OUTPUT_FIELD_MISMATCH"
        Next i
    End If
End Sub

Private Sub RPX_WriteFieldsCore(ByVal adaptive As Boolean)''')
edit('RPX_Output.bas', '    RPX_InternalWrite = True', '    RPX_AssertUpperOutput adaptive\n    RPX_InternalWrite = True')
edit('RPX_Output.bas', '    If adaptive Then RPX_RestoreWitnessState RBestLower', '''    If adaptive Then
        RPX_RestoreWitnessState RBestLower
        RLowerField = RBestLower.field
    End If''')
edit('RPX_Output.bas', '    If adaptive Then RPX_RestoreWitnessState RBestUpper', '    If adaptive Then Call RPX_PrepareUpperOutput(True)')

edit('RPX_Robust.bas', 'Public Function RPX_RobustEligible(ByVal number As Long, ByVal message As String) As Boolean', '''' Accept only the detailed token emitted by FactorSystem; retain its text in evidence.
Private Function FactorizationFailure(ByVal message As String) As Boolean
    Const prefix As String = "NEWTON_FACTORIZATION_FAILED pivot="
    Dim digits As String, i As Long, first As Long, negative As Boolean, value As Double
    If Left$(message, Len(prefix)) <> prefix Then Exit Function
    digits = Mid$(message, Len(prefix) + 1): first = 1
    If Left$(digits, 1) = "-" Then negative = True: first = 2
    If Len(digits) < first Or Len(digits) - first + 1 > 10 Then Exit Function
    For i = first To Len(digits)
        If Mid$(digits, i, 1) < "0" Or Mid$(digits, i, 1) > "9" Then Exit Function
    Next i
    value = CDbl(Mid$(digits, first))
    If negative Then
        If value > 2147483648# Then Exit Function
    Else
        If value > 2147483647# Then Exit Function
    End If
    FactorizationFailure = True
End Function
Public Function RPX_RobustEligible(ByVal number As Long, ByVal message As String) As Boolean''')
edit('RPX_Robust.bas', '''            RPX_RobustEligible = True
    End Select''', '''            RPX_RobustEligible = True
        Case Else
            RPX_RobustEligible = FactorizationFailure(Trim$(message))
    End Select''')

diff = []
for path in sorted((ROOT / 'src').glob('*.bas')):
    old = (ROOT / 'original' / path.name).read_text(encoding='utf-8')
    new = path.read_text(encoding='utf-8')
    if old != new:
        diff.extend(difflib.unified_diff(old.splitlines(True), new.splitlines(True), 'G1R11/' + path.name, 'G1R12/' + path.name))
        for folder, text in [('import_cp932', new), ('rollback_cp932', old)]:
            payload = text.replace('\r\n', '\n').replace('\n', '\r\n').encode('cp932', errors='strict')
            assert payload.decode('cp932').replace('\r\n', '\n') == text.replace('\r\n', '\n')
            (ROOT / folder / path.name).write_bytes(payload)
(ROOT / 'changes.patch').write_text(''.join(diff), encoding='utf-8', newline='\n')
print('Implemented RPX_Control, RPX_Output, RPX_Robust; CP932 strict roundtrip PASS')
