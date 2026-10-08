Option Explicit
' P2-S: immutable cell values/formulas, exact model identity, shared failure policy.
Private inputValues(0 To 2) As Variant, inputFormulas(0 To 2) As Variant
Private inputCols(0 To 2) As Variant, inputValueType(0 To 2) As Variant, inputFormulaType(0 To 2) As Variant
Private inputColN(0 To 2) As Long, inputRows(0 To 2) As Long
Private inputActive As Boolean, inputSaved As Boolean, inputUseFast As Boolean
Private Const Hopt11Faster As Double = 0.8
Private lastModel As String, modelRevision As Long
Private Type DoubleBits
    value As Double
End Type
Private Type ByteBits
    value(0 To 7) As Byte
End Type
Private injectStage As String, injectNumber As Long, injectText As String
Private injectSkip As Long, injectArmed As Boolean
Private nextStage As String, nextNumber As Long, nextText As String
Private forceTrialPositive As Boolean
Public RPX_TestHits As Long
Public RPX_LastErrorLine As Long, RPX_LastErrorSource As String
Public RPX_InputEpoch As Long
Public RPX_LoadPhase As String
Public Function RPX_EngineIdentity() As String
    RPX_EngineIdentity = "20260926_p2s_f01_f07"
End Function

Public Function RPX_InputSheetIndex(ByVal name As String) As Long
    RPX_InputSheetIndex = -1
    Select Case name
        Case "要素定義": RPX_InputSheetIndex = 0
        Case "材料データ": RPX_InputSheetIndex = 1
        Case "設定": RPX_InputSheetIndex = 2
    End Select
End Function
Private Function InputRange(ByVal i As Long) As Range
    Select Case i
        Case 0: Set InputRange = ThisWorkbook.Worksheets("要素定義").Range("A7:T1006")
        Case 1: Set InputRange = ThisWorkbook.Worksheets("材料データ").Range("A5:Q1004")
        Case 2: Set InputRange = ThisWorkbook.Worksheets("設定").Range("A4:D100")
    End Select
End Function
Private Function InputColumn(ByVal i As Long, ByVal c As Long) As Boolean
    Select Case i
        Case 0: InputColumn = (c <= 3 Or (c >= 5 And c <= 8) Or c >= 10)
        Case 1: InputColumn = True
        Case 2: InputColumn = (c = 2 Or c = 4)
    End Select
End Function
Public Function RPX_IsInputChange(ByVal ws As Object, ByVal changed As Range) As Boolean
    Dim area As Range, part As Range, i As Long, c As Long
    i = RPX_InputSheetIndex(CStr(ws.name))
    If i < 0 Then Exit Function
    Set area = Application.Intersect(changed, InputRange(i))
    If area Is Nothing Then Exit Function
    For Each part In area.Areas
        For c = part.column To part.column + part.columns.count - 1
            If InputColumn(i, c) Then RPX_IsInputChange = True: Exit Function
        Next c
    Next part
End Function
Public Sub RPX_InputBegin()
    Dim i As Long
    inputActive = False: inputSaved = False
    ' Recalculate before capture so values and formulas describe the same input.
    For i = 0 To 2: InputRange(i).Calculate: Next i
    For i = 0 To 2
        inputValues(i) = InputRange(i).Value2
        inputFormulas(i) = InputRange(i).Formula
    Next i
    inputUseFast = False
    BuildInputMeta
    RPX_InputEpoch = RPX_InputEpoch + 1
    inputSaved = True: inputActive = True
End Sub
Private Sub BuildInputMeta()
    Dim i As Long, r As Long, c As Long, k As Long, n As Long, ub As Long, rows As Long
    Dim saved As Variant, forms As Variant
    Dim cols() As Long, valueTypes() As Long, formulaTypes() As Long
    For i = 0 To 2
        saved = inputValues(i): forms = inputFormulas(i)
        rows = UBound(saved, 1): ub = UBound(saved, 2)
        ReDim cols(1 To ub): n = 0
        For c = 1 To ub
            If InputColumn(i, c) Then n = n + 1: cols(n) = c
        Next c
        ReDim valueTypes(1 To rows, 1 To ub)
        ReDim formulaTypes(1 To rows, 1 To ub)
        For r = 1 To rows
            For k = 1 To n
                c = cols(k)
                valueTypes(r, c) = VarType(saved(r, c))
                formulaTypes(r, c) = VarType(forms(r, c))
            Next k
        Next r
        inputRows(i) = rows: inputColN(i) = n
        inputCols(i) = cols: inputValueType(i) = valueTypes: inputFormulaType(i) = formulaTypes
    Next i
End Sub
Private Function LegacyCore(ByRef live As Variant, ByRef formulas As Variant, ByRef saved As Variant, ByRef oldFormulas As Variant, ByRef cols As Variant, ByVal n As Long, ByVal rowCount As Long) As Boolean
    Dim r As Long, k As Long, c As Long
    For r = 1 To rowCount
        For k = 1 To n
            c = cols(k)
            If Not SameCell(live(r, c), saved(r, c)) Or Not SameCell(formulas(r, c), oldFormulas(r, c)) Then
                LegacyCore = True
                Exit Function
            End If
        Next k
    Next r
End Function
Private Function FastCore(ByRef live As Variant, ByRef formulas As Variant, ByRef saved As Variant, ByRef oldFormulas As Variant, ByRef valueTypes As Variant, ByRef formulaTypes As Variant, ByRef cols As Variant, ByVal n As Long, ByVal rowCount As Long) As Boolean
    ' Inlined. A helper per cell would cost as much as SameCell.
    ' VBA inequality keeps +0 equal to -0. Not a raw bit compare.
    Dim r As Long, k As Long, c As Long, liveType As Long, savedType As Long
    For r = 1 To rowCount
        For k = 1 To n
            c = cols(k)
            liveType = VarType(live(r, c)): savedType = valueTypes(r, c)
            If liveType <> savedType Then FastCore = True: Exit Function
            If liveType <> vbEmpty Then
                If liveType = vbError Then
                    If CStr(live(r, c)) <> CStr(saved(r, c)) Then FastCore = True: Exit Function
                ElseIf liveType = vbDouble Or liveType = vbString Or liveType = vbBoolean Or liveType = vbDate Or liveType = vbCurrency Or liveType = vbLong Or liveType = vbInteger Or liveType = vbByte Or liveType = vbSingle Then
                    If live(r, c) <> saved(r, c) Then FastCore = True: Exit Function
                ElseIf Not SameCell(live(r, c), saved(r, c)) Then
                    FastCore = True: Exit Function
                End If
            End If
            liveType = VarType(formulas(r, c)): savedType = formulaTypes(r, c)
            If liveType <> savedType Then FastCore = True: Exit Function
            If liveType <> vbEmpty Then
                If liveType = vbError Then
                    If CStr(formulas(r, c)) <> CStr(oldFormulas(r, c)) Then FastCore = True: Exit Function
                ElseIf liveType = vbDouble Or liveType = vbString Or liveType = vbBoolean Or liveType = vbDate Or liveType = vbCurrency Or liveType = vbLong Or liveType = vbInteger Or liveType = vbByte Or liveType = vbSingle Then
                    If formulas(r, c) <> oldFormulas(r, c) Then FastCore = True: Exit Function
                ElseIf Not SameCell(formulas(r, c), oldFormulas(r, c)) Then
                    FastCore = True: Exit Function
                End If
            End If
        Next k
    Next r
End Function
Public Sub RPX_BenchInputCompare()
    Dim i As Long, rep As Long, agree As Long
    Dim oldSeconds As Double, newSeconds As Double, t0 As Double
    Dim live As Variant, formulas As Variant, saved As Variant, oldFormulas As Variant
    Dim cols As Variant, valueTypes As Variant, formulaTypes As Variant, probe As Variant
    Dim modeName As String
    If Not inputSaved Then Exit Sub
    oldSeconds = 0#: newSeconds = 0#: agree = 0: inputUseFast = False: modeName = "legacy"
    For i = 0 To 2
        live = inputValues(i): formulas = inputFormulas(i)
        saved = inputValues(i): oldFormulas = inputFormulas(i)
        cols = inputCols(i): valueTypes = inputValueType(i): formulaTypes = inputFormulaType(i)
        t0 = RPX_DiagClock()
        For rep = 1 To 20
            If LegacyCore(live, formulas, saved, oldFormulas, cols, inputColN(i), inputRows(i)) Then err.Raise 5, , "HOPT11_IDENTICAL_LEGACY"
        Next rep
        oldSeconds = oldSeconds + (RPX_DiagClock() - t0)
        t0 = RPX_DiagClock()
        For rep = 1 To 20
            If FastCore(live, formulas, saved, oldFormulas, valueTypes, formulaTypes, cols, inputColN(i), inputRows(i)) Then err.Raise 5, , "HOPT11_IDENTICAL_FAST"
        Next rep
        newSeconds = newSeconds + (RPX_DiagClock() - t0)
    Next i
    probe = inputValues(2): saved = inputValues(2): oldFormulas = inputFormulas(2)
    cols = inputCols(2): valueTypes = inputValueType(2): formulaTypes = inputFormulaType(2)
    probe(1, 2) = "HOPT11_PROBE"
    If LegacyCore(probe, oldFormulas, saved, oldFormulas, cols, inputColN(2), inputRows(2)) Then
        If FastCore(probe, oldFormulas, saved, oldFormulas, valueTypes, formulaTypes, cols, inputColN(2), inputRows(2)) Then agree = 1
    End If
    If agree = 1 And oldSeconds > 0# And newSeconds <= oldSeconds * Hopt11Faster Then
        inputUseFast = True: modeName = "fast"
    End If
    RPX_DiagEvent "hopt11_bench", "reps=20;old_seconds=" & oldSeconds & ";new_seconds=" & newSeconds & ";agree=" & agree & ";mode=" & modeName
End Sub
Public Sub RPX_InputEnd()
    inputActive = False
End Sub
Public Sub RPX_DumpInputSnapshot(ByVal stream As Object)
    Dim i As Long, r As Long, c As Long, values As Variant, formulas As Variant, value As Variant, encoded As String
    stream.WriteLine "input_epoch" & vbTab & RPX_InputEpoch & vbTab & "saved=" & inputSaved
    For i = 0 To 2
        If inputSaved Then
            values = inputValues(i): formulas = inputFormulas(i)
        Else
            values = InputRange(i).Value2: formulas = InputRange(i).Formula
        End If
        For r = 1 To UBound(values, 1)
            For c = 1 To UBound(values, 2)
                If InputColumn(i, c) Then
                    value = values(r, c)
                    If Not IsEmpty(value) Then
                        If VarType(value) = vbDouble Then
                            encoded = RPX_ExactDouble(CDbl(value))
                        Else
                            encoded = Replace(Replace(Replace(CStr(value), vbTab, "\t"), vbCr, "\r"), vbLf, "\n")
                        End If
                        stream.WriteLine "input" & vbTab & i & vbTab & r & vbTab & c & vbTab & VarType(value) & vbTab & encoded & vbTab & Replace(Replace(Replace(CStr(formulas(r, c)), vbTab, "\t"), vbCr, "\r"), vbLf, "\n")
                    End If
                End If
            Next c
        Next r
    Next i
End Sub
Public Function RPX_InputCell(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long) As Variant
    Dim i As Long, a As Variant, firstRow As Long
    i = RPX_InputSheetIndex(ws.name)
    If inputActive And i >= 0 Then
        Select Case i
            Case 0: firstRow = 7
            Case 1: firstRow = 5
            Case 2: firstRow = 4
        End Select
        RPX_InputCell = SavedCell(inputValues(i), r - firstRow + 1, c)
    Else
        RPX_InputCell = ws.cells(r, c).Value2
    End If
End Function
Private Function SavedCell(ByRef values As Variant, ByVal r As Long, ByVal c As Long) As Variant
    SavedCell = values(r, c)
End Function
Private Function SameCell(ByVal a As Variant, ByVal b As Variant) As Boolean
    If VarType(a) <> VarType(b) Then Exit Function
    If IsError(a) Then
        SameCell = (CStr(a) = CStr(b))
    Else
        SameCell = (a = b)
    End If
End Function
Public Function RPX_InputsChanged() As Boolean
    Dim i As Long
    Dim live As Variant, saved As Variant, formulas As Variant, oldFormulas As Variant
    Dim cols As Variant, valueTypes As Variant, formulaTypes As Variant
    Dim mark As Long, failNumber As Long, failSource As String, failText As String
    mark = RPX_HoptMark()
    On Error GoTo HoptFailed
    If Not inputSaved Then GoTo HoptDone
    For i = 0 To 2
        RPX_HoptEnter HoptInputRead
        live = InputRange(i).Value2: formulas = InputRange(i).Formula
        RPX_HoptLeave HoptInputRead
        ' R4 B10: pass the saved arrays ByRef. A local assignment copies the sheet snapshot.
        RPX_HoptEnter HoptInputCompare
        If inputUseFast Then
            If FastCore(live, formulas, inputValues(i), inputFormulas(i), inputValueType(i), inputFormulaType(i), inputCols(i), inputColN(i), inputRows(i)) Then
                RPX_InputsChanged = True
                RPX_HoptLeave HoptInputCompare
                GoTo HoptDone
            End If
        ElseIf LegacyCore(live, formulas, inputValues(i), inputFormulas(i), inputCols(i), inputColN(i), inputRows(i)) Then
            RPX_InputsChanged = True
            RPX_HoptLeave HoptInputCompare
            GoTo HoptDone
        End If
        RPX_HoptLeave HoptInputCompare
    Next i
HoptDone:
    Exit Function
HoptFailed:
    failNumber = err.number: failSource = err.source: failText = err.description
    RPX_HoptUnwind mark
    err.Raise failNumber, failSource, failText
End Function
Public Sub RPX_AssertInputs(ByVal stage As String)
    If XCancel Then err.Raise 18, "RPX_Guard", "ANALYSIS_CANCELLED"
    If Not inputActive Then Exit Sub
    If RPX_InputsChanged() Then
        XCancel = True: RPX_ResultCurrent = False: RPX_MeshCurrent = False
        RPX_FsBracketReset
        err.Raise 5, "RPX_Guard", "INPUT_MISMATCH " & stage
    End If
End Sub
Public Function RPX_ExactDouble(ByVal value As Double) As String
    Dim d As DoubleBits, b As ByteBits, i As Long, s As String
    d.value = value: LSet b = d
    For i = 0 To 7: s = s & Right$("0" & Hex$(b.value(i)), 2): Next i
    RPX_ExactDouble = s
End Function
Public Function RPX_ExactModel() As String
    Dim p() As String, n As Long, i As Long, j As Long, key As Variant, rec As Variant
    ReDim p(0 To rn + re + RM + RR + RNEdge + RConstraints.count + 1)
    p(n) = rn & "," & re & "," & RM & "," & RR & "," & RNEdge & "," & RGravityMode & "," & RPX_LoadPhase & "," & RPX_ExactDouble(RStressScale) & "," & RPX_AnalysisPolicy() & ",raw_audit=" & RRawAuditEnabled: n = n + 1
    For i = 0 To rn - 1
        p(n) = RPX_ExactDouble(RXY(i, 0)) & RPX_ExactDouble(RXY(i, 1)): n = n + 1
    Next i
    For i = 0 To re - 1
        p(n) = RTri(i, 0) & "," & RTri(i, 1) & "," & RTri(i, 2) & "," & RMatId(i) & "," & RRigidId(i)
        For j = 0 To 1: p(n) = p(n) & "," & RPX_ExactDouble(RBody0(i, j)) & RPX_ExactDouble(RBody1(i, j)): Next j
        n = n + 1
    Next i
    For i = 0 To RM - 1
        With RMaterials(i)
            p(n) = RPX_ExactDouble(.cohesion) & RPX_ExactDouble(.friction) & RPX_ExactDouble(.dilation) & RPX_ExactDouble(.gamma)
        End With
        n = n + 1
    Next i
    For i = 0 To RR - 1
        With RRigids(i)
            p(n) = RPX_ExactDouble(.center(0)) & RPX_ExactDouble(.center(1))
            For j = 0 To 2: p(n) = p(n) & "," & .fixed(j) & RPX_ExactDouble(.load0(j)) & RPX_ExactDouble(.load1(j)): Next j
        End With
        n = n + 1
    Next i
    For i = 0 To RNEdge - 1
        With REdges(i)
            p(n) = .a & "," & .b & "," & .element(0) & "," & .element(1) & "," & .kind & "," & .rigidId
            For j = 0 To 1: p(n) = p(n) & "," & RPX_ExactDouble(.load0(j)) & RPX_ExactDouble(.load1(j)): Next j
        End With
        n = n + 1
    Next i
    For Each key In RConstraints.keys
        rec = RConstraints(key): p(n) = CStr(key) & "," & rec(0) & "," & rec(1): n = n + 1
    Next key
    RPX_ExactModel = Join(p, "|")
End Function
Public Function RPX_ModelVersion() As Long
    Dim live As String
    RPX_AssertInputs "model_identity"
    live = RPX_ExactModel()
    If live <> lastModel Then
        modelRevision = modelRevision + 1
        lastModel = live
        RPX_FsBracketReset
        RProgramKind = ""
    End If
    RPX_ModelVersion = modelRevision
End Function
Public Function RPX_Recoverable(ByVal number As Long, ByVal message As String, ByVal stage As String, ByVal verifiedBase As Boolean) As Boolean
    Dim token As String
    If XCancel Or number <> 5 Then Exit Function
    token = Split(Trim$(message), " ")(0)
    ' Exact identifiers and explicit phases; initial construction/restoration never recovers.
    Select Case stage
        Case "fs_trial", "evaluate", "pilot_upper"
            Select Case token
                Case "HSD_LINEAR_QUALITY_FAILED", "HSD_PIVOT_FACTORIZATION_FAILED", "HSD_BORDER_QUALITY_FAILED", "HSD_STABLE_UNAVAILABLE_MEMORY"
                    RPX_Recoverable = (stage = "fs_trial" Or stage = "evaluate")
                Case "CONIC_STEP_STALLED", "CONIC_NOT_CONVERGED", "NEWTON_FACTORIZATION_FAILED", "HSD_STEP_STALLED", "HSD_NOT_CONVERGED", "HSD_SCHUR_SINGULAR", "HSD_SOLVE_BUDGET_EXCEEDED"
                    RPX_Recoverable = True
                Case "PILOT_UPPER_INCOMPLETE", "PILOT_UPPER_NOT_OPTIMAL", "PILOT_UPPER_NOT_POSITIVE", "PILOT_PAIR_INCOMPLETE"
                    RPX_Recoverable = verifiedBase And stage <> "fs_trial"
            End Select
        Case "band_recover"
            RPX_Recoverable = verifiedBase And token = "CONSTRAINED_EDGE_RECOVERY_FAILED"
        Case "band_regions", "band_relocate", "refine"
            RPX_Recoverable = verifiedBase And (token = "NONPOSITIVE_TRIANGLE" Or token = "DEGENERATE_MESH_ELEMENT" Or token = "NO_ELEMENTS_IN_REGIONS" Or token = "RIGID_GEOMETRY_MISMATCH")
        Case "pilot_weight"
            RPX_Recoverable = verifiedBase And token = "PILOT_WEIGHT_EMPTY"
        Case "g1_trial"
            Select Case token
                Case "HSD_LINEAR_QUALITY_FAILED", "HSD_PIVOT_FACTORIZATION_FAILED", "HSD_BORDER_QUALITY_FAILED"
                    RPX_Recoverable = True
                Case "CONIC_STEP_STALLED", "CONIC_NOT_CONVERGED", "NEWTON_FACTORIZATION_FAILED", "HSD_STEP_STALLED", "HSD_NOT_CONVERGED", "HSD_SCHUR_SINGULAR", "HSD_SOLVE_BUDGET_EXCEEDED"
                    RPX_Recoverable = True
                Case "FS_UPPER_ENDPOINT_INVALID", "G1_OBJECTIVE_REJECTED"
                    RPX_Recoverable = True
            End Select
        Case "g1_restore"
            RPX_Recoverable = False
    End Select
End Function
Public Sub RPX_ErrorEvidence(ByVal number As Long, ByVal source As String, ByVal message As String, ByVal line As Long, ByVal stage As String)
    RPX_LastErrorLine = line: RPX_LastErrorSource = source
    RPX_DiagEvent "error_origin", "number=" & number & ";source=" & source & ";line=" & line & ";stage=" & stage & ";message=" & message
End Sub
Public Sub RPX_RefineAttempt(ByRef score() As Double, ByVal fraction As Double)
    RPX_TestHook "refine"
    RPX_RefineMesh score, fraction
End Sub
Public Function RPX_TargetMet(ByVal mode As String, ByVal lower As Double, ByVal upper As Double, ByVal pairComplete As Boolean) As Boolean
    If Not pairComplete Or XCancel Or RPX_AuditRejected Then Exit Function
    If upper < lower Or Abs(lower) + Abs(upper) <= 0# Then Exit Function
    If mode = "Fs" Then
        ' Endpoint code certifies width <= tolerance * Max(1, midpoint).
        ' RFsLargestRootBracket is absolute; comparing it directly to tolerance is incorrect.
        If lower <= 0# Or RFsSearchIncomplete Then Exit Function
    End If
    If 2# * (upper - lower) / (Abs(upper) + Abs(lower)) > CDbl(RPX_Setting("GAP", 0.01)) Then Exit Function
    RPX_TargetMet = True
End Function
' Injection requires BOTH setting=1 and explicit arming. A single run consumes it once.
Public Sub RPX_TestArm(ByVal stage As String, ByVal number As Long, ByVal message As String, Optional ByVal skip As Long = 0)
    If RPX_Busy Then err.Raise 5, , "TEST_ARM_WHILE_BUSY"
    If CLng(RPX_Setting("TEST_INJECTION", 0)) <> 1 Then err.Raise 5, , "TEST_INJECTION_DISABLED"
    injectStage = stage: injectNumber = number: injectText = message: injectSkip = skip
    injectArmed = True: RPX_TestHits = 0
End Sub
Public Sub RPX_TestDisarm()
    injectArmed = False
    nextStage = ""
    forceTrialPositive = False
End Sub
Public Sub RPX_TestTrialResult(ByRef residual As Double)
    RPX_TestHook "fs_trial_result"
    If forceTrialPositive Then residual = Abs(residual) + 1#: forceTrialPositive = False
End Sub
Public Sub RPX_TestQueue(ByVal stage As String, ByVal number As Long, ByVal message As String)
    If Not injectArmed Or RPX_Busy Then err.Raise 5, , "TEST_QUEUE_NOT_ARMED"
    nextStage = stage: nextNumber = number: nextText = message
End Sub
Public Sub RPX_TestHook(ByVal stage As String)
    Dim number As Long, message As String
    If Not injectArmed Or stage <> injectStage Then Exit Sub
    If CLng(RPX_Setting("TEST_INJECTION", 0)) <> 1 Then Exit Sub
    If injectSkip > 0 Then injectSkip = injectSkip - 1: Exit Sub
    number = injectNumber: message = injectText
    injectArmed = False: RPX_TestHits = RPX_TestHits + 1
    If Len(nextStage) > 0 Then
        injectStage = nextStage: injectNumber = nextNumber: injectText = nextText
        injectSkip = 0: injectArmed = True: nextStage = ""
    End If
    RPX_DiagEvent "test_injection", "stage=" & stage & ";number=" & number
    Select Case message
        Case "ROOT_INCOMPLETE": RFsSearchIncomplete = True: RFsSearchNote = "TEST_ROOT_WIDTH_UNMET": Exit Sub
        Case "POSITIVE_TRIAL": forceTrialPositive = True: Exit Sub
        Case "MUTATE_FACE": ThisWorkbook.Worksheets("要素定義").Range("R7").Value2 = -10#: Exit Sub
        Case "MUTATE_MATERIAL": ThisWorkbook.Worksheets("材料データ").Range("B5").Value2 = 42#: Exit Sub
        Case "MUTATE_FIXED": ThisWorkbook.Worksheets("材料データ").Range("I5").Value2 = 0: Exit Sub
        Case "MUTATE_EDGE": REdges(0).load0(0) = REdges(0).load0(0) + 1#: Exit Sub
        Case "MUTATE_RIGID": RRigids(0).center(0) = RRigids(0).center(0) + 1#: Exit Sub
    End Select
    If number = 18 Then XCancel = True
7001 err.Raise number, "RPX_TestInjection", message
End Sub