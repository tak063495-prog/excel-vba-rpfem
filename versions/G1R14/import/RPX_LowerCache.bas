Attribute VB_Name = "RPX_LowerCache"
Option Explicit
' P3-B: one immutable compiled lower structure. Never stores iterates or numeric factors.
Private Const CacheByteLimit As Double = 67108864#
Private cacheValid As Boolean, cacheKey As String, cacheBytes As Double
Private savedN As Long, savedM As Long, savedS As Long, savedB As Long
Private savedQScale As Double, savedLoad As Long
Private savedCones() As RPX_ConeData
Private savedQ() As Double, savedEPtr() As Long, savedEId() As Long
Private savedRowNorm() As Double
Private savedEVal() As Double, savedERhs() As Double
Private savedKPtr() As Long, savedKId() As Long, savedKBase() As Double, savedKDiag() As Long
Private savedPerm() As Long, savedInv() As Long, savedStressStart() As Long
Private savedReactions As Collection, savedEqRoles As Collection

Private Function SecondsSince(ByVal started As Double) As Double
    SecondsSince = Timer - started
    If SecondsSince < 0# Then SecondsSince = SecondsSince + 86400#
End Function
Public Sub RPX_LowerCacheReset()
    cacheValid = False: cacheKey = "": cacheBytes = 0#
    Erase savedCones: Erase savedQ: Erase savedEPtr: Erase savedEId
    Erase savedEVal: Erase savedERhs: Erase savedKPtr: Erase savedKId
    Erase savedKBase: Erase savedKDiag: Erase savedPerm: Erase savedInv
    Erase savedRowNorm
    Erase savedStressStart
    Set savedReactions = Nothing: Set savedEqRoles = Nothing
End Sub
Private Function Eligible(ByRef reason As String) As Boolean
    Dim profile As RPX_ProblemProfile
    If CLng(RPX_Setting("LOWER_MODEL_CACHE", 0)) <> 1 Then reason = "disabled": Exit Function
    RPX_BuildModelProfile profile
    Eligible = RPX_CanUseFeature("LOWER_MODEL_CACHE", profile, reason)
End Function
Public Function RPX_LowerExactIdentity() As String
    Dim parts() As String, e As Long, k As Long, j As Long, i As Long, n As Long
    ' Model identity includes exact scale, loads and raw/reference_total phase.
    ' Lower assembly is independent of RSubdivisions and strengthFactor.
    RPX_AssertInputs "lower_cache_identity"
    ReDim parts(0 To re + RNEdge)
    parts(n) = "lower_layout_v1|" & RPX_ExactModel(): n = n + 1
    For e = 0 To re - 1
        parts(n) = RPX_ExactDouble(RArea(e))
        For k = 0 To 2
            For j = 0 To 1: parts(n) = parts(n) & RPX_ExactDouble(RGrad(e, k, j)): Next j
        Next k
        n = n + 1
    Next e
    For i = 0 To RNEdge - 1
        With REdges(i)
            parts(n) = RPX_ExactDouble(.length)
            For j = 0 To 1
                parts(n) = parts(n) & RPX_ExactDouble(.normal(j)) & RPX_ExactDouble(.tangent(j))
                For k = 0 To 2: parts(n) = parts(n) & "," & .loc(j, k): Next k
            Next j
        End With
        n = n + 1
    Next i
    RPX_LowerExactIdentity = Join(parts, "|")
End Function
Private Sub CopyCone(ByRef dest As RPX_ConeData, ByRef source As RPX_ConeData, ByVal index As Long)
    dest.dimn = source.dimn: dest.count = source.count: dest.first = source.first
    dest.Pattern = source.Pattern: dest.modelKind = source.modelKind: dest.owner = source.owner
    dest.parameter1 = source.parameter1: dest.parameter2 = source.parameter2
    dest.directionSign = source.directionSign
    dest.ids = source.ids: dest.coef = source.coef: dest.offset = source.offset
    ' Repeated patterns deliberately have an unallocated slots array after Compile.
    If source.Pattern = index Then dest.slots = source.slots Else Erase dest.slots
End Sub
Private Function CopyReactions(ByVal source As Collection) As Collection
    Dim result As New Collection, item As Variant
    For Each item In source
        result.Add Array(item(0), item(1), item(2), item(3), item(4))
    Next item
    Set CopyReactions = result
End Function
Public Function RPX_LowerCacheMatch() As Boolean
    Dim reason As String, started As Double
    started = Timer
    RPX_AssertInputs "lower_cache_match"
    If Not Eligible(reason) Then
        Call RPX_LowerCacheReset
    ElseIf Not cacheValid Then
        reason = "empty"
    ElseIf cacheKey <> RPX_LowerExactIdentity() Then
        reason = "model_identity"
        Call RPX_LowerCacheReset
    Else
        RPX_LowerCacheMatch = True
    End If
    RPX_DiagEvent "lower_cache_lookup", "hit=" & Abs(CLng(RPX_LowerCacheMatch)) & ";reason=" & reason & ";seconds=" & SecondsSince(started)
End Function
Public Sub RPX_LowerCacheStore()
    Dim reason As String, key As String, started As Double, bytes As Double, i As Long
    RPX_AssertInputs "lower_cache_store"
    If Not Eligible(reason) Then Exit Sub
    started = Timer: key = RPX_LowerExactIdentity()
    If cacheValid Then
        If cacheKey = key Then
            RPX_DiagEvent "lower_cache_keep", "bytes=" & cacheBytes & ";seconds=" & SecondsSince(started)
            Exit Sub
        End If
    End If
    Call RPX_LowerCacheReset
    ' Conservative retained-payload estimate including array/UDT and key overhead.
    bytes = 2# * Len(key) + 8# * XN + 12# * (UBound(XEId) + 1#)
    bytes = bytes + 20# * (XM + 1#) + 16# * (XN + XM) + 4#
    bytes = bytes + 12# * (UBound(XKId) + 1#) + 4# * re + 512#
    For i = 0 To XB - 1
        With XC(i)
            bytes = bytes + 512# + 4# * .count + 8# * .dimn * (.count + 1#)
            If .Pattern = i Then bytes = bytes + 4# * .count * .count
        End With
    Next i
    bytes = bytes + 256# * RReactions.count + 256# * XM
    If bytes > CacheByteLimit Then
        RPX_DiagEvent "lower_cache_bypass", "reason=retained_budget;bytes=" & bytes & ";limit=" & CacheByteLimit
        Exit Sub
    End If
    RPX_CapacityCheck "lower_cache_store", bytes
    RPX_TestHook "lower_cache_store"
    savedN = XN: savedM = XM: savedS = XS: savedB = XB
    savedQScale = XQScale: savedLoad = RLoadVariable
    savedRowNorm = XEqRowNorm
    savedQ = XQ: savedEPtr = XEPtr: savedEId = XEId: savedEVal = XEVal: savedERhs = XERhs
    savedKPtr = XKPtr: savedKId = XKId: savedKBase = XKBase: savedKDiag = XKDiag
    savedPerm = XPerm: savedInv = XInv: savedStressStart = RStressStart
    Set savedReactions = CopyReactions(RReactions)
    Set savedEqRoles = RPX_CopyEqRoles(RLowerEqRoles)
    ReDim savedCones(0 To XB - 1)
    For i = 0 To XB - 1
        CopyCone savedCones(i), XC(i), i
        If i Mod 256 = 0 Then RPX_Stage "lower cache store " & i
    Next i
    RPX_TestHook "lower_cache_store_commit"
    If key <> RPX_LowerExactIdentity() Then err.Raise 5, "RPX_LowerCache", "LOWER_CACHE_IDENTITY_CHANGED"
    cacheKey = key: cacheBytes = bytes: cacheValid = True
    RPX_DiagEvent "lower_cache_store", "engine=20260927_p3b;bytes=" & bytes & ";seconds=" & SecondsSince(started) & ";xn=" & XN & ";xm=" & XM & ";xb=" & XB
End Sub
Public Sub RPX_LowerCacheRestore()
    Dim i As Long, started As Double
    started = Timer
    RPX_Stage "lower cache restore"
    RPX_TestHook "lower_cache_restore"
    If Not cacheValid Then err.Raise 5, "RPX_LowerCache", "LOWER_CACHE_NOT_VALID"
    If cacheKey <> RPX_LowerExactIdentity() Then err.Raise 5, "RPX_LowerCache", "LOWER_CACHE_IDENTITY_CHANGED"
    RPX_CapacityCheck "lower_cache_restore", cacheBytes
    Call RPX_LDLRelease
    Call RPX_ResetPreparedState
    XN = savedN: XM = savedM: XS = savedS: XB = savedB
    XQScale = savedQScale: RLoadVariable = savedLoad
    XEqRowNorm = savedRowNorm
    XQ = savedQ: XEPtr = savedEPtr: XEId = savedEId: XEVal = savedEVal: XERhs = savedERhs
    XKPtr = savedKPtr: XKId = savedKId: XKBase = savedKBase: XKDiag = savedKDiag
    XPerm = savedPerm: XInv = savedInv: RStressStart = savedStressStart
    Set RReactions = CopyReactions(savedReactions)
    Set RLowerEqRoles = RPX_CopyEqRoles(savedEqRoles)
    ReDim XC(0 To XB - 1)
    For i = 0 To XB - 1
        CopyCone XC(i), savedCones(i), i
        If i Mod 256 = 0 Then RPX_Stage "lower cache restore " & i
    Next i
    ReDim XKVal(0 To UBound(XKId)): ReDim XKScale(0 To XN + XM - 1)
    ' Rebuild cheap symbolic/reach data. No LDL values, Newton directions or iterates restored.
    RPX_DiagBegin dpSymbolic
    RPX_Symbolic XN + XM, XKPtr, XKId
    RPX_DiagEnd dpSymbolic
    RPX_TestHook "lower_cache_restore_commit"
    If cacheKey <> RPX_LowerExactIdentity() Then err.Raise 5, "RPX_LowerCache", "LOWER_CACHE_IDENTITY_CHANGED"
    Call RPX_DiagCounts
    RPX_DiagEvent "lower_cache_restore", "bytes=" & cacheBytes & ";seconds=" & SecondsSince(started) & ";prepared=1;cold_start=1;symbolic=rebuilt"
End Sub
