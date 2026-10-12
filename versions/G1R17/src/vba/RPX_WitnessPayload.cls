Attribute VB_Name = "RPX_WitnessPayload"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
' Owned immutable payload. Publication is a single Set after Capture succeeds.
Private ready As Boolean, nodeN As Long, elementN As Long, materialN As Long, rigidN As Long
Private xy() As Double, triangles() As Long, materialId() As Long, rigidId() As Long, values() As Double
Private body0() As Double, body1() As Double
Private materials() As RPX_Material, rigids() As RPX_Rigid
Private boundary As Object, constraints As Object
Private modelKey As String, policy As String, phase As String, gravity As String
Private stressScale As Double, fs As Double, q As Long, epoch As Long
Public Sub Capture(ByRef field() As Double, Optional ByVal acceptedFs As Double = 0#, Optional ByVal acceptedQ As Long = 0)
    Dim i As Long, j As Long, key As Variant
    If ready Then err.Raise 5, "RPX_WitnessPayload", "IMMUTABLE_PAYLOAD"
    RPX_AssertInputs "payload_capture"
    nodeN = rn: elementN = re: materialN = RM: rigidN = RR
    modelKey = RPX_ExactModel(): policy = RPX_AnalysisPolicy(): phase = RPX_LoadPhase
    stressScale = RStressScale: fs = RStrengthFactor: q = RSubdivisions: epoch = RPX_InputEpoch
    If acceptedFs > 0# Then fs = acceptedFs
    If acceptedQ > 0 Then q = acceptedQ
    gravity = RGravityMode
    ReDim xy(0 To nodeN - 1, 0 To 1): ReDim triangles(0 To elementN - 1, 0 To 2)
    ReDim materialId(0 To elementN - 1): ReDim rigidId(0 To elementN - 1)
    ReDim body0(0 To elementN - 1, 0 To 1): ReDim body1(0 To elementN - 1, 0 To 1)
    ReDim materials(0 To materialN - 1): ReDim rigids(0 To rigidN)
    For i = 0 To nodeN - 1
        For j = 0 To 1: xy(i, j) = RXY(i, j): Next j
    Next i
    For i = 0 To elementN - 1
        For j = 0 To 2: triangles(i, j) = RTri(i, j): Next j
        For j = 0 To 1: body0(i, j) = RBody0(i, j): body1(i, j) = RBody1(i, j): Next j
        materialId(i) = RMatId(i): rigidId(i) = RRigidId(i)
    Next i
    For i = 0 To materialN - 1: materials(i) = RMaterials(i): Next i
    For i = 0 To rigidN - 1: rigids(i) = RRigids(i): Next i
    values = field
    For i = LBound(values) To UBound(values)
        If Not RPX_SchurFinite(values(i)) Then err.Raise 5, "RPX_WitnessPayload", "NONFINITE_WITNESS_PAYLOAD"
    Next i
    Set boundary = RPX_Row(): Set constraints = RPX_Row()
    For Each key In RBoundary.keys: boundary(key) = RBoundary(key): Next key
    For Each key In RConstraints.keys: constraints(key) = RConstraints(key): Next key
    RPX_TestHook "payload_commit"
    RPX_AssertInputs "payload_commit"
    If modelKey <> RPX_ExactModel() Then err.Raise 5, "RPX_WitnessPayload", "PAYLOAD_MODEL_MISMATCH"
    ready = True
End Sub
Public Function node(ByVal i As Long, ByVal j As Long) As Double
    node = xy(i, j)
End Function
Public Function Triangle(ByVal i As Long, ByVal j As Long) As Long
    Triangle = triangles(i, j)
End Function
Public Function Material(ByVal i As Long) As Long
    Material = materialId(i)
End Function
Public Function Rigid(ByVal i As Long) As Long
    Rigid = rigidId(i)
End Function
Public Function field(ByVal i As Long) As Double
    field = values(i)
End Function
Public Sub RestoreUpper()
    Dim key As Variant
    If Not ready Then err.Raise 5, "RPX_WitnessPayload", "PAYLOAD_NOT_READY"
    RPX_AssertInputs "payload_restore"
    If epoch <> RPX_InputEpoch Or policy <> RPX_AnalysisPolicy() Then err.Raise 5, "RPX_WitnessPayload", "PAYLOAD_INPUT_MISMATCH"
    rn = nodeN: re = elementN: RM = materialN: RR = rigidN
    RXY = xy: RTri = triangles: RMatId = materialId: RRigidId = rigidId
    RBody0 = body0: RBody1 = body1: RMaterials = materials: RRigids = rigids
    RStressScale = stressScale: RStrengthFactor = fs: RSubdivisions = q
    RGravityMode = gravity: RPX_LoadPhase = phase
    Set RBoundary = RPX_Row(): Set RConstraints = RPX_Row()
    For Each key In boundary.keys: RBoundary(key) = boundary(key): Next key
    For Each key In constraints.keys: RConstraints(key) = constraints(key): Next key
    RUpperField = values
    Call RPX_Topology
    If modelKey <> RPX_ExactModel() Then err.Raise 5, "RPX_WitnessPayload", "PAYLOAD_RESTORE_MISMATCH"
    RProgramKind = "": Call RPX_ResetPreparedState
End Sub
