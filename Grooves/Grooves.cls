VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Grooves"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'===============================================================================
' # Fields

Public Shapes As ShapeRange
Public Punches As ShapeRange
Public GrooveSize As Double
Public PunchLength As Double
Public PunchColor As Color
Public PunchName As String
Public ProbeRadius As Double
Public ProbeSteps As Long
Public ConcavityMult As Double

'===============================================================================
' # Defaults

Private Sub Class_Initialize()
    Set Punches = CreateShapeRange
    GrooveSize = 4
    PunchLength = GrooveSize * 4
    Set PunchColor = CreateGrayColor(0)
    PunchName = "PUNCH"
    ProbeRadius = GrooveSize / 10
    ProbeSteps = 36
    ConcavityMult = 0.5
End Sub

'===============================================================================
' # Public

'строит пуансоны
Public Sub MakePunches()
    Dim Shape As Shape
    For Each Shape In Shapes
        MakePunchesOnShape Shape:=Shape, Cut:=False
    Next Shape
End Sub

'строит пуансоны и вырезает
Public Sub MakeGrooves()
    Dim Shape As Shape
    For Each Shape In Shapes
        MakePunchesOnShape Shape:=Shape, Cut:=True
    Next Shape
End Sub

'===============================================================================
' # Helpers

Private Sub MakePunchesOnShape(ByVal Shape As Shape, ByVal Cut As Boolean)
    Dim ShapePunches As New ShapeRange
    Dim Punch As Shape
    Dim Node As Node
    For Each Node In Shape.Curve.Nodes.All
        Set Punch = MakePunchOnNode(Node)
        If IsSome(Punch) Then ShapePunches.Add Punch
    Next Node
    If Cut Then
        CutShape ShapePunches, Shape
    Else
        Punches.AddRange ShapePunches
    End If
End Sub

Private Function MakePunchOnNode(ByVal Node As Node) As Shape
    
    If IsNodeConvex(Node) Then Exit Function
    
    Dim Angle As Double: Angle = AngleOutside(Node)
    Dim Punch As Shape: Set Punch = MakeGroovePunch
    With Punch
        .LeftX = Node.PositionX - GrooveSize / 2
        .TopY = Node.PositionY + GrooveSize / 2
        .Rotate AngleOutside(Node)
    End With
    Set MakePunchOnNode = Punch
End Function

Private Function MakeGroovePunch() As Shape
    Set MakeGroovePunch = _
        MakePunch( _
            PunchLength, GrooveSize, _
            OutlineColor:=PunchColor, _
            Name:=PunchName _
        )
End Function

Private Property Get AngleOutside(ByVal Node As Node) As Double
    With NodeData.New_(Node)
        AngleOutside = (.Angle1 + .Angle2) / 2
        Dim ProbePoint As Point: Set ProbePoint = _
            Probe(.Position, AngleOutside, ProbeRadius)
        If Not Node.Parent.IsPointInside(ProbePoint.x, ProbePoint.y) Then _
            Exit Property
        AngleOutside = AngleOutside + 180
    End With
End Property

Private Property Get IsNodeConvex(ByVal Node As Node) As Boolean
    Dim MaxHits As Long: MaxHits = ProbeSteps
    Dim Hits As Long
    Hits = _
        ProbeHits( _
            Node.Parent, _
            Point.New_(Node.PositionX, Node.PositionY), _
            ProbeRadius, _
            MaxHits _
        )
    If Hits < MaxHits * ConcavityMult Then IsNodeConvex = True
End Property

Private Sub MakeVectors(ByVal Node As Node)
    Dim n As NodeData: Set n = NodeData.New_(Node)
    With Probe(n.Position, 120, PunchLength)
        ActiveLayer.CreateEllipse2(.x, .y, GrooveSize / 2).Fill.ApplyUniformFill _
            CreateCMYKColor(100, 0, 0, 0)
    End With
End Sub

Private Property Get Probe( _
                         ByVal StartingPoint As Point, _
                         ByVal Angle As Double, _
                         ByVal Length As Double _
                     ) As Point
    Set Probe = StartingPoint.GetCopy
    Probe.Move Length
    Probe.RotateAroundPoint StartingPoint, Angle
End Property

Public Function MakePunch( _
                    ByVal Width As Double, _
                    ByVal Diameter As Double, _
                    Optional ByVal FillColor As Color, _
                    Optional ByVal OutlineColor As Color, _
                    Optional ByVal Name As String _
                ) As Shape
    Set MakePunch = _
        ActiveLayer.CreateRectangle2(0, 0, Width, Diameter)
    With MakePunch
        With .Rectangle
            .CornerType = cdrCornerTypeRound
            '.SetRoundness 100 не работает, поэтому так
            .CornerLowerLeft = 100
            .CornerLowerRight = 100
            .CornerUpperLeft = 100
            .CornerUpperRight = 100
        End With
        .RotationCenterX = .LeftX + Diameter / 2
        .RotationCenterY = .BottomY + Diameter / 2
        If IsSome(FillColor) Then .Fill.ApplyUniformFill FillColor
        If IsSome(OutlineColor) Then .Outline.Color.CopyAssign OutlineColor
    End With
    If Not Name = vbNullString Then MakePunch.Name = Name
End Function

Public Property Get ProbeHits( _
                        ByVal ClosedCurve As Curve, _
                        ByVal Center As Point, _
                        ByVal Radius As Double, _
                        ByVal ProbeSteps As Long _
                    ) As Long
    Dim Step As Double: Step = 360 / ProbeSteps
    Dim Probe As Point: Set Probe = _
        Point.New_(Center.x + Radius, Center.y)
    Dim Angle As Double
    For Angle = Step To 360 Step Step
        Probe.RotateAroundPoint Center, Step
        If Probe.Inside(ClosedCurve) Then ProbeHits = ProbeHits + 1
    Next Angle
End Property

Private Function CutShape( _
                     ByVal Punches As ShapeRange, _
                     ByVal Shape As Shape _
                 )
    If Punches.Count > 0 And IsSome(Shape) Then
        Punches.Combine.Trim Shape, False, False
    End If
End Function
