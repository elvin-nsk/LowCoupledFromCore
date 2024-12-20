VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NodeData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
Option Explicit

'===============================================================================
' # Declarations

Private Type This
    Node As Node
End Type
Private This As This

'===============================================================================
' # Constructors

Friend Function New_(Node As Node) As NodeData
    Set New_ = New NodeData
    New_.Inject Node
End Function

Friend Sub Inject( _
               Node As Node _
           )
    With This
        Set .Node = Node
    End With
End Sub

'===============================================================================
' # Public

Public Property Get Angle1() As Double
    If This.Node.IsEnding Then Exit Property
    Angle1 = This.Node.Segment.EndingControlPointAngle
End Property

Public Property Get Angle2() As Double
    If This.Node.IsEnding Then Exit Property
    Angle2 = This.Node.NextSegment.StartingControlPointAngle
End Property

Public Property Get ControlPoint1() As Point
    If This.Node.Segment.Type = cdrLineSegment Then
        Set ControlPoint1 = Position
        Exit Property
    End If
    Set ControlPoint1 = _
        Point.New_( _
            This.Node.Segment.EndingControlPointX, _
            This.Node.Segment.EndingControlPointY _
        )
End Property

Public Property Get ControlPoint2() As Point
    If This.Node.NextSegment.Type = cdrLineSegment Then
        Set ControlPoint2 = Position
        Exit Property
    End If
    Set ControlPoint2 = _
        Point.New_( _
            This.Node.NextSegment.StartingControlPointX, _
            This.Node.NextSegment.StartingControlPointY _
        )
End Property

Public Property Get Position() As Point
    Set Position = Point.New_(This.Node.PositionX, This.Node.PositionY)
End Property

Public Sub SampleAction(ByVal Value As Variant)
End Sub

Public Function SampleFunction(ByVal Value As Variant) As Variant
End Function

Public Property Get Self() As NodeData
    Set Self = Me
End Property

Public Property Get Unwrap() As Node
    Set Unwrap = This.Node
End Property

'===============================================================================
' # Helpers
