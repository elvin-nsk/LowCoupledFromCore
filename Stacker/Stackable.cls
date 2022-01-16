VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Stackable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
Option Explicit

'===============================================================================

Private Type typeThis
  Shapes As ShapeRange
  IsShape As Boolean
  Name As String
End Type
Private This As typeThis

'===============================================================================

Public Function Create( _
                  ByVal ShapeOrRange As Object _
                ) As Stackable
  With New Stackable
    Set Create = .Init(ShapeOrRange)
  End With
End Function

Friend Function Init( _
                  ByVal ShapeOrRange As Object _
                ) As Stackable
  With This
    If TypeOf ShapeOrRange Is Shape Then
      Set .Shapes = New ShapeRange
      .Shapes.Add ShapeOrRange
      .IsShape = True
    ElseIf TypeOf ShapeOrRange Is ShapeRange Then
      Set .Shapes = ShapeOrRange
    Else
      VBA.Err.Raise 13, Source:="Stackable", _
                    Description:="Type mismatch: ShapeOrRange должен быть Shape или ShapeRange"
    End If
  End With
  Set Init = Me
End Function

'===============================================================================

Public Property Get Width() As Double
  Width = This.Shapes.SizeWidth
End Property
Public Property Let Width(ByVal RHS As Double)
  This.Shapes.SizeWidth = RHS
End Property

Public Property Get Height() As Double
  Height = This.Shapes.SizeHeight
End Property
Public Property Let Height(ByVal RHS As Double)
  This.Shapes.SizeHeight = RHS
End Property

Public Property Get PivotX() As Double
  PivotX = This.Shapes.LeftX
End Property
Public Property Let PivotX(ByVal RHS As Double)
  This.Shapes.LeftX = RHS
End Property

Public Property Get PivotY() As Double
  PivotY = This.Shapes.TopY
End Property
Public Property Let PivotY(ByVal RHS As Double)
  This.Shapes.TopY = RHS
End Property

Public Property Get Name() As String
  If This.IsShape Then
    Name = This.Shapes.FirstShape.ObjectData("Name").FormattedValue
  Else
    Name = This.Name
  End If
End Property
Public Property Let Name(ByVal RHS As String)
  If This.IsShape Then
    This.Shapes.FirstShape.Name = RHS
  Else
    This.Name = RHS
  End If
End Property

Public Property Get Self() As Stackable
  Set Self = Me
End Property

'===============================================================================
