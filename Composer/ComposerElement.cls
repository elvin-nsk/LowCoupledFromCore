VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComposerElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
Option Explicit

Implements IComposerElement

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
                ) As ComposerElement
    With New ComposerElement
        Set Create = .Init(ShapeOrRange)
    End With
End Function

Friend Function Init( _
                    ByVal ShapeOrRange As Object _
                ) As ComposerElement
    With This
        If TypeOf ShapeOrRange Is Shape Then
            Set .Shapes = New ShapeRange
            .Shapes.Add ShapeOrRange
            .IsShape = True
        ElseIf TypeOf ShapeOrRange Is ShapeRange Then
            Set .Shapes = ShapeOrRange
        Else
            VBA.Err.Raise _
                13, Source:="ComposerElement", _
                Description:="Type mismatch: ShapeOrRange должен быть Shape или ShapeRange"
        End If
        End With
    Set Init = Me
End Function

'===============================================================================

Private Property Get IComposerElement_Height() As Double
    IComposerElement_Height = This.Shapes.SizeHeight
End Property

Private Property Let IComposerElement_PivotX(ByVal RHS As Double)
    This.Shapes.LeftX = RHS
End Property
Private Property Get IComposerElement_PivotX() As Double
    IComposerElement_PivotX = This.Shapes.LeftX
End Property

Private Property Let IComposerElement_PivotY(ByVal RHS As Double)
    This.Shapes.TopY = RHS
End Property
Private Property Get IComposerElement_PivotY() As Double
    IComposerElement_PivotY = This.Shapes.TopY
End Property

Private Property Get IComposerElement_Width() As Double
    IComposerElement_Width = This.Shapes.SizeWidth
End Property

Private Property Get IComposerElement_Shapes() As ShapeRange
    Set IComposerElement_Shapes = This.Shapes
End Property

Private Property Get IComposerElement_Self() As IComposerElement
    Set IComposerElement_Self = Me
End Property

'===============================================================================
