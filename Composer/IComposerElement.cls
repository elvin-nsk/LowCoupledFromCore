VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IComposerElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Interface
'@ModuleDescription("Interface")

Option Explicit

'===============================================================================

Public Property Get Height() As Double
End Property

Public Property Get PivotX() As Double
End Property
Public Property Let PivotX(ByVal RHS As Double)
End Property

Public Property Get PivotY() As Double
End Property
Public Property Let PivotY(ByVal RHS As Double)
End Property

Public Property Get Shapes() As ShapeRange
End Property

Public Property Get Width() As Double
End Property

'===============================================================================

Public Property Get Self() As IComposerElement
End Property
