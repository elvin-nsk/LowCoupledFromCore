VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@ModuleDescription "Low-coupled version from Core framework"
'@Interface
Option Explicit

'===============================================================================

Public Property Get X() As Double
End Property
Public Property Let X(ByVal RHS As Double)
End Property

Public Property Get Y() As Double
End Property
Public Property Let Y(ByVal RHS As Double)
End Property

Public Function GetDistanceFrom(ByVal Point As IPoint) As Double
End Function

Public Sub Move(ByVal DeltaX As Double, ByVal DeltaY As Double)
End Sub

Public Sub RotateAroundPoint(ByVal RotationCenter As IPoint, ByVal Angle As Double)
End Sub

'===============================================================================

Public Function GetCopy() As IPoint
End Function

Public Property Get Self() As IPoint
End Property
