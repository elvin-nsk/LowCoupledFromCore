VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Point"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'===============================================================================
'   Модуль          : Point
'   Версия          : 2024.02.13
'   Автор           : elvin-nsk (me@elvin.nsk.ru)
'   Сайт            : https://github.com/elvin-nsk/LowCoupledFromCore
'   Назначение      : Хранит координаты точки
'   Зависимости     : Нет
'===============================================================================

'@Interface
'@PredeclaredId
Option Explicit

'===============================================================================

Private Type This
    x As Double
    y As Double
End Type
Private This As This

'===============================================================================

Friend Function New_(ByVal x As Double, ByVal y As Double) As Point
    Set New_ = New Point
    New_.x = x
    New_.y = y
End Function

'===============================================================================

Public Function GetCopy() As Point
    Set GetCopy = New_(This.x, This.y)
End Function

'todo
Public Function GetDistanceFrom(ByVal Point As Point) As Double
End Function

Public Sub Move(ByVal DeltaX As Double, ByVal DeltaY As Double)
    With This
        .x = .x + DeltaX
        .y = .y + DeltaY
    End With
End Sub

Public Sub RotateAroundPoint( _
               ByVal RotationCenter As Point, _
               ByVal Angle As Double _
           )
    With RotationCenter
        This.x = _
            (This.x - .x) * VBA.Cos(DegToRad(Angle)) _
          - (This.y - .y) * VBA.Sin(DegToRad(Angle)) + .x
        This.y = _
            (This.x - .x) * VBA.Sin(DegToRad(Angle)) _
          + (This.y - .y) * VBA.Cos(DegToRad(Angle)) + .y
    End With
End Sub

Public Property Get Self() As Point
    Set Self = Me
End Property

Public Property Get x() As Double
    x = This.x
End Property
Public Property Let x(ByVal RHS As Double)
    This.x = RHS
End Property

Public Property Get y() As Double
    y = This.y
End Property
Public Property Let y(ByVal RHS As Double)
    This.y = RHS
End Property

'===============================================================================

Private Function DegToRad(Degree As Double) As Double
    DegToRad = Degree * 3.14159265358979 / 180
End Function
