VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Size"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'===============================================================================
'   Модуль          : Size
'   Версия          : 2024.02.27
'   Автор           : elvin-nsk (me@elvin.nsk.ru)
'   Сайт            : https://github.com/elvin-nsk/LowCoupledFromCore
'   Назначение      : хранит и обрабатывает абстрактный размер
'   Зависимости     : LibCore
'===============================================================================

'@Interface
'@PredeclaredId
Option Explicit

'===============================================================================
' # Declarations

Private Type This
    Height As Double
    Width As Double
End Type
Private This As This

Private Const DEFAULT_DECIMAL_SEPARATOR = ","

'===============================================================================
' # Constructors

Friend Function New_( _
                    ByVal Width As Double, _
                    ByVal Height As Double _
                ) As Size
    Set New_ = New Size
    New_.Inject Width, Height
End Function

Friend Function NewFromRect( _
                    ByVal Rect As Rect _
                ) As Size
    Set NewFromRect = New Size
    NewFromRect.Inject Rect.Width, Rect.Height
End Function

Friend Function NewFromShape( _
                    ByVal Shape As Shape _
                ) As Size
    Set NewFromShape = New Size
    NewFromShape.Inject Shape.SizeWidth, Shape.SizeHeight
End Function

Friend Sub Inject( _
               ByVal Width As Double, _
               ByVal Height As Double _
           )
    With This
        .Width = Width
        .Height = Height
    End With
End Sub

'===============================================================================
' # Public

Public Sub ApplyToRect(ByVal Rect As Rect)
    Rect.Width = This.Width
    Rect.Height = This.Height
End Sub

Public Sub ApplyToShape(ByVal Shape As Shape)
    Shape.SetSize This.Width, This.Height
End Sub

Public Property Get ConvertUnits( _
                        ByVal FromUnit As cdrUnit, _
                        ByVal ToUnit As cdrUnit _
                    ) As Size
    If FromUnit = ToUnit Then
        Set ConvertUnits = Me
    Else
        Set ConvertUnits = _
            New_( _
                Application.ConvertUnits(This.Width, FromUnit, ToUnit), _
                Application.ConvertUnits(This.Height, FromUnit, ToUnit) _
            )
    End If
End Property

Public Property Get DisplayHeight( _
                        Optional ByVal NumDigitsAfterDecimal As Long, _
                        Optional ByVal DecimalSeparator As String = _
                            DEFAULT_DECIMAL_SEPARATOR _
                    ) As String
    DisplayHeight = _
        DisplaySize(This.Height, NumDigitsAfterDecimal, DecimalSeparator)
End Property

Public Property Get DisplayWidth( _
                        Optional ByVal NumDigitsAfterDecimal As Long, _
                        Optional ByVal DecimalSeparator As String = _
                            DEFAULT_DECIMAL_SEPARATOR _
                    ) As String
    DisplayWidth = _
        DisplaySize(This.Width, NumDigitsAfterDecimal, DecimalSeparator)
End Property

Public Property Get Equilateral(Optional ByVal Tolerance As Double) As Boolean
    Equilateral = _
        LibCore.DiffWithinTolerance(This.Width, This.Height, Tolerance)
End Property

Public Property Get Height() As Double
    Height = This.Height
End Property

Public Property Get Landscape() As Boolean
    Landscape = This.Width > This.Height
End Property

Public Property Get Longest() As Double
    If This.Width > This.Height Then
        Longest = This.Width
    Else
        Longest = This.Height
    End If
End Property

Public Property Get Mult(ByVal Multiplier As Double) As Size
    Set Mult = New_(This.Width * Multiplier, This.Height * Multiplier)
End Property

Public Property Get Portrait() As Boolean
    Portrait = This.Width < This.Height
End Property

Public Property Get ResizeToHeight(ByVal NewHeight As Double) As Size
    Set ResizeToHeight = New_(NewHeight * Ratio, NewHeight)
End Property

Public Property Get ResizeToLongest(ByVal NewLongest As Double) As Size
    If Landscape Then
        Set ResizeToLongest = ResizeToWidth(NewLongest)
    Else
        Set ResizeToLongest = ResizeToHeight(NewLongest)
    End If
End Property

Public Property Get ResizeToShortest(ByVal NewShortest As Double) As Size
    If Landscape Then
        Set ResizeToShortest = ResizeToHeight(NewShortest)
    Else
        Set ResizeToShortest = ResizeToWidth(NewShortest)
    End If
End Property

Public Property Get ResizeToWidth(ByVal NewWidth As Double) As Size
    Set ResizeToWidth = New_(NewWidth, NewWidth / Ratio)
End Property

Public Property Get Shortest() As Double
    If This.Width < This.Height Then
        Shortest = This.Width
    Else
        Shortest = This.Height
    End If
End Property

Public Property Get Swap() As Size
    Set Swap = New_(This.Height, This.Width)
End Property

Public Property Get Width() As Double
    Width = This.Width
End Property

Public Property Get Self() As Size
    Set Self = Me
End Property

'===============================================================================
' # Helpers

Private Property Get DisplaySize( _
                        ByVal Size As Double, _
                        ByVal NumDigitsAfterDecimal As Long, _
                        ByVal DecimalSeparator As String _
                    ) As String
    DisplaySize = _
        LibCore.ToStr(VBA.Round(Size, NumDigitsAfterDecimal), DecimalSeparator)
End Property

Private Property Get Ratio() As Double
    Ratio = This.Width / This.Height
End Property
