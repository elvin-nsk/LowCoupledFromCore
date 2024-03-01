VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextBoxHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'===============================================================================
'   Модуль          : TextBoxHandler
'   Версия          : 2024.02.21
'   Автор           : elvin-nsk (me@elvin.nsk.ru)
'   Сайт            : https://github.com/elvin-nsk/LowCoupledFromCore
'   Назначение      : сохранение значений TextBox на форме в заданном диапазоне
'   Зависимости     : нет
'===============================================================================

'@PredeclaredId
Option Explicit

'===============================================================================
' # Declarations

Private Const CUSTOM_ERROR = vbObjectError Or 32

Public Enum enumTextBoxType
    TextBoxTypeLong = 3
    TextBoxTypeDouble = 5
    TextBoxTypeString = 8
End Enum

Private Type This
    TextBoxType As enumTextBoxType
    Min As Variant
    Max As Variant
    AllowMinus As Boolean
End Type
Private This As This
Private WithEvents Bound As MSForms.TextBox
Attribute Bound.VB_VarHelpID = -1

'===============================================================================
' # Constructors

Public Function New_( _
                    ByVal TextBox As MSForms.TextBox, _
                    ByVal TextBoxType As enumTextBoxType, _
                    Optional ByVal Min As Variant = -2147483647, _
                    Optional ByVal Max As Variant = 2147483647 _
                ) As TextBoxHandler
    Set New_ = New TextBoxHandler
    New_.Inject TextBox, TextBoxType, Min, Max
End Function

Friend Sub Inject( _
               ByVal TextBox As MSForms.TextBox, _
               ByVal TextBoxType As enumTextBoxType, _
               ByVal Min As Variant, _
               ByVal Max As Variant _
           )
    With This
        .TextBoxType = TextBoxType
        .Min = Min
        .Max = Max
        If .Min > .Max Then ThrowMinGreaterThanMax
        .AllowMinus = .Min < 0
    End With
    Set Bound = TextBox
End Sub

'===============================================================================
' # Public

'@DefaultMember
Public Property Get Value() As Variant
Attribute Value.VB_UserMemId = 0
    Select Case This.TextBoxType
        Case TextBoxTypeLong
            Value = GetLong
        Case TextBoxTypeDouble
            Value = GetDouble
        Case TextBoxTypeString
            Value = VBA.CStr(Bound.Value)
    End Select
End Property
Public Property Let Value(ByVal RHS As Variant)
    Bound.Value = VBA.CStr(RHS)
End Property

Public Property Get Self() As TextBoxHandler
    Set Self = Me
End Property

'===============================================================================
' # Handlers

Private Sub Bound_Change()
    '
End Sub

Private Sub Bound_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case This.TextBoxType
        Case TextBoxTypeLong
            GuardInt KeyAscii
        Case TextBoxTypeDouble
            GuardFloat KeyAscii
        Case TextBoxTypeString
    End Select
End Sub

Private Sub Bound_KeyUp( _
                ByVal KeyCode As MSForms.ReturnInteger, _
                ByVal Shift As Integer _
            )
    Select Case This.TextBoxType
        Case TextBoxTypeLong
            GuardRangeLng This.Min, This.Max
        Case TextBoxTypeDouble
            GuardRangeDbl This.Min, This.Max
        Case TextBoxTypeString
    End Select
End Sub

'===============================================================================
' # Helpers

Private Sub GuardInt(ByVal KeyAscii As MSForms.ReturnInteger)
    GuardNumeric KeyAscii, False
End Sub

Private Sub GuardFloat(ByVal KeyAscii As MSForms.ReturnInteger)
    GuardNumeric KeyAscii, True
End Sub

Private Sub GuardNumeric( _
                ByVal KeyAscii As MSForms.ReturnInteger, _
                ByVal AllowFloat As Boolean _
            )
    Select Case KeyAscii.Value
        Case VBA.Asc("0") To VBA.Asc("9")
        Case VBA.Asc(","), VBA.Asc(".")
            If FloatCharExists Then
                KeyAscii.Value = 0
            ElseIf AllowFloat And _
                         Bound.SelStart > 0 Then
                KeyAscii.Value = VBA.Asc(",")
            Else
                KeyAscii.Value = 0
            End If
        Case VBA.Asc("-")
            If This.AllowMinus Then
                If Bound.SelStart > 0 Then KeyAscii.Value = 0
            Else
                KeyAscii.Value = 0
            End If
        Case Else
            KeyAscii.Value = 0
    End Select
End Sub

Private Sub GuardRangeDbl( _
                ByVal Min As Double, _
                Optional ByVal Max As Double = 1.79769313486231E+308 _
            )
    With Bound
        If .Value = "" Or .Value = "-" Then Exit Sub
        If VBA.CDbl(.Value) > Max Then .Value = VBA.CStr(Max)
        If VBA.CDbl(.Value) < Min Then .Value = VBA.CStr(Min)
    End With
End Sub

Private Sub GuardRangeLng( _
                ByVal Min As Long, _
                Optional ByVal Max As Long = 2147483647 _
            )
    With Bound
        If .Value = "" Or .Value = "-" Then Exit Sub
        If VBA.CLng(.Value) > Max Then .Value = VBA.CStr(Max)
        If VBA.CLng(.Value) < Min Then .Value = VBA.CStr(Min)
    End With
End Sub

Private Function FloatCharExists() As Boolean
    FloatCharExists = _
        Not (VBA.InStr(1, Bound.Value, ",", vbBinaryCompare) = 0)
End Function

Private Function GetDouble() As Double
    With Bound
        If .Value = "" Or .Value = "-" Then Exit Function
        GetDouble = .Value
    End With
End Function

Private Function GetLong() As Long
    With Bound
        If .Value = "" Or .Value = "-" Then Exit Function
        GetLong = .Value
    End With
End Function

Private Sub ThrowMinGreaterThanMax()
    VBA.Err.Raise CUSTOM_ERROR, "RecordList", "MinGreaterThanMax"
End Sub
