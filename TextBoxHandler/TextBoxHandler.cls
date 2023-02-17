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
'   Версия          : 2022.04.23
'   Автор           : elvin-nsk (me@elvin.nsk.ru)
'   Сайт            : https://github.com/elvin-nsk/LowCoupledFromCore
'   Назначение      : сохранение значений TextBox на форме в заданном диапазоне
'   Зависимости     : нет
'===============================================================================

'@PredeclaredId
Option Explicit

'===============================================================================

Private Const CustomError = vbObjectError Or 32

Public Enum enumTextBoxType
    TextBoxTypeLong = 3
    TextBoxTypeDouble = 5
    TextBoxTypeString = 8
End Enum

Private Type typeThis
    TextBoxType As enumTextBoxType
    Min As Variant
    Max As Variant
    AllowMinus As Boolean
End Type
Private This As typeThis
Private WithEvents Encapsulated As MSForms.TextBox
Attribute Encapsulated.VB_VarHelpID = -1

'===============================================================================

Public Function Create( _
                    ByVal TextBox As MSForms.TextBox, _
                    ByVal TextBoxType As enumTextBoxType, _
                    Optional ByVal Min As Variant = 0, _
                    Optional ByVal Max As Variant = 2147483647 _
                ) As TextBoxHandler
    With New TextBoxHandler
        Set Create = .Init(TextBox, TextBoxType, Min, Max)
    End With
End Function

Friend Function Init( _
                    ByVal TextBox As MSForms.TextBox, _
                    ByVal TextBoxType As enumTextBoxType, _
                    ByVal Min As Variant, _
                    ByVal Max As Variant _
                ) As TextBoxHandler
    With This
        .TextBoxType = TextBoxType
        .Min = Min
        .Max = Max
        If .Min > .Max Then ThrowMinGreaterThanMax
        .AllowMinus = .Min < 0
    End With
    Set Encapsulated = TextBox
    Set Init = Me
End Function

'===============================================================================

'@DefaultMember
Public Property Get Value() As Variant
Attribute Value.VB_UserMemId = 0
    Select Case This.TextBoxType
        Case TextBoxTypeLong
            Value = GetLong
        Case TextBoxTypeDouble
            Value = GetDouble
        Case TextBoxTypeString
            Value = VBA.CStr(Encapsulated.Value)
    End Select
End Property
Public Property Let Value(ByVal RHS As Variant)
    Encapsulated.Value = VBA.CStr(RHS)
End Property

Public Property Get Self() As TextBoxHandler
    Set Self = Me
End Property

'===============================================================================

Private Sub Encapsulated_Change()
    '
End Sub

Private Sub Encapsulated_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case This.TextBoxType
        Case TextBoxTypeLong
            GuardInt KeyAscii
        Case TextBoxTypeDouble
            GuardFloat KeyAscii
        Case TextBoxTypeString
    End Select
End Sub

Private Sub Encapsulated_KeyUp( _
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
                         Encapsulated.SelStart > 0 Then
                KeyAscii.Value = VBA.Asc(",")
            Else
                KeyAscii.Value = 0
            End If
        Case VBA.Asc("-")
            If This.AllowMinus Then
                If Encapsulated.SelStart > 0 Then KeyAscii.Value = 0
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
    With Encapsulated
        If .Value = "" Or .Value = "-" Then Exit Sub
        If VBA.CDbl(.Value) > Max Then .Value = VBA.CStr(Max)
        If VBA.CDbl(.Value) < Min Then .Value = VBA.CStr(Min)
    End With
End Sub

Private Sub GuardRangeLng( _
                ByVal Min As Long, _
                Optional ByVal Max As Long = 2147483647 _
            )
    With Encapsulated
        If .Value = "" Or .Value = "-" Then Exit Sub
        If VBA.CLng(.Value) > Max Then .Value = VBA.CStr(Max)
        If VBA.CLng(.Value) < Min Then .Value = VBA.CStr(Min)
    End With
End Sub

Private Function FloatCharExists() As Boolean
    FloatCharExists = Not (VBA.InStr(1, Encapsulated.Value, ",", vbBinaryCompare) = 0)
End Function

Private Function GetDouble() As Double
    With Encapsulated
        If .Value = "" Or .Value = "-" Then Exit Function
        GetDouble = .Value
    End With
End Function

Private Function GetLong() As Long
    With Encapsulated
        If .Value = "" Or .Value = "-" Then Exit Function
        GetLong = .Value
    End With
End Function

Private Sub ThrowMinGreaterThanMax()
    VBA.Err.Raise CustomError, "RecordList", "MinGreaterThanMax"
End Sub
