VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Either"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
Option Explicit

Implements IEither

'===============================================================================

Private Type typeThis
    SuccessValue As Variant
    ErrorValue As Variant
    IsSuccess As Boolean
    IsError As Boolean
End Type
Private This As typeThis

'===============================================================================

Public Function Create( _
                    Optional ByRef SuccessValue As Variant, _
                    Optional ByRef ErrorValue As Variant _
                ) As IEither
    With New Either
        Set Create = .Init(SuccessValue, ErrorValue)
    End With
End Function

Friend Function Init( _
                    ByRef SuccessValue As Variant, _
                    ByRef ErrorValue As Variant _
                ) As IEither
    With This
        AssignUnknown .SuccessValue, SuccessValue
        AssignUnknown .ErrorValue, ErrorValue
        If VBA.IsMissing(.SuccessValue) Then
            .IsError = True
            If VBA.IsMissing(.ErrorValue) Then .ErrorValue = 0
        Else
            .IsSuccess = True
        End If
    End With
    Set Init = Me
End Function

'===============================================================================

Private Property Get IEither_ErrorValue() As Variant
    AssignUnknown IEither_ErrorValue, This.ErrorValue
End Property

Private Property Get IEither_IsError() As Boolean
    IEither_IsError = This.IsError
End Property

Private Property Get IEither_IsSuccess() As Boolean
    IEither_IsSuccess = This.IsSuccess
End Property

Private Property Get IEither_Self() As IEither
    Set IEither_Self = Me
End Property

Private Property Get IEither_SuccessValue() As Variant
    AssignUnknown IEither_SuccessValue, This.SuccessValue
End Property

'===============================================================================

Private Sub AssignUnknown( _
                ByRef Destination As Variant, _
                ByRef Value As Variant _
            )
    If VBA.IsObject(Value) Then
        Set Destination = Value
    Else
        Destination = Value
    End If
End Sub
