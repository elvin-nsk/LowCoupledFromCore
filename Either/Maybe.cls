VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Maybe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
'@ModuleDescription("Option: Some or None.")

Option Explicit

'===============================================================================
' # Declarations

Private Type This
    Some As Variant
    None As Boolean
End Type
Private This As This

'===============================================================================
' # Constructors

Public Property Get New_(ByVal x As Variant) As Maybe
    If TypeOf x Is Maybe Then
        Set New_ = x
        Exit Property
    End If
    Set New_ = New Maybe
    New_.Inject x
End Property

Friend Sub Inject( _
               ByVal x As Variant _
           )
    With This
        .None = IsNone(x)
        If Not .None Then Assign .Some, x
    End With
End Sub

'===============================================================================
' # Public

Public Property Get None() As Boolean
    None = This.None
End Property

Public Function Some(Optional ByRef Out As Variant) As Boolean
    Some = Not This.None
    If Not This.None Then Assign Out, This.Some
End Function

Public Property Get Self() As Maybe
    Set Self = Me
End Property

Public Property Get Unwrap() As Variant
    If This.None Then
        VBA.Err.Raise CUSTOM_ERROR, "Maybe.Unwrap", "Maybe -> None"
    Else
        Assign Unwrap, This.Some
    End If
End Property

'===============================================================================
' # Interfaces

'-------------------------------------------------------------------------------


'===============================================================================
' # Helpers


