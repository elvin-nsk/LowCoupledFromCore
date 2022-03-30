VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IStringLocalizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Interface
Option Explicit

'===============================================================================

Public Sub ForceLocale(ByVal LCID As Variant)
End Sub

'@DefaultMember
Public Property Get Item(ByVal Key As String, _
                                                 ParamArray Args() As Variant _
                                                 ) As String
Attribute Item.VB_UserMemId = 0
End Property

Public Property Get CurrentLocale() As String
End Property

'===============================================================================

Public Property Get Self() As IStringLocalizer
End Property
