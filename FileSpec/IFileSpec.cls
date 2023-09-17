VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IFileSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Interface
Option Explicit

'===============================================================================

Public Property Get Ext() As String
End Property
Public Property Let Ext(ByVal Value As String)
End Property

Public Property Get GetCopy() As IFileSpec
End Property

'@DefaultMember
Public Function ToString() As String
Attribute ToString.VB_UserMemId = 0
End Function

Public Property Get Name() As String
End Property
Public Property Let Name(ByVal RHS As String)
End Property

Public Property Get NameWithoutExt() As String
End Property
Public Property Let NameWithoutExt(ByVal RHS As String)
End Property

Public Property Get Path() As String
End Property
Public Property Let Path(ByVal RHS As String)
End Property

'===============================================================================

Public Property Get Self() As IFileSpec
End Property
