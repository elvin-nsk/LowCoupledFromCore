VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Surface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Struct
Option Explicit

'===============================================================================

Public Content As Shape
Public File As IFileSpec

'===============================================================================

Public Property Get GetCopy() As Surface
    With New Surface
        Set .Content = Content
        Set .File = File.GetCopy
        Set GetCopy = .Self
    End With
End Property

Public Property Get Self() As Surface
    Set Self = Me
End Property
