VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Motif"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Struct
Option Explicit

'===============================================================================

Public Quantity As Long
Public SurfaceA As Surface
Public SurfaceB As Surface

'===============================================================================

Public Property Get HasSurfaceA() As Boolean
    HasSurfaceA = Not SurfaceA Is Nothing
End Property

Public Property Get HasSurfaceB() As Boolean
    HasSurfaceB = Not SurfaceB Is Nothing
End Property

Public Property Get GetCopy() As Motif
    With New Motif
        Set .SurfaceA = SurfaceA.GetCopy
        Set .SurfaceB = SurfaceB.GetCopy
        .Quantity = Quantity
        Set GetCopy = .Self
    End With
End Property

Public Property Get Self() As Motif
    Set Self = Me
End Property
