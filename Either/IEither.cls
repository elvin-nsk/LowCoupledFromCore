VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IEither"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Interface
Option Explicit

'===============================================================================

Public Property Get IsError() As Boolean
End Property

Public Property Get IsSuccess() As Boolean
End Property

Public Property Get ErrorValue() As Variant
End Property

Public Property Get SuccessValue() As Variant
End Property

'===============================================================================

Public Property Get Self() As IEither
End Property
