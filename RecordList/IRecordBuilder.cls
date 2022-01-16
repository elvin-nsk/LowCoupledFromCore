VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IRecordBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Interface
Option Explicit

'===============================================================================

Public Function WithField(ByVal Key As String, _
                          ByVal Value As Variant _
                          ) As IRecordBuilder
End Function

Public Function Build() As IRecord
End Function

Public Property Get Events() As IRecordBuilderEvents
End Property

'===============================================================================
