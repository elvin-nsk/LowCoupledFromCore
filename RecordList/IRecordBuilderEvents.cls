VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IRecordBuilderEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Events pseudo-interface
Option Explicit

'===============================================================================

Public Event Built(ByVal outRecord As IRecord)

'===============================================================================

Friend Sub Built(ByVal Record As IRecord)
  RaiseEvent Built(Record)
End Sub
