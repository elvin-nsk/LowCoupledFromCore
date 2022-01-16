VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IRecordListFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Interface
Option Explicit

'===============================================================================

Public Function Fields(ByRef PrimitiveOrEnumerable As Variant, _
                       Optional ByVal Key As String) As IRecordList
End Function

Public Function FieldsLike(ByVal Pattern As String, _
                           Optional ByVal Key As String) As IRecordList
End Function

Public Function FieldsEmpty(ByVal Key As String) As IRecordList
End Function

Public Function NotFields(ByRef PrimitiveOrEnumerable As Variant, _
                          Optional ByVal Key As String) As IRecordList
End Function

Public Function NotFieldsLike(ByVal Pattern As String, _
                              Optional ByVal Key As String) As IRecordList
End Function

Public Function NotFieldsEmpty(ByVal Key As String) As IRecordList
End Function

'===============================================================================
