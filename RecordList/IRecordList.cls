VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IRecordList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Interface
Option Explicit

'===============================================================================

'@DefaultMember
Public Property Get Record(ByVal NumberOrPrimaryField As Variant) As IRecord
Attribute Record.VB_UserMemId = 0
End Property

Public Function BuildRecord() As IRecordBuilder
End Function

Public Function GetFields(ByVal Name As String) As IList
End Function

Public Property Get Count() As Long
End Property

Public Function RecordExists(ByVal NumberOrPrimaryField As Variant) As Boolean
End Function

Public Function KeyExists(ByVal Key As String) As Boolean
End Function

Public Function PrimaryFieldExists(ByVal Value As Variant) As Boolean
End Function

Public Property Get PrimaryKeySet() As Boolean
End Property

Public Function Filter() As IRecordListFilter
End Function

Public Function GetCopy() As IRecordList
End Function

Public Property Get NewEnum() As EnumHelper
End Property

'===============================================================================

Public Property Get Self() As IRecordList
End Property
