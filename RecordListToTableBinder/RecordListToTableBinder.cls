VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RecordListToTableBinder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
Option Explicit

Implements IRecordListToTableBinder
Implements IRecordListToTableBinderBuilder

'===============================================================================

Private Enum enumAction
  ActionCreateAndRead
  ActionRead
  ActionWrite
End Enum

Private Type typeThis
  RecordList As IRecordList
  Table As ITableFile
  Bindings As Scripting.Dictionary
  RecordsToRowsCollation As IList
  PrimaryKey As String
  MandatoryKeys As IList
  IsKeyFieldSet As Boolean
  Keys As IList
  RecordFactory As IRecordFactory
End Type
Private This As typeThis

'===============================================================================

Public Function Builder(ByVal Table As ITableFile, _
                        Optional ByVal RecordFactory As IRecordFactory _
                        ) As IRecordListToTableBinderBuilder
  With New RecordListToTableBinder
    Set Builder = .Init(Table, RecordFactory)
  End With
End Function

Private Sub Class_Terminate()
  With This
    If .RecordList Is Nothing Or _
       .Table Is Nothing Then
      Exit Sub
    End If
    If Not .Table.ReadOnly Then
      TableWalk ActionWrite
    End If
  End With
End Sub

Friend Function Init(ByVal Table As ITableFile, _
                     ByVal RecordFactory As IRecordFactory _
                     ) As IRecordListToTableBinderBuilder
  With This
    Set .Table = Table
    Set .Bindings = New Scripting.Dictionary
    Set .RecordsToRowsCollation = List.Create
    Set .RecordFactory = RecordFactory
    Set .Keys = List.Create
    Set .MandatoryKeys = List.Create
  End With
  Set Init = Me
End Function

'-------------------------------------------------------------------------------

Private Function IRecordListToTableBinderBuilder_WithKey _
                 (ByVal Key As String, ByVal ColumnToBind As Long _
                 ) As IRecordListToTableBinderBuilder
  This.Bindings.Add Key, ColumnToBind
  This.Keys.Add Key
End Function

Private Function IRecordListToTableBinderBuilder_WithMandatoryKey _
                 (ByVal Key As String, ByVal ColumnToBind As Long _
                 ) As IRecordListToTableBinderBuilder
  This.Bindings.Add Key, ColumnToBind
  This.Keys.Add Key
  This.MandatoryKeys.Add Key
End Function

Private Function IRecordListToTableBinderBuilder_WithOptionalKey _
                 (ByVal Key As String, ByVal ColumnToBind As Long _
                 ) As IRecordListToTableBinderBuilder
  If ColumnExists(ColumnToBind) Then This.Bindings.Add Key, ColumnToBind
  This.Keys.Add Key
End Function

Private Function IRecordListToTableBinderBuilder_WithPrimaryKey _
                 (ByVal Key As String, ByVal ColumnToBind As Long _
                 ) As IRecordListToTableBinderBuilder
  If Not This.PrimaryKey = "" Then ThrowPrimaryKeyAlreadySet
  This.Bindings.Add Key, ColumnToBind
  This.Keys.Add Key
  This.PrimaryKey = Key
End Function


Private Function IRecordListToTableBinderBuilder_WithUnboundKey _
                 (ByVal Key As String) As IRecordListToTableBinderBuilder
  This.Keys.Add Key
End Function

Private Function IRecordListToTableBinderBuilder_Build() As IRecordListToTableBinder
  With This
    .IsKeyFieldSet = Not (.PrimaryKey = "")
    Set .RecordList = RecordList.Create(.Keys, .PrimaryKey, .RecordFactory)
    TableWalk ActionCreateAndRead
    Set IRecordListToTableBinderBuilder_Build = Me
    Set .Keys = Nothing
    Set .RecordFactory = Nothing
  End With
End Function

'===============================================================================

Private Property Get IRecordListToTableBinder_RecordList() As IRecordList
  Set IRecordListToTableBinder_RecordList = This.RecordList
End Property

Private Property Get IRecordListToTableBinder_Self() As IRecordListToTableBinder
  Set IRecordListToTableBinder_Self = Me
End Property

'===============================================================================

Private Sub TableWalk(ByVal Action As enumAction)
  Dim Row As Long
  For Row = 1 To This.Table.MaxRow
    ProcessRow Row, Action
  Next Row
End Sub

Private Sub ProcessRow(ByVal Row As Long, ByVal Action As enumAction)
  
  Dim RecordBuilder As IRecordBuilder
  Dim RecordIndex As Long
  
  If Action = ActionCreateAndRead Then
    If Not IsRowValidToRecord(Row) Then Exit Sub
    Set RecordBuilder = This.RecordList.BuildRecord
    This.RecordsToRowsCollation.Add Row
  ElseIf Action = ActionRead Then
    RecordIndex = FindRecordIndex(Row)
    If RecordIndex = 0 Then Exit Sub
  ElseIf Action = ActionWrite Then
    RecordIndex = FindRecordIndex(Row)
    If RecordIndex = 0 Then Exit Sub
    If Not This.RecordList(RecordIndex).IsChanged Then Exit Sub
  End If
  
  Dim Key As Variant
  For Each Key In This.Bindings
    
    If Action = ActionCreateAndRead Then
      RecordBuilder.WithField Key, This.Table.Cell(Row, This.Bindings(Key))
    ElseIf Action = ActionRead Then
      This.RecordList(RecordIndex)(Key) = This.Table.Cell(Row, This.Bindings(Key))
    ElseIf Action = ActionWrite Then
      This.Table.Cell(Row, This.Bindings(Key)) = This.RecordList(RecordIndex)(Key)
    End If
  
  Next Key
  
  If Action = ActionCreateAndRead Then RecordBuilder.Build
  
End Sub

Private Function IsRowValidToRecord(ByVal Row As Long) As Boolean
  With This
    If .IsKeyFieldSet Then
      If .Table.Cell(Row, .Bindings(.PrimaryKey)) = "" Then Exit Function
    End If
    If IsMandatoryFieldsEmpty(Row) Then Exit Function
    IsRowValidToRecord = True
  End With
End Function

Private Function IsMandatoryFieldsEmpty(ByVal Row As Long) As Boolean
  With This
    Dim Key As Variant
    For Each Key In .MandatoryKeys.NewEnum
      If .Table.Cell(Row, .Bindings(Key)) = "" Then
        IsMandatoryFieldsEmpty = True
        Exit Function
      End If
    Next Key
  End With
End Function

Private Function FindRecordIndex(ByVal Row As Long) As Long
  If This.RecordsToRowsCollation.Count = 0 Then Exit Function
  Dim i As Long
  For i = 1 To This.RecordsToRowsCollation.Count
    If This.RecordsToRowsCollation(i) = Row Then
      FindRecordIndex = i
      Exit Function
    End If
  Next i
End Function

Private Function ColumnExists(ByVal Column As Long) As Boolean
  On Error GoTo ErrorHandler
  Dim Temp As Variant
  Temp = This.Table.Cell(1, Column)
  ColumnExists = True
  Exit Function
ErrorHandler:
  On Error GoTo 0
End Function

Private Sub ThrowPrimaryKeyAlreadySet()
  VBA.Err.Raise CustomError, "RecordListToTableBinder", "PrimaryKeyAlreadySet"
End Sub

Private Property Get Self() As IRecordListToTableBinder
  Set Self = Me
End Property
