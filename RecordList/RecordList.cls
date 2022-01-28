VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RecordList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
Option Explicit

Implements IRecordList
Implements IRecordListFilter

'===============================================================================

Private Const CustomError = vbObjectError Or 32

Private Type typeThis
  Records As IList
  Keys As IList
  PrimaryKey As String
  IsKeyFieldSet As Boolean
  RecordFactory As IRecordFactory
End Type
Private This As typeThis
Private WithEvents PendingRecord As IRecordBuilderEvents
Attribute PendingRecord.VB_VarHelpID = -1

'===============================================================================

Public Function Create(ByVal Keys As IList, _
                       Optional ByVal PrimaryKey As String, _
                       Optional ByVal RecordFactory As IRecordFactory _
                       ) As IRecordList
  With New RecordList
    Set Create = .Init( _
      Records:=List.Create, _
      Keys:=GetKeysAndCheck(Keys), _
      PrimaryKey:=PrimaryKey, _
      RecordFactory:=GetRecordFactory(RecordFactory) _
      )
  End With
End Function

Friend Function CreateClone(ByVal Records As IList, _
                            ByVal Keys As IList, _
                            ByVal PrimaryKey As String, _
                            ByVal RecordFactory As IRecordFactory _
                            ) As IRecordList
  With New RecordList
    Set CreateClone = .Init( _
      Records:=Records, _
      Keys:=Keys, _
      PrimaryKey:=PrimaryKey, _
      RecordFactory:=RecordFactory _
      )
  End With
End Function

Friend Function Init(ByVal Records As IList, _
                     ByVal Keys As IList, _
                     ByVal PrimaryKey As String, _
                     ByVal RecordFactory As IRecordFactory _
                     ) As IRecordList
  With This
    Set .Records = Records
    Set .Keys = Keys
    CheckAndSetKeyFieldName PrimaryKey
    Set .RecordFactory = RecordFactory
  End With
  Set Init = Me
End Function

Private Sub CheckAndSetKeyFieldName(ByVal PrimaryKey As String)
  With This
    If PrimaryKey = "" Then Exit Sub
    .IsKeyFieldSet = True
    .PrimaryKey = PrimaryKey
  End With
End Sub

Private Function GetKeysAndCheck(ByVal Keys As IList) As IList
  Set GetKeysAndCheck = List.Create
  Dim Item As Variant
  For Each Item In Keys.NewEnum
    If GetKeysAndCheck.Contains(Item) Then
      ThrowDuplicateFieldName
    Else
      GetKeysAndCheck.Add Item
    End If
  Next Item
End Function

Private Function GetRecordFactory _
                 (ByVal RecordFactory As IRecordFactory) As IRecordFactory
  If RecordFactory Is Nothing Then
    Set GetRecordFactory = Record
  Else
    Set GetRecordFactory = RecordFactory
  End If
End Function

'===============================================================================

Private Function IRecordList_BuildRecord() As IRecordBuilder
  Set IRecordList_BuildRecord = This.RecordFactory.Builder(This.Keys)
  Set PendingRecord = IRecordList_BuildRecord.Events
End Function

Private Property Get IRecordList_Count() As Long
  IRecordList_Count = This.Records.Count
End Property

Private Function IRecordList_GetCopy() As IRecordList
  With This
    Dim RecordsCopies As IList
    Set RecordsCopies = List.Create
    Dim Item As IRecord
    For Each Item In .Records.NewEnum
      RecordsCopies.Add Item.GetCopy
    Next Item
    Set IRecordList_GetCopy = _
      CreateClone(RecordsCopies, .Keys, .PrimaryKey, .RecordFactory)
  End With
End Function

Private Function IRecordList_KeyExists(ByVal Key As String) As Boolean
  IRecordList_KeyExists = This.Keys.Contains(Key)
End Function

Private Function IRecordList_Filter() As IRecordListFilter
  Set IRecordList_Filter = Me
End Function

Private Function IRecordList_GetFields(ByVal Name As String) As IList
  Set IRecordList_GetFields = List.Create
  Dim Record As IRecord
  For Each Record In This.Records.NewEnum
    IRecordList_GetFields.Add Record.Field(Name)
  Next Record
End Function

Private Function IRecordList_PrimaryFieldExists(ByVal Value As Variant) As Boolean
  If Not IsKeyFieldValid(Value) Then ThrowKeyFieldNotPrimitive
  Dim Record As IRecord
  For Each Record In This.Records.NewEnum
    If Record.Field(This.PrimaryKey) = Value Then
      IRecordList_PrimaryFieldExists = True
      Exit Function
    End If
  Next Record
End Function

Private Property Get IRecordList_PrimaryKeySet() As Boolean
  IRecordList_PrimaryKeySet = This.IsKeyFieldSet
End Property

Private Property Get IRecordList_NewEnum() As EnumHelper
  Set IRecordList_NewEnum = This.Records.NewEnum
End Property

Private Property Get IRecordList_Record(ByVal NumberOrPrimaryField As Variant) As IRecord
  Dim Index As Long
  Index = FindRecordIndex(NumberOrPrimaryField)
  If Index > 0 Then
    Set IRecordList_Record = This.Records(Index)
  Else
    ThrowInvalidRecordNumberOrName
  End If
End Property

Private Function IRecordList_RecordExists _
                 (ByVal NumberOrPrimaryField As Variant) As Boolean
  IRecordList_RecordExists = Not (FindRecordIndex(NumberOrPrimaryField) = 0)
End Function

Private Property Get IRecordList_Self() As IRecordList
  Set IRecordList_Self = Me
End Property

'-------------------------------------------------------------------------------

Private Function IRecordListFilter_Fields _
                 (ByRef PrimitiveOrEnumerable As Variant, _
                  Optional ByVal Key As String) As IRecordList
  Set IRecordListFilter_Fields = CreateClone _
    (RecordsMatchUnknown(PrimitiveOrEnumerable, Key, True), _
     This.Keys, This.PrimaryKey, This.RecordFactory)
End Function

Private Function IRecordListFilter_FieldsEmpty _
                 (ByVal Key As String) As IRecordList
  Dim Records As IList
  Set Records = List.Create
  Dim Record As IRecord
  For Each Record In This.Records.NewEnum
    If Record.Field(Key) = "" Then Records.Add Record
  Next Record
  Set IRecordListFilter_FieldsEmpty = CreateClone _
    (Records, This.Keys, This.PrimaryKey, This.RecordFactory)
End Function

Private Function IRecordListFilter_FieldsLike _
                 (ByVal Pattern As String, _
                  Optional ByVal Key As String) As IRecordList
  Dim Records As IList
  Set Records = List.Create
  Dim Record As IRecord
  For Each Record In This.Records.NewEnum
    If IsRecordLike(Record, Pattern, Key) Then Records.Add Record
  Next Record
  Set IRecordListFilter_FieldsLike = CreateClone _
    (Records, This.Keys, This.PrimaryKey, This.RecordFactory)
End Function

Private Function IRecordListFilter_NotFields _
                 (ByRef PrimitiveOrEnumerable As Variant, _
                  Optional ByVal Key As String) As IRecordList
 Set IRecordListFilter_NotFields = CreateClone _
    (RecordsMatchUnknown(PrimitiveOrEnumerable, Key, False), _
     This.Keys, This.PrimaryKey, This.RecordFactory)
End Function

Private Function IRecordListFilter_NotFieldsEmpty _
                 (ByVal Key As String) As IRecordList
  Dim Records As IList
  Set Records = List.Create
  Dim Record As IRecord
  For Each Record In This.Records.NewEnum
    If Not Record.Field(Key) = "" Then Records.Add Record
  Next Record
  Set IRecordListFilter_NotFieldsEmpty = CreateClone _
    (Records, This.Keys, This.PrimaryKey, This.RecordFactory)
End Function

Private Function IRecordListFilter_NotFieldsLike _
                 (ByVal Pattern As String, _
                  Optional ByVal Key As String) As IRecordList
  Dim Records As IList
  Set Records = List.Create
  Dim Record As IRecord
  For Each Record In This.Records.NewEnum
    If Not IsRecordLike(Record, Pattern, Key) Then Records.Add Record
  Next Record
  Set IRecordListFilter_NotFieldsLike = CreateClone _
    (Records, This.Keys, This.PrimaryKey, This.RecordFactory)
End Function

'===============================================================================

Private Sub PendingRecord_Built(ByVal Record As IRecord)
  CheckRecordKeyFieldOrThrow Record
  This.Records.Add Record
  Set PendingRecord = Nothing
End Sub

'===============================================================================

Private Function IsKeyFieldValid(ByVal Value As Variant) As Boolean
  IsKeyFieldValid = VBA.IsNumeric(Value) Or _
                    VBA.VarType(Value) = vbString
End Function

Private Sub CheckRecordKeyFieldOrThrow(ByVal Record As IRecord)
  With This
    If Not .IsKeyFieldSet Then Exit Sub
    If Not IsKeyFieldValid(Record.Field(.PrimaryKey)) Then _
      ThrowKeyFieldNotPrimitive
    If Self.PrimaryFieldExists(Record.Field(.PrimaryKey)) Then _
      ThrowKeyFieldDuplicate
  End With
End Sub

Private Function FindRecordIndex(ByVal NumberOrPrimaryField As Variant) As Long
  If This.Records.Count = 0 Then Exit Function
  If VBA.VarType(NumberOrPrimaryField) = vbString Then
    If Not This.IsKeyFieldSet Then ThrowKeyFieldsAreNotSet
    Dim Index As Long
    For Index = 1 To This.Records.Count
      If This.Records(Index).Field(This.PrimaryKey) = NumberOrPrimaryField Then
        FindRecordIndex = Index
        Exit Function
      End If
    Next Index
  ElseIf VBA.IsNumeric(NumberOrPrimaryField) Then
    If This.Records.Count >= VBA.CLng(NumberOrPrimaryField) Then
      FindRecordIndex = VBA.CLng(NumberOrPrimaryField)
    End If
  End If
End Function

Private Function RecordsMatchUnknown(ByRef PrimitiveOrEnumerable As Variant, _
                                     ByVal Key As String, _
                                     ByVal Match As Boolean _
                                     ) As IList
  Set RecordsMatchUnknown = List.Create
  Dim Record As IRecord
  For Each Record In This.Records.NewEnum
    If RecordMatchUnknown(Record, PrimitiveOrEnumerable, Key) = Match Then
      RecordsMatchUnknown.Add Record
    End If
  Next Record
End Function

Private Function RecordMatchUnknown(ByVal Record As IRecord, _
                                    ByRef PrimitiveOrEnumerable As Variant, _
                                    ByVal Key As String _
                                    ) As Boolean
  Dim Item As Variant
  If VBA.IsArray(PrimitiveOrEnumerable) Or _
     TypeOf PrimitiveOrEnumerable Is VBA.Collection Then
    For Each Item In PrimitiveOrEnumerable
      If IsRecordMatchValue(Record, Item, Key) Then
        RecordMatchUnknown = True
        Exit Function
      End If
    Next Item
  ElseIf TypeOf PrimitiveOrEnumerable Is Scripting.Dictionary Then
    For Each Item In PrimitiveOrEnumerable.Items
      If IsRecordMatchValue(Record, Item, Key) Then
        RecordMatchUnknown = True
        Exit Function
      End If
    Next Item
  ElseIf TypeOf PrimitiveOrEnumerable Is IList Then
    For Each Item In PrimitiveOrEnumerable.NewEnum
      If IsRecordMatchValue(Record, Item, Key) Then
        RecordMatchUnknown = True
        Exit Function
      End If
    Next Item
  Else
    RecordMatchUnknown = IsRecordMatchValue(Record, PrimitiveOrEnumerable, Key)
  End If
End Function

Private Function IsRecordMatchValue(ByVal Record As IRecord, _
                                    ByVal Value As Variant, _
                                    ByVal Key As String _
                                    ) As Boolean
  If Key = "" Then
    IsRecordMatchValue = Record.Contains(Value)
  Else
    IsRecordMatchValue = (Record.Field(Key) = Value)
  End If
End Function

Private Function IsRecordLike(ByVal Record As IRecord, _
                              ByVal Pattern As String, _
                              ByVal Key As String _
                              ) As Boolean
  If Key = "" Then
    IsRecordLike = Record.ContainsLike(Pattern)
  Else
    IsRecordLike = (Record.Field(Key) Like Pattern)
  End If
End Function

Private Sub ThrowDuplicateFieldName()
  VBA.Err.Raise CustomError, "RecordList", "DuplicateFieldName"
End Sub

Private Sub ThrowKeyFieldsAreNotSet()
  VBA.Err.Raise CustomError, "RecordList", "KeyFieldsAreNotSet"
End Sub

Private Sub ThrowInvalidRecordNumberOrName()
  VBA.Err.Raise CustomError, "RecordList", "InvalidRecordNumberOrName"
End Sub

Private Sub ThrowKeyFieldNotPrimitive()
  VBA.Err.Raise CustomError, "RecordList", "KeyFieldNotPrimitive"
End Sub

Private Sub ThrowKeyFieldDuplicate()
  VBA.Err.Raise CustomError, "RecordList", "KeyFieldDuplicate"
End Sub

Private Property Get Self() As IRecordList
  Set Self = Me
End Property
