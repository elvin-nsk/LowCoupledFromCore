VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Record"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
Option Explicit

Implements IRecord
Implements IRecordFactory
Implements IRecordBuilder

'===============================================================================

Private Const CustomError = vbObjectError Or 32

Private Type typeThis
  Encapsulated As Scripting.Dictionary
  Changed As Boolean
  BuilderEvents As IRecordBuilderEvents
  BuilderAssignedKeys As IList
End Type
Private This As typeThis

'===============================================================================

Private Function IRecordFactory_Builder _
                 (ByVal Keys As IList) As IRecordBuilder
  With New Record
    Set IRecordFactory_Builder = .Init(Keys)
  End With
End Function

Friend Function Init(ByVal Keys As IList) As IRecordBuilder
  With This
    Set .Encapsulated = New Scripting.Dictionary
    Dim Item As Variant
    For Each Item In Keys.NewEnum
      .Encapsulated.Add Item, Empty
    Next Item
    Set .BuilderEvents = New IRecordBuilderEvents
    Set .BuilderAssignedKeys = List.Create
  End With
  Set Init = Me
End Function

'-------------------------------------------------------------------------------

Private Function IRecordBuilder_WithField _
                 (ByVal Key As String, ByVal Value As Variant) As IRecordBuilder
  CheckKeyOrThrow Key
  CheckAlreadyAssigned Key
  AssignUnknownItem Key, Value
  This.BuilderAssignedKeys.Add Key
  Set IRecordBuilder_WithField = Me
End Function

Private Function IRecordBuilder_Build() As IRecord
  Set IRecordBuilder_Build = Me
  This.BuilderEvents.Built Me
  Set This.BuilderEvents = Nothing
  Set This.BuilderAssignedKeys = Nothing
End Function

Private Property Get IRecordBuilder_Events() As IRecordBuilderEvents
  Set IRecordBuilder_Events = This.BuilderEvents
End Property

'===============================================================================

Private Function IRecord_Contains(ByVal Value As Variant) As Boolean
  Dim Item As Variant
  For Each Item In This.Encapsulated.Items
    If Item = Value Then
      IRecord_Contains = True
      Exit Function
    End If
  Next Item
End Function

Private Function IRecord_ContainsLike(ByVal Pattern As String) As Boolean
  Dim Item As Variant
  For Each Item In This.Encapsulated.Items
    If Item Like Pattern Then
      IRecord_ContainsLike = True
      Exit Function
    End If
  Next Item
End Function

Private Function IRecord_Exists(ByVal Key As String) As Boolean
  IRecord_Exists = This.Encapsulated.Exists(Key)
End Function

Private Property Get IRecord_Field(ByVal Key As String) As Variant
  CheckKeyOrThrow Key
  AssignUnknown IRecord_Field, This.Encapsulated(Key)
End Property
Private Property Let IRecord_Field(ByVal Key As String, RHS As Variant)
  CheckKeyOrThrow Key
  This.Encapsulated(Key) = RHS
  This.Changed = True
End Property
Private Property Set IRecord_Field(ByVal Key As String, RHS As Variant)
  CheckKeyOrThrow Key
  Set This.Encapsulated(Key) = RHS
  This.Changed = True
End Property

Private Function IRecord_GetCopy() As IRecord
  Dim ThisFactory As IRecordFactory
  Set ThisFactory = Me
  With ThisFactory.Builder(Self.Keys)
    Dim Key As Variant
    For Each Key In This.Encapsulated.Keys
      .WithField Key, This.Encapsulated.Item(Key)
    Next Key
    Set IRecord_GetCopy = .Build
  End With
End Function

Private Property Get IRecord_Keys() As IList
  Set IRecord_Keys = List.Create
  Dim Key As Variant
  For Each Key In This.Encapsulated.Keys
    IRecord_Keys.Add Key
  Next Key
End Property

Private Property Get IRecord_IsChanged() As Boolean
  IRecord_IsChanged = This.Changed
End Property

Private Property Get IRecord_Self() As IRecord
  Set IRecord_Self = Me
End Property

'===============================================================================

Private Sub CheckAlreadyAssigned(ByVal Key As String)
  Dim Item As Variant
  For Each Item In This.BuilderAssignedKeys.NewEnum
    If Item = Key Then ThrowFieldAlreadySet Key
  Next Item
End Sub

Private Sub CheckKeyOrThrow(ByVal Key As String)
  If Not This.Encapsulated.Exists(Key) Then ThrowWrongKey
End Sub

Public Sub AssignUnknown(ByRef Destination As Variant, ByRef Value As Variant)
  If VBA.IsObject(Value) Then
    Set Destination = Value
  Else
    Destination = Value
  End If
End Sub

Public Sub AssignUnknownItem(ByVal Key As Variant, _
                             ByRef Value As Variant)
  If VBA.IsObject(Value) Then
    Set This.Encapsulated(Key) = Value
  Else
    This.Encapsulated(Key) = Value
  End If
End Sub

Private Sub ThrowFieldAlreadySet(ByVal Name As String)
  VBA.Err.Raise CustomError, "Record", "Значение поля '" & Name & "' уже задано"
End Sub

Private Sub ThrowWrongKey()
  VBA.Err.Raise CustomError, "Record", "WrongKey"
End Sub

Private Property Get Self() As IRecord
  Set Self = Me
End Property
