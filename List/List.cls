VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@ModuleDescription "Low-coupled version from Core framework"
'@PredeclaredId
Option Explicit

Implements IList

'===============================================================================

Private Const CustomError = vbObjectError Or 32

Private Type typeThis
  Encapsulated As VBA.Collection
End Type
Private This As typeThis

'===============================================================================

Public Function Create() As IList
  With New List
    Set Create = .Init()
  End With
End Function

Public Function CreateFrom(ParamArray Items() As Variant) As IList
  With New List
    Set CreateFrom = .Init()
    Dim Item As Variant
    For Each Item In Items
      CreateFrom.Add Item
    Next Item
  End With
End Function

Friend Function Init() As IList
  With This
    Set .Encapsulated = New VBA.Collection
  End With
  Set Init = Me
End Function

'===============================================================================

Private Sub IList_Add(ByRef Item As Variant, _
                      Optional ByVal Before As Long, _
                      Optional ByVal After As Long)
  If Before = 0 And (Not After = 0) Then
    This.Encapsulated.Add Item, , , After
  ElseIf (Not Before = 0) And After = 0 Then
    This.Encapsulated.Add Item, , Before
  ElseIf (Not Before = 0) And (Not After = 0) Then
    ThrowBeforeAndAfterAreMutuallyExclusive
  Else
    This.Encapsulated.Add Item
  End If
End Sub

Private Sub IList_Append(ByVal List As IList, _
                         Optional ByVal Before As Long, _
                         Optional ByVal After As Long)
  If List.Count = 0 Then Exit Sub
  Dim i As Long
  For i = 1 To List.Count
    Self.Add List(i), Before, After
  Next i
End Sub

Private Sub IList_AppendCollection(ByVal Collection As VBA.Collection, _
                                   Optional ByVal Before As Long, _
                                   Optional ByVal After As Long)
  If Collection.Count = 0 Then Exit Sub
  Dim i As Long
  For i = 1 To Collection.Count
    Self.Add Collection(i), Before, After
  Next i
End Sub

Private Function IList_Contains(ByRef Item As Variant) As Boolean
  IList_Contains = (Self.IndexOf(Item) > 0)
End Function

Private Property Get IList_Count() As Long
  IList_Count = This.Encapsulated.Count
End Property

Private Function IList_GetCopy() As IList
  With This
    Set IList_GetCopy = Create
    If .Encapsulated.Count = 0 Then Exit Function
    Dim i As Long
    For i = 1 To .Encapsulated.Count
      IList_GetCopy.Add .Encapsulated.Item(i)
    Next i
  End With
End Function

Private Function IList_IndexOf(ByRef Item As Variant) As Long
  Dim i As Long
  For i = 1 To This.Encapsulated.Count
    If IsSame(This.Encapsulated(i), Item) Then
      IList_IndexOf = i
      Exit Function
    End If
  Next i
End Function

Private Function IList_Intersect(ByVal List As IList) As IList
'todo
End Function

Private Property Set IList_Item(ByVal Index As Long, ByVal RHS As Variant)
  UpdateItem Index, RHS
End Property
Private Property Let IList_Item(ByVal Index As Long, RHS As Variant)
  UpdateItem Index, RHS
End Property
Private Property Get IList_Item(ByVal Index As Long) As Variant
  AssignUnknown IList_Item, This.Encapsulated.Item(Index)
End Property

Private Property Get IList_NewEnum() As EnumHelper
  With New EnumHelper
    Set .EnumVariant = This.Encapsulated.[_NewEnum]
    Set IList_NewEnum = .Self
  End With
End Property

Private Sub IList_Remove(ByVal Index As Long)
  This.Encapsulated.Remove Index
End Sub

Private Sub IList_RemoveList(ByVal List As IList)
  Dim i As Long
  For i = 1 To List.Count
    If Self.Contains(List(i)) Then Self.Remove Self.IndexOf(List(i))
  Next i
End Sub

Private Property Get IList_Self() As IList
  Set IList_Self = Me
End Property

'===============================================================================

Private Sub UpdateItem(ByVal Index As Long, ByRef Item As Variant)
  With This.Encapsulated
    .Add Item, , , Index
    .Remove Index
  End With
End Sub

Private Sub AssignUnknown(ByRef Destination As Variant, ByRef Value As Variant)
  If VBA.IsObject(Value) Then
    Set Destination = Value
  Else
    Destination = Value
  End If
End Sub

Private Function IsSame(ByRef Value1 As Variant, ByRef Value2 As Variant) As Boolean
  If VBA.IsObject(Value1) And VBA.IsObject(Value2) Then
    IsSame = Value1 Is Value2
  ElseIf Not VBA.IsObject(Value1) And Not VBA.IsObject(Value2) Then
    IsSame = (Value1 = Value2)
  End If
End Function

Private Sub ThrowBeforeAndAfterAreMutuallyExclusive()
  VBA.Err.Raise CustomError, "List", "BeforeAndAfterAreMutuallyExclusive"
End Sub

Private Property Get Self() As IList
  Set Self = Me
End Property
