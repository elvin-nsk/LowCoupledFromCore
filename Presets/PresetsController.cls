VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PresetsController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
Option Explicit

'===============================================================================
' # Declarations

Private Type typeThis
    PresetsForm As MSForms.UserForm
    BoundDictionary As Scripting.IDictionary
    CurrentKey As String
    LastSelectedIndex As Long
    ProtectedKey As String
    WarningPresetNameAlreadyExists As String
    WarningPresetNameIsEmpty As String
    WarningPresetRemove As String
End Type
Private This As typeThis

Private WithEvents PresetsList As MSForms.ListBox
Attribute PresetsList.VB_VarHelpID = -1
Private WithEvents CurrentTextBox As MSForms.TextBox
Attribute CurrentTextBox.VB_VarHelpID = -1
Private WithEvents AddButton As MSForms.CommandButton
Attribute AddButton.VB_VarHelpID = -1
Private WithEvents RemoveButton As MSForms.CommandButton
Attribute RemoveButton.VB_VarHelpID = -1

'===============================================================================
' # Events

Event OnSelectionChange(ByRef outNewKey As String)

'===============================================================================
' # Constructors

Public Function Create( _
                    ByVal PresetsForm As MSForms.UserForm, _
                    ByVal BoundDictionary As Scripting.IDictionary, _
                    Optional ByVal CurrentKey As String _
                ) As PresetsController
    With New PresetsController
        Set Create = .Init(PresetsForm, BoundDictionary, CurrentKey)
    End With
End Function

Friend Function Init( _
                    ByVal PresetsForm As MSForms.UserForm, _
                    ByVal BoundDictionary As Scripting.IDictionary, _
                    ByVal CurrentKey As String _
                ) As PresetsController
    With This
        Set .PresetsForm = PresetsForm
        Set .BoundDictionary = BoundDictionary
        
        Set PresetsList = .PresetsForm("PresetsList")
        Set CurrentTextBox = .PresetsForm("PresetCurrent")
        Set AddButton = .PresetsForm("PresetAdd")
        Set RemoveButton = .PresetsForm("PresetRemove")
        .WarningPresetNameAlreadyExists = "Пресет с таким названием уже существует"
        .WarningPresetNameIsEmpty = "Название пресета не может быть пустым"
        .WarningPresetRemove = "Удалить пресет?"
        RebuildList
        If CurrentKey = vbNullString Then
            SelectIndex 0
        Else
            SelectIndex GetListItemIndex(CurrentKey)
        End If
    End With
    
    Set Init = Me
End Function

'===============================================================================
' # Public

Public Property Get CurrentItem() As Variant
    If Not This.BoundDictionary.Exists(This.CurrentKey) Then Exit Property
    Assign CurrentItem, This.BoundDictionary(This.CurrentKey)
End Property

Public Property Get CurrentKey() As String
    CurrentKey = This.CurrentKey
End Property
Public Property Let CurrentKey(RHS As String)
    This.CurrentKey = CurrentKey
End Property

Public Sub Refresh()
    RebuildList
    If PresetsList.ListCount = 0 Then Exit Sub
    If This.LastSelectedIndex <= GetListLastIndex Then
        SelectIndex This.LastSelectedIndex
    Else
        SelectIndex 0
    End If
End Sub

Public Sub SetProtectedKey(ByVal Key As String)
    This.ProtectedKey = Key
End Sub

Public Sub SetWarningPresetNameAlreadyExists(ByVal Message As String)
    This.WarningPresetNameAlreadyExists = Message
End Sub

Public Sub SetWarningPresetNameIsEmpty(ByVal Message As String)
    This.WarningPresetNameIsEmpty = Message
End Sub

Public Sub SetWarningPresetRemove(ByVal Message As String)
    This.WarningPresetRemove = Message
End Sub

Public Property Get Self() As PresetsController
    Set Self = Me
End Property

'===============================================================================
' # Handlers

Private Sub PresetsList_Click()
    OnChange
End Sub

Private Sub PresetsList_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    OnChange
End Sub

Private Sub AddButton_Click()
    If CurrentTextBox.Value = vbNullString Then
        Warning This.WarningPresetNameIsEmpty
        Exit Sub
    End If
    If This.BoundDictionary.Exists(CurrentTextBox.Value) Then
        Warning This.WarningPresetNameAlreadyExists
        Exit Sub
    End If
    This.BoundDictionary.Add CurrentTextBox.Value, CurrentItem
    RebuildList
    SelectIndex GetListLastIndex
End Sub

Private Sub RemoveButton_Click()
    If Not AskYesNo(This.WarningPresetRemove) Then Exit Sub
    Dim Key As String
    Key = PresetsList.List(This.LastSelectedIndex)
    This.BoundDictionary.Remove Key
    Refresh
    CheckButtonsState
End Sub

'===============================================================================
' # Helpers

Private Sub RebuildList()
    PresetsList.Clear
    Dim Key As Variant
    For Each Key In This.BoundDictionary.Keys
        PresetsList.AddItem Key
    Next Key
End Sub

Private Sub SelectIndex(ByVal Index As Long)
    If Index < 0 Then Exit Sub
    PresetsList.Selected(Index) = True
    This.LastSelectedIndex = Index
    CurrentTextBox.Value = This.CurrentKey
    CheckButtonsState
End Sub

Private Property Get GetListLastIndex() As Long
    GetListLastIndex = PresetsList.ListCount - 1
End Property

Private Property Get GetListItemIndex(ByVal Key As String) As Long
    GetListItemIndex = -1
    If PresetsList.ListCount = 0 Then Exit Property
    Dim Index As Long
    For Index = 0 To PresetsList.ListCount - 1
        If PresetsList.List(Index) = Key Then
            GetListItemIndex = Index
            Exit Property
        End If
    Next Index
End Property

Private Sub OnChange()
    If Not SelectionChanged Then Exit Sub
    SyncDictionaryToList
    CurrentTextBox.Value = This.CurrentKey
    CheckButtonsState
    RaiseEvent OnSelectionChange(This.CurrentKey)
End Sub

Private Sub SyncDictionaryToList()
    Dim Index As Long
    Index = PresetsList.ListIndex
    If Index = -1 Then Exit Sub
    This.CurrentKey = This.BoundDictionary.Keys(Index)
End Sub

Private Sub CheckButtonsState()
    If Protected(This.CurrentKey) _
    Or PresetsList.ListIndex = -1 _
    Or PresetsList.ListCount = 0 Then
        RemoveButton.Enabled = False
    Else
        RemoveButton.Enabled = True
    End If
End Sub

Private Property Get SelectionChanged() As Boolean
    If Not This.LastSelectedIndex = PresetsList.ListIndex Then
        This.LastSelectedIndex = PresetsList.ListIndex
        SelectionChanged = True
    End If
End Property

Private Property Get Protected(ByVal Key As String) As Boolean
    If This.ProtectedKey = vbNullString Then Exit Property
    Protected = (Key = This.ProtectedKey)
End Property

Private Sub Warning(ByVal Text As String)
    VBA.MsgBox Text, vbExclamation
End Sub

Private Function AskYesNo(ByVal Text As String) As Boolean
    AskYesNo = (VBA.MsgBox(Text, vbYesNo) = 6)
End Function
