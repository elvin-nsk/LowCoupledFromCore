VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PresetsConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'===============================================================================
'   Модуль          : PresetsConfig
'   Версия          : 2024.06.16
'   Автор           : elvin-nsk (me@elvin.nsk.ru)
'   Сайт            : https://github.com/elvin-nsk/LowCoupledFromCore
'   Назначение      : Хранит настройки в json-формате в виде пресетов
'   Зависимости     : LibCore, JsonConverter
'===============================================================================

'@PredeclaredId
Option Explicit

'===============================================================================
' # Declarations

Private Const CURRENT_ID_ROOT_KEY As String = "Current"
Private Const DEFAULT_ID As String = "Default"
Private Const PRESETS_DICTIONARY_KEY As String = "Presets"

Private Type This
    RootDictionary As Scripting.IDictionary
    CurrentID As String
    DefaultID As String
    DefaultPreset As Scripting.IDictionary
    FileSpec As String
    ReadOnly As Boolean
End Type
Private This As This

'===============================================================================
' # Events

Event OnCurrentIDChange(ByRef OldID As String, ByRef NewID As String)
Event OnItemChange(ByRef PresetID As String, ByRef ItemKey As String)

'===============================================================================
' # Constructors

Public Function New_( _
                    ByVal FileBaseName As String, _
                    Optional ByVal DefaultPreset As Scripting.IDictionary, _
                    Optional ByVal DefaultID As String = DEFAULT_ID _
                ) As PresetsConfig
    Set New_ = New PresetsConfig
    New_.Inject FileBaseName, DefaultPreset, DefaultID, False
End Function

Friend Function NewReadOnly( _
                    ByVal FileBaseName As String _
                ) As PresetsConfig
    Set NewReadOnly = New PresetsConfig
    NewReadOnly.Inject FileBaseName, Nothing, vbNullString, True
End Function

Friend Sub Inject( _
               ByVal FileBaseName As String, _
               ByVal DefaultPreset As Scripting.IDictionary, _
               ByVal DefaultID As String, _
               ByVal ReadOnly As Boolean _
           )
    With This
        .ReadOnly = ReadOnly
        .DefaultID = DefaultID
        Set .DefaultPreset = DefaultPreset
        InitRootDictonary FileBaseName
        .CurrentID = .RootDictionary(CURRENT_ID_ROOT_KEY)
    End With
End Sub

Private Sub Class_Terminate()
    If Not This.FileSpec = vbNullString Then
        ForceSave
    End If
End Sub

'===============================================================================
' # Public

Public Property Get Current() As Scripting.IDictionary
    FillWithDefaultIfEmpty This.CurrentID
    Set Current = Presets(This.CurrentID)
End Property

Public Property Get CurrentID() As String
    CurrentID = This.CurrentID
End Property
Public Property Let CurrentID(RHS As String)
    Dim OldID As String: OldID = This.CurrentID
    This.CurrentID = RHS
    This.RootDictionary(CURRENT_ID_ROOT_KEY) = This.CurrentID
    RaiseEvent OnCurrentIDChange(OldID, RHS)
End Property

Public Property Get DefaultID() As String
    DefaultID = This.DefaultID
End Property

Public Sub ForceSave()
    If This.ReadOnly Then Exit Sub
    WriteFile _
        JsonConverter.New_.ConvertToJson(This.RootDictionary), This.FileSpec
End Sub

'@DefaultMember
Public Property Get Item(ByVal Key As String) As Variant
    FillWithDefaultIfEmpty This.CurrentID
    Item = Presets(This.CurrentID)(Key)
End Property
Public Property Let Item(ByVal Key As String, RHS As Variant)
Attribute Item.VB_UserMemId = 0
    FillWithDefaultIfEmpty This.CurrentID
    Presets(This.CurrentID)(Key) = RHS
    RaiseEvent OnItemChange(This.CurrentID, Key)
End Property

Public Property Get Root() As Scripting.IDictionary
    Set Root = This.RootDictionary
End Property

Public Property Get Presets() As Scripting.IDictionary
    Set Presets = This.RootDictionary(PRESETS_DICTIONARY_KEY)
End Property

Public Property Get Self() As PresetsConfig
    Set Self = Me
End Property

'===============================================================================
' # Helpers

Private Sub InitRootDictonary( _
                 ByVal FileBaseName As String _
            )
    This.FileSpec = GetConfigFile(FileBaseName)
    If FileExists(This.FileSpec) Then
        Set This.RootDictionary = _
            JsonConverter.New_.ParseJson(ReadFile(This.FileSpec))
    Else
        CreateRoot
    End If
End Sub

Private Sub CreateRoot()
    Set This.RootDictionary = New Scripting.Dictionary
    With This.RootDictionary
        .Add CURRENT_ID_ROOT_KEY, This.DefaultID
        .Add PRESETS_DICTIONARY_KEY, New Scripting.Dictionary
        FillWithDefaultIfEmpty This.DefaultID
    End With
End Sub

Private Sub FillWithDefaultIfEmpty(ByVal ID As String)
    If TypeOf Presets(ID) Is Scripting.IDictionary Then Exit Sub
    If This.DefaultPreset Is Nothing Then
        Set Presets(ID) = New Scripting.Dictionary
    Else
        Set Presets(ID) = GetDictionaryCopy(This.DefaultPreset)
    End If
End Sub

Private Function GetConfigFile(ByVal FileBaseName As String) As String
    GetConfigFile = GMSManager.UserGMSPath & FileBaseName & ".json"
End Function
