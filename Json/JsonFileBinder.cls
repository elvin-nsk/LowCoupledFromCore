VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JsonFileBinder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'===============================================================================
'   Модуль          : JsonFileBinder
'   Версия          : 2024.06.20
'   Автор           : elvin-nsk (me@elvin.nsk.ru)
'   Сайт            : https://github.com/elvin-nsk/LowCoupledFromCore
'   Назначение      : Хранит настройки в словаре в файле json
'   Зависимости     : LibCore, JsonConverter
'===============================================================================

'@PredeclaredId
Option Explicit

'===============================================================================
' # Declarations

Private Type This
    Dictionary As Scripting.IDictionary
    FileSpec As String
    ReadOnly As Boolean
End Type
Private This As This

'===============================================================================
' # Constructors

Public Function New_( _
                    ByVal FileBaseName As String _
                ) As JsonFileBinder
    Set New_ = New JsonFileBinder
    New_.Inject FileBaseName, False
End Function

Friend Function NewReadOnly( _
                    ByVal FileBaseName As String _
                ) As JsonFileBinder
    Set NewReadOnly = New JsonFileBinder
    NewReadOnly.Inject FileBaseName, True
End Function

Friend Sub Inject( _
               ByVal FileBaseName As String, _
               ByVal ReadOnly As Boolean _
           )
    With This
        .ReadOnly = ReadOnly
        BindDictonary FileBaseName
    End With
End Sub

Private Sub Class_Terminate()
    If Not This.FileSpec = vbNullString Then ForceSave
End Sub

'===============================================================================
' # Public

Public Property Get Content() As Scripting.IDictionary
    Set Content = This.Dictionary
End Property

Public Sub ForceSave()
    If This.ReadOnly Then Exit Sub
    WriteFile _
        JsonConverter.New_.ConvertToJson(This.Dictionary), This.FileSpec
End Sub

Public Property Get GetOrMakeSubArray( _
                        ByVal Key As String _
                    ) As Variant
    With This.Dictionary
        If Not .Exists(Key) Then .Item(Key) = Pack()
        GetOrMakeSubArray = .Item(Key)
    End With
End Property

Public Property Get GetOrMakeSubCollection( _
                        ByVal Key As String _
                    ) As VBA.Collection
    With This.Dictionary
        If Not .Exists(Key) Then Set .Item(Key) = New VBA.Collection
        Set GetOrMakeSubCollection = .Item(Key)
    End With
End Property

Public Property Get GetOrMakeSubDictionary( _
                        ByVal Key As String _
                    ) As Scripting.IDictionary
    With This.Dictionary
        If Not .Exists(Key) Then Set .Item(Key) = New Scripting.Dictionary
        Set GetOrMakeSubDictionary = .Item(Key)
    End With
End Property

'@DefaultMember
Public Property Get Item(ByVal Key As String) As Variant
    If VBA.IsObject(This.Dictionary(Key)) Then
        Set Item = This.Dictionary(Key)
    Else
        Item = This.Dictionary(Key)
    End If
End Property
Public Property Let Item(ByVal Key As String, RHS As Variant)
Attribute Item.VB_UserMemId = 0
    If VBA.IsObject(RHS) Then
        Set This.Dictionary(Key) = RHS
    Else
        This.Dictionary(Key) = RHS
    End If
End Property

Public Property Get Self() As JsonFileBinder
    Set Self = Me
End Property

'===============================================================================
' # Helpers

Private Sub BindDictonary( _
                 ByVal FileBaseName As String _
            )
    This.FileSpec = GetConfigFile(FileBaseName)
    If FileExists(This.FileSpec) Then
        Set This.Dictionary = _
            JsonConverter.New_.ParseJson(ReadFile(This.FileSpec))
    Else
        Set This.Dictionary = New Scripting.Dictionary
    End If
End Sub

Private Function GetConfigFile(ByVal FileBaseName As String) As String
    GetConfigFile = GMSManager.UserGMSPath & FileBaseName & ".json"
End Function
