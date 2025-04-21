VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ViewToDictionaryBinder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'===============================================================================
'   Модуль          : FormToDictionaryBinder
'   Версия          : 2024.06.18
'   Автор           : elvin-nsk (me@elvin.nsk.ru)
'   Сайт            : https://github.com/elvin-nsk/LowCoupledFromCore
'   Назначение      : Привязывает контролы формы к словарю
'                     для автоматической загрузки/сохранения состояния
'   Зависимости     : LibCore, JsonConverter
'===============================================================================

'@Interface
'@PredeclaredId
Option Explicit

'===============================================================================
' # Declarations

Private Type This
    ControlNames As Variant
    Dictionary As Scripting.IDictionary
    ReadOnly As Boolean
    View As MSForms.UserForm
End Type
Private This As This
Attribute This.VB_VarHelpID = -1

'===============================================================================
' # Constructors

Friend Function New_( _
                    ByVal Dictionary As Scripting.IDictionary, _
                    ByVal View As MSForms.UserForm, _
                    ByVal ControlNames As Variant _
                ) As ViewToDictionaryBinder
    Set New_ = New ViewToDictionaryBinder
    New_.Inject Dictionary, View, ControlNames
End Function

Friend Sub Inject( _
               ByVal Dictionary As Scripting.IDictionary, _
               ByVal View As MSForms.UserForm, _
               ByVal ControlNames As Variant _
           )
    With This
        Set .Dictionary = Dictionary
        Assign .ControlNames, ControlNames
        Set .View = View
        .ReadOnly = ReadOnly
        RefreshForm
    End With
End Sub

'===============================================================================
' # Public

Public Property Get Dictionary() As Scripting.Dictionary
    Set Dictionary = This.Dictionary
End Property

'@DefaultMember
Public Property Get Item(ByVal Key As String) As Variant
    Item = This.Dictionary(Key)
End Property
Public Property Let Item(ByVal Key As String, RHS As Variant)
Attribute Item.VB_UserMemId = 0
    This.Dictionary(Key) = RHS
End Property

Public Function RefreshDictionary() As Scripting.Dictionary
    With This
        Dim ItemName As Variant
        For Each ItemName In This.ControlNames
            .Dictionary(ItemName) = .View.Controls(ItemName)
        Next ItemName
        Set RefreshDictionary = .Dictionary
    End With
End Function

Public Sub RefreshForm()
    With This
        Dim ItemName As Variant
        For Each ItemName In This.ControlNames
            If .Dictionary.Exists(ItemName) Then
                .View.Controls(ItemName) = .Dictionary(ItemName)
            End If
        Next ItemName
    End With
End Sub

Public Property Get Self() As ViewToDictionaryBinder
    Set Self = Me
End Property

'===============================================================================
' # Helpers
