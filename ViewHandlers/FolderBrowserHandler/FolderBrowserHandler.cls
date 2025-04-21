VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FolderBrowserHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'===============================================================================
'   Модуль          : FolderBrowserHandler
'   Версия          : 2024.02.21
'   Автор           : elvin-nsk (me@elvin.nsk.ru)
'   Сайт            : https://github.com/elvin-nsk/LowCoupledFromCore
'   Назначение      : обслуживание связки TextBox+Button для выбора папки
'   Зависимости     : FileSpec
'===============================================================================

'@Interface
'@PredeclaredId
Option Explicit

'===============================================================================
' # Declarations

Private WithEvents TextBoxPath As MSForms.TextBox
Attribute TextBoxPath.VB_VarHelpID = -1
Private WithEvents Button As MSForms.CommandButton
Attribute Button.VB_VarHelpID = -1

'===============================================================================
' # Constructors

Friend Function New_( _
                    ByVal TextBox As MSForms.TextBox, _
                    ByVal BrowseButton As MSForms.CommandButton, _
                    Optional ByVal InitialPath As String _
                ) As FolderBrowserHandler
    Set New_ = New FolderBrowserHandler
    New_.Inject TextBox, BrowseButton, InitialPath
End Function

Friend Sub Inject( _
               ByVal TextBox As MSForms.TextBox, _
               ByVal BrowseButton As MSForms.CommandButton, _
               ByVal InitialPath As String _
           )
    Set TextBoxPath = TextBox
    If Not InitialPath = vbNullString Then TextBoxPath.Value = InitialPath
    Set Button = BrowseButton
End Sub

'===============================================================================
' # Handlers

Private Sub Button_Click()
    Dim LastPath As String: LastPath = TextBoxPath.Value
    Dim Folder As FileSpec: Set Folder = FileSpec.New_(TextBoxPath.Value)
    Folder.Path = CorelScriptTools.GetFolder(Folder.Path)
    If Folder.Path = "\" Then
        TextBoxPath.Value = LastPath
    Else
        TextBoxPath.Value = Folder.Path
    End If
    #If DEV = 1 Then
    DebugOut "FolderBrowserHandler", "Browser return", TextBoxPath.Value
    #End If
End Sub

'===============================================================================
' # Public

'@DefaultMember
Public Property Get Value() As Variant
Attribute Value.VB_UserMemId = 0
    Value = TextBoxPath.Value
End Property
Public Property Let Value(ByVal RHS As Variant)
    TextBoxPath.Value = RHS
End Property

Public Property Get Self() As FolderBrowserHandler
    Set Self = Me
End Property
