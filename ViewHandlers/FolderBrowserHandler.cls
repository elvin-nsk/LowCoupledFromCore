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
'   Версия          : 2023.11.19
'   Автор           : elvin-nsk (me@elvin.nsk.ru)
'   Сайт            : https://github.com/elvin-nsk/LowCoupledFromCore
'   Назначение      : обслуживание связки TextBox+Button для выбора папки
'   Зависимости     : FileSpec
'===============================================================================

'@Folder UI
'@PredeclaredId
Option Explicit

'===============================================================================
' # Declarations

Private Const CustomError = vbObjectError Or 32

Private WithEvents TextBoxPath As MSForms.TextBox
Attribute TextBoxPath.VB_VarHelpID = -1
Private WithEvents ButtonBrowse As MSForms.CommandButton
Attribute ButtonBrowse.VB_VarHelpID = -1

'===============================================================================
' # Constructors

Public Function New_( _
                    ByVal TextBox As MSForms.TextBox, _
                    ByVal Button As MSForms.CommandButton, _
                    Optional ByVal InitialPath As String _
                ) As FolderBrowserHandler
    With New FolderBrowserHandler
        Set New_ = .Init(TextBox, Button, InitialPath)
    End With
End Function

Friend Function Init( _
                    ByVal TextBox As MSForms.TextBox, _
                    ByVal Button As MSForms.CommandButton, _
                    ByVal InitialPath As String _
                ) As FolderBrowserHandler
    Set TextBoxPath = TextBox
    If Not InitialPath = vbNullString Then TextBoxPath.Value = InitialPath
    Set ButtonBrowse = Button
    Set Init = Me
End Function

'===============================================================================
' # Handlers

Private Sub ButtonBrowse_Click()
    Dim LastPath As String
    LastPath = TextBoxPath.Value
    Dim Folder As New FileSpec
    Folder.Inject TextBoxPath.Value
    Folder.Path = CorelScriptTools.GetFolder(Folder.Path)
    If Folder.Path = "\" Then
        TextBoxPath.Value = LastPath
    Else
        TextBoxPath.Value = Folder.Path
    End If
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
