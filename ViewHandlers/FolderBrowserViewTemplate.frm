VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FolderBrowserViewTemplate 
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6975
   OleObjectBlob   =   "FolderBrowserViewTemplate.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FolderBrowserViewTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'===============================================================================
' # State

Public IsOk As Boolean
Public IsCancel As Boolean

Private SourceFolder As FolderBrowserHandler
Private OutputFolder As FolderBrowserHandler

'===============================================================================
' # Constructor

Private Sub UserForm_Initialize()    
    Set SourceFolder = _
        FolderBrowserHandler.New_(SourceFolderBox, SourceFolderBrowse)
    Set OutputFolder = _
        FolderBrowserHandler.New_(OutputFolderBox, OutputFolderBrowse)
End Sub

'===============================================================================
' # Handlers

Private Sub UserForm_Activate()
    '
End Sub

Private Sub btnOk_Click()
    FormŒ 
End Sub

Private Sub btnCancel_Click()
    FormCancel
End Sub

'===============================================================================
' # Logic

Private Sub FormŒ ()
    Hide
    IsOk = True
End Sub

Private Sub FormCancel()
    Hide
    IsCancel = True
End Sub

'===============================================================================
' # Helpers


'===============================================================================
' # Boilerplate

Private Sub UserForm_QueryClose(—ancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        —ancel = True
        FormCancel
    End If
End Sub
