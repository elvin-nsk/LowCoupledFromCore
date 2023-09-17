VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SimpleLogger 
   ClientHeight    =   5190
   ClientLeft      =   90
   ClientTop       =   420
   ClientWidth     =   5820
   OleObjectBlob   =   "SimpleLogger.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SimpleLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================================
' Форма            : SimpleLogger
' Версия           : 2023.08.28
' Автор            : elvin-nsk (me@elvin.nsk.ru)
' Назначение:      : ведение лога событий и ошибок
'===============================================================================

Option Explicit

'===============================================================================

Private Type typeThis
  Messages() As String
  MessagesCount As Long
End Type
Private This As typeThis

'===============================================================================

'добавить сообщение в лог, с опциональной привязкой к объекту
Public Sub Add(ByVal Text As String)
  This.MessagesCount = This.MessagesCount + 1
  ReDim Preserve This.Messages(1 To This.MessagesCount)
  This.Messages(This.MessagesCount) = Text
  lstMain.AddItem Text
End Sub

Public Property Get Count()
  Count = This.MessagesCount
End Property

'вывести лог, если он не пуст
Public Sub Check(Optional ByVal ListCaption As String = "Лог")
  If This.MessagesCount = 0 Then Exit Sub
  Caption = ListCaption
  Show vbModeless
  lstMain.SetFocus
  lstMain.ListIndex = 0
End Sub

'===============================================================================

Private Sub UserForm_Initialize()
  '
End Sub

Private Sub CloseButton_Click()
  FormCancel
End Sub

'===============================================================================

Private Sub FormCancel()
  Me.Hide
End Sub

'===============================================================================

Private Sub UserForm_QueryClose(Сancel As Integer, CloseMode As Integer)
  If CloseMode = VbQueryClose.vbFormControlMenu Then
    Сancel = True
    FormCancel
  End If
End Sub
