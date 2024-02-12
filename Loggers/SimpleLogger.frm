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
' �����            : SimpleLogger
' ������           : 2023.11.07
' �����            : elvin-nsk (me@elvin.nsk.ru)
' ����������:      : ������� ���� ������� � ������
'===============================================================================

Option Explicit

'===============================================================================

Private Type This
  MessagesCount As Long
End Type
Private This As This

'===============================================================================

Public Sub Add(ByVal Text As String)
  This.MessagesCount = This.MessagesCount + 1
  Report.Value = Report.Value & Text & vbCrLf
End Sub

Public Property Get Count()
  Count = This.MessagesCount
End Property

'������� ���, ���� �� �� ����
Public Sub Check(Optional ByVal ListCaption As String = "���")
  If This.MessagesCount = 0 Then Exit Sub
  Caption = ListCaption
  Show vbModeless
  Report.SetFocus
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

Private Sub UserForm_QueryClose(�ancel As Integer, CloseMode As Integer)
  If CloseMode = VbQueryClose.vbFormControlMenu Then
    �ancel = True
    FormCancel
  End If
End Sub
