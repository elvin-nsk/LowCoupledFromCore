VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Logger 
   ClientHeight    =   5190
   ClientLeft      =   90
   ClientTop       =   420
   ClientWidth     =   5820
   OleObjectBlob   =   "Logger.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Logger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================================
' Форма            : Logger
' Версия           : 2022.01.15
' Автор            : elvin-nsk (me@elvin.nsk.ru)
' Назначение:      : ведение лога событий и ошибок
'===============================================================================

Option Explicit

'===============================================================================

Private Type typeMessage
  Text As String
  Object As Object
End Type

Private Type typeThis
  Messages() As typeMessage
  MessagesCount As Long
End Type
Private This As typeThis

'===============================================================================

'добавить сообщение в лог, с опциональной привязкой к объекту
Public Sub Add(ByVal Text As String, Optional ConnectedObject As Object)
  This.MessagesCount = This.MessagesCount + 1
  ReDim Preserve This.Messages(1 To This.MessagesCount)
  This.Messages(This.MessagesCount).Text = Text
  lstMain.AddItem Text
  If Not ConnectedObject Is Nothing Then
    If TypeOf ConnectedObject Is Document Or _
       TypeOf ConnectedObject Is Page Or _
       TypeOf ConnectedObject Is Layer Or _
       TypeOf ConnectedObject Is ShapeRange Or _
       TypeOf ConnectedObject Is Shape Then _
      Set This.Messages(This.MessagesCount).Object = ConnectedObject
  End If
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

Private Sub lstMain_Change()
  If lstMain.ListCount = 0 Then Exit Sub
  lbMain.Caption = lstMain.List(lstMain.ListIndex)
  If This.Messages(ActiveMessageNum).Object Is Nothing Then
    btnZoom.Enabled = False
  Else
    btnZoom.Enabled = True
  End If
End Sub

Private Sub lstMain_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
  If KeyAscii = 13 Then Focus 'если Enter
End Sub

Private Sub lstMain_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
  Focus
End Sub

Private Sub btnZoom_Click()
  Focus
End Sub

Private Sub btnNext_Click()
  With lstMain
    If .ListIndex + 1 < .ListCount Then .ListIndex = .ListIndex + 1
  End With
End Sub

Private Sub btnPrev_Click()
  With lstMain
    If .ListIndex > 0 Then .ListIndex = .ListIndex - 1
  End With
End Sub

'===============================================================================

'фокусируется на привязанном к ошибке объекте
Private Sub Focus()
  With This.Messages(ActiveMessageNum)
    If .Object Is Nothing Then Exit Sub
    If TypeOf .Object Is Document Then
      ActivateDocument .Object
    End If
    If TypeOf .Object Is Page Then
      ActivatePage .Object
    End If
    If TypeOf .Object Is Layer Then
      ActivateLayer .Object
    End If
    If TypeOf .Object Is ShapeRange Then
      ActivateShapeRange .Object
    End If
    If TypeOf .Object Is Shape Then
      ActivateShape .Object
    End If
  End With
  'ActiveWindow.Refresh
  'Application.Refresh
  'Application.Windows.Refresh
End Sub

Private Function ActiveMessageNum() As Long
  ActiveMessageNum = lstMain.ListIndex + 1
End Function

Private Sub ActivateDocument(Document As Document)
  Document.Activate
End Sub

Private Sub ActivatePage(Page As Page)
  Page.Parent.Parent.Activate
  On Error Resume Next
  Page.Activate
  ActiveWindow.ActiveView.ToFitPage
End Sub

Private Sub ActivateLayer(Layer As Layer)
  ActivatePage Layer.Parent.Parent
  If Not Layer.Master Then Layer.Activate
End Sub

Private Sub ActivateShapeRange(ShapeRange As ShapeRange)
  ActivateLayer ShapeRange.FirstShape.Layer
  ShapeRange.CreateSelection
  ActiveWindow.ActiveView.ToFitSelection
End Sub

Private Sub ActivateShape(Shape As Shape)
  ActivateLayer Shape.Layer
  Shape.CreateSelection
  ActiveWindow.ActiveView.ToFitSelection
End Sub

Private Sub FormCancel()
  Me.Hide
End Sub

'===============================================================================

Private Sub GuardInt(ByVal KeyAscii As MSForms.ReturnInteger)
  Select Case KeyAscii
    Case VBA.Asc("0") To VBA.Asc("9")
    Case Else
      KeyAscii = 0
  End Select
End Sub

Private Sub GuardNum(ByVal KeyAscii As MSForms.ReturnInteger)
  Select Case KeyAscii
    Case VBA.Asc("0") To VBA.Asc("9")
    Case VBA.Asc(",")
    Case Else
      KeyAscii = 0
  End Select
End Sub

Private Sub GuardRangeDbl(TextBox As MSForms.TextBox, _
                          ByVal Min As Double, _
                          Optional ByVal Max As Double = 1.79769313486231E+308)
  With TextBox
    If .Value = "" Then .Value = VBA.CStr(Min)
    If VBA.CDbl(.Value) > Max Then .Value = VBA.CStr(Max)
    If VBA.CDbl(.Value) < Min Then .Value = VBA.CStr(Min)
  End With
End Sub

Private Sub GuardRangeLng(TextBox As MSForms.TextBox, _
                          ByVal Min As Long, _
                          Optional ByVal Max As Long = 2147483647)
  With TextBox
    If .Value = "" Then .Value = VBA.CStr(Min)
    If VBA.CLng(.Value) > Max Then .Value = VBA.CStr(Max)
    If VBA.CLng(.Value) < Min Then .Value = VBA.CStr(Min)
  End With
End Sub

Private Sub UserForm_QueryClose(Сancel As Integer, CloseMode As Integer)
  If CloseMode = VbQueryClose.vbFormControlMenu Then
    Сancel = True
    FormCancel
  End If
End Sub
