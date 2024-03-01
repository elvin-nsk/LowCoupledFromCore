VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressBarView 
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "ProgressBarView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ProgressBarView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Original: https://www.erlandsendata.no
Option Explicit

'===============================================================================
' # Declarations

Private Const GWL_STYLE = -16
Private Const WS_CAPTION = &HC00000
Private Const WS_SYSMENU = &H80000
Private Const SC_CLOSE = &HF060

#If VBA7 Then

    Private Declare PtrSafe Function FindWindowA _
        Lib "user32" (ByVal lpClassName As String, _
        ByVal lpWindowName As String) As Long
    Private Declare PtrSafe Function DeleteMenu _
        Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, _
        ByVal wFlags As Long) As Long
    Private Declare PtrSafe Function GetSystemMenu _
        Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
                
#Else

    Private Declare Function FindWindowA _
        Lib "user32" (ByVal lpClassName As String, _
        ByVal lpWindowName As String) As Long
    Private Declare Function DeleteMenu _
        Lib "user32" (ByVal hMenu As Long, _
        ByVal nPosition As Long, ByVal wFlags As Long) As Long
    Public Declare Function GetSystemMenu _
        Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
                
#End If

Public Event QueryCancel()
Public NumericMiddleText As String
Private LastDoneWidth As Long

'===============================================================================
' # Handlers

Private Sub UserForm_Initialize()
    CancelDisable
    Me.Caption = "Подождите..."
    Me.btnCancel.Caption = "Отмена"
    NumericMiddleText = "из"
    With Me.lblDone ' set the "progress bar" to it's initial length
        .Top = Me.lblRemain.Top + 1
        .Left = Me.lblRemain.Left + 1
        .Height = Me.lblRemain.Height - 2
        .Width = 0
    End With
End Sub

Private Sub btnCancel_Click()
    RaiseEvent QueryCancel
End Sub

'===============================================================================
' # Logic

Public Sub UpdateTo( _
               ByVal Current As Long, ByVal Max As Long, _
               Optional ByVal ShowAsPercentage = True _
           )
    If Current < 0 Then Current = VBA.Abs(Current)
    If Current > Max Then Current = Max
    Dim Dec As Double
    Dec = Current / Max
    Dim DoneWidth As Long
    With Me
        DoneWidth = VBA.CLng(Dec * (.lblRemain.Width - 2))
        If DoneWidth = LastDoneWidth Then
            Exit Sub
        Else
            LastDoneWidth = DoneWidth
        End If
        .lblDone.Width = DoneWidth
        If ShowAsPercentage Then
            .lblPct.Caption = VBA.Format(Dec, "0%")
        Else
            .lblPct.Caption = Current & " " & NumericMiddleText & " " & Max
        End If
    End With
    DoEvents 'The DoEvents statement is responsible for the form updating
End Sub

Public Property Let CancelButtonCaption(Value As String)
    Me.btnCancel.Caption = Value
End Property
Public Property Get CancelButtonCaption() As String
    CancelButtonCaption = Me.btnCancel.Caption
End Property

Public Property Let Cancelable(Value As Boolean)
    If Value Then CancelEnable Else CancelDisable
End Property

'===============================================================================
' # Helpers

Private Sub CancelDisable()
    Me.Height = 55
    With Me.btnCancel
        .Enabled = False
        .Cancel = False
    End With
    CloseButtonSettings Me, False
End Sub

Private Sub CancelEnable()
    Me.Height = 90
    With Me.btnCancel
        .Enabled = True
        .Cancel = True
        .SetFocus
    End With
    CloseButtonSettings Me, True
End Sub

'https://exceloffthegrid.com/hide-or-disable-a-vba-userform-X-close-button/
Private Sub CloseButtonSettings(Form As Object, Show As Boolean)
    Dim WindowHandle As Long
    Dim MenuHandle As Long
    WindowHandle = FindWindowA(vbNullString, Form.Caption)
    If Show Then
        MenuHandle = GetSystemMenu(WindowHandle, 1)
    Else
        MenuHandle = GetSystemMenu(WindowHandle, 0)
        DeleteMenu MenuHandle, SC_CLOSE, 0&
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        RaiseEvent QueryCancel
        Cancel = True
    End If
End Sub
