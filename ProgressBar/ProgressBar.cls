VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
Option Explicit

'===============================================================================
' # Declarations

Public Enum ProgressBarUnit
    ProgressBarNumeric = 0
    ProgressBarDecimal
    ProgressBarPercent
End Enum

Private Type This
    Cancelable As Boolean
    Canceled As Boolean
    Iteration As Long
    Max As Long
    Step As Long
    Unit As ProgressBarUnit
End Type
Private This As This

Private WithEvents View As ProgressBarView
Attribute View.VB_VarHelpID = -1

'===============================================================================
' # Constructors

Public Function New_( _
                    Optional ByVal Unit As ProgressBarUnit = ProgressBarNumeric, _
                    Optional ByVal Max As Long = 100, _
                    Optional ByVal Step As Long = 1 _
                ) As ProgressBar
    Set New_ = New ProgressBar
    New_.Inject Unit, Max, Step
End Function

Private Sub Class_Initialize()
    With This
        .Max = 100
        .Step = 1
        .Iteration = 0
        Set View = New ProgressBarView
    End With
End Sub

Friend Sub Inject( _
               ByVal Unit As ProgressBarUnit, _
               ByVal Max As Long, _
               ByVal Step As Long _
           )
    With This
        .Unit = Unit
        If IsValid(Max) Then .Max = Max Else Throw
        .Step = Step
        .Iteration = 0
        Set View = New ProgressBarView
    End With
End Sub

Private Sub Class_Terminate()
    DisposeView
End Sub

'===============================================================================
' # Handlers

Private Sub View_QueryCancel()
    With This
        If .Cancelable Then
            .Canceled = True
            DisposeView
        End If
    End With
End Sub

'===============================================================================
' # Public

Public Property Get Cancelable() As Boolean
    Cancelable = This.Cancelable
End Property
Public Property Let Cancelable(ByVal RHS As Boolean)
    This.Cancelable = RHS
    View.Cancelable = RHS
End Property

Public Property Get Canceled() As Boolean
    Canceled = This.Canceled
End Property

Public Property Let Caption(ByVal RHS As String)
    View.Caption = RHS
End Property
Private Property Get Caption() As String
    Caption = View.Caption
End Property

Public Property Get CancelButtonCaption() As String
    CancelButtonCaption = View.CancelButtonCaption
End Property
Public Property Let CancelButtonCaption(ByVal RHS As String)
    View.CancelButtonCaption = RHS
End Property

Public Property Get Max() As Long
    Max = This.Max
End Property

Public Property Get NumericMiddleText() As String
    NumericMiddleText = View.NumericMiddleText
End Property
Private Property Let NumericMiddleText(ByVal RHS As String)
    View.NumericMiddleText = RHS
End Property

Public Property Get Step() As Long
    Step = This.Step
End Property

Public Sub Update()
    ShowView
    Dim CurrentValue As Long
    With This
        .Iteration = .Iteration + 1
        CurrentValue = .Iteration * .Step
        View.UpdateTo CurrentValue, .Max, Not (.Unit = ProgressBarNumeric)
    End With
End Sub

Public Sub UpdateTo(ByVal RHS As Variant)
    ShowView
    With This
        If Not IsValid(RHS) Then Exit Sub
        Select Case .Unit
        Case ProgressBarDecimal
            View.UpdateTo VBA.CLng(RHS / 100), 100
        Case ProgressBarNumeric
            View.UpdateTo VBA.CLng(RHS), VBA.CLng(.Max), False
        Case ProgressBarPercent
            View.UpdateTo VBA.CLng(RHS), 100
        End Select
    End With
End Sub

Public Property Get Self() As ProgressBar
    Set Self = Me
End Property

'===============================================================================
' # Helpers

Private Sub DisposeView()
    On Error Resume Next
    View.Hide
    On Error GoTo 0
    Set View = Nothing
End Sub

Private Function IsValid(ByVal Value As Variant) As Boolean
    If Not VBA.IsNumeric(Value) Then Exit Function
    With This
        Select Case .Unit
        Case ProgressBarDecimal
            If Value > 0 And Value <= 1 Then IsValid = True
        Case ProgressBarNumeric
            If Value > 0 Then IsValid = True
        Case ProgressBarPercent
            If Value > 0 And Value <= 100 Then IsValid = True
        End Select
    End With
End Function

Private Sub ShowView()
    If Not View.Visible Then View.Show vbModeless
End Sub

Private Sub Throw()
    VBA.Err.Raise _
        vbObjectError Or 32, _
        "ProgressBar", _
        "Wrong value in current unit context"
End Sub
