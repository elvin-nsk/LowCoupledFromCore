Attribute VB_Name = "a__Main"
'===============================================================================
'   Макрос          : MacroName
'   Версия          : 2025.01.01
'   Сайт            : https://github.com/AuthorName
'   Автор           : AuthorName
'===============================================================================

Option Explicit

'===============================================================================
' # Manifest

Public Const APP_NAME As String = "MacroName"
Public Const APP_DISPLAYNAME As String = APP_NAME
Public Const APP_FILEBASENAME As String = APP_NAME
Public Const APP_VERSION As String = "2025.01.01"

'===============================================================================
' # Globals

Private Const SOME_CONST As String = ""

'===============================================================================
' # Entry points

Sub Start()

    #If DebugMode = 0 Then
    On Error GoTo Catch
    #End If
    
    Dim Shapes As ShapeRange
    If Not InputData.ExpectShapes.Ok(Shapes) Then GoTo Finally
    
    Dim Source As ShapeRange: Set Source = ActiveSelectionRange
    
    BoostStart APP_DISPLAYNAME
    
    '??? PROFIT!
    
    Source.CreateSelection
    
Finally:
    BoostFinish
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Source & ": " & VBA.Err.Description, vbCritical, "Error"
    Resume Finally

End Sub

'===============================================================================
' # Helpers

Private Sub Helper()
'
End Sub

'===============================================================================
' # Tests

Private Sub TestSomething()
'
End Sub
