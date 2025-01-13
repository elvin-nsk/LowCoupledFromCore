VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "подключение к файлу экселя через ADODB."
'===============================================================================
'   Модуль          : ExcelConnection
'   Версия          : 2025.01.13
'   Автор           : elvin-nsk (me@elvin.nsk.ru)
'   Описание        : подключение к файлу экселя через ADODB
'   Использование   :
'   Зависимости     : LibCore, TableFile
'===============================================================================

'@ModuleDescription("подключение к файлу экселя через ADODB.")

'@PredeclaredId
Option Explicit

Implements TableFile

'===============================================================================

Private Const adLockReadOnly = 1
Private Const adLockOptimistic = 3
Private Const adOpenStatic = 3
Private Const adUseServer = 2
Private Const adUseClient = 3
Private Const adSchemaTables = 20
Private Const adEditInProgress = 1
Private Const adResyncAllValues = 2

Private Type This
    File As String
    SheetName As String
    StartingRow As Long
    StartingColumn As Long
    RowOffset As Long
    ColumnOffset As Long
    ReadOnly As Boolean
    Connection As Object
    RecordSet As Object
End Type
Private This As This

'===============================================================================

Public Function Create( _
                    ByVal FileToBind As String, _
                    Optional ByVal SheetName As String, _
                    Optional ByVal StartingRow As Long = 1, _
                    Optional ByVal StartingColumn As Long = 1 _
                ) As TableFile
    With New ExcelConnection
        Set Create = _
            .Init(FileToBind, SheetName, StartingRow, StartingColumn, False)
    End With
End Function

Public Function CreateReadOnly( _
                    ByVal FileToBind As String, _
                    Optional ByVal SheetName As String, _
                    Optional ByVal StartingRow As Long = 1, _
                    Optional ByVal StartingColumn As Long = 1 _
                ) As TableFile
    With New ExcelConnection
        Set CreateReadOnly = _
            .Init(FileToBind, SheetName, StartingRow, StartingColumn, True)
    End With
End Function

Private Sub Class_Terminate()
    With This
        If Not .RecordSet Is Nothing Then
            Self.ForceSave
            .RecordSet.Close
        End If
        If Not .Connection Is Nothing Then .Connection.Close
    End With
End Sub

Friend Function Init( _
                    ByVal FileToBind As String, _
                    ByVal SheetName As String, _
                    ByVal StartingRow As Long, _
                    ByVal StartingColumn As Long, _
                    Optional ByVal ReadOnly As Boolean = False _
                ) As TableFile
    With This
        .File = FileToBind
        .SheetName = SheetName
        .StartingRow = StartingRow
        .StartingColumn = StartingColumn
        .RowOffset = .StartingRow - 2
        .ColumnOffset = .StartingColumn - 2
        .ReadOnly = ReadOnly
        Connect
    End With
    Set Init = Me
End Function

'===============================================================================

Private Property Let TableFile_Cell( _
                         ByVal Row As Long, _
                         ByVal Column As Long, _
                         RHS As String _
                     )
    With This.RecordSet
        If This.ReadOnly Then ThrowReadOnly
        If Not tryMoveCursorToRow(Row) Then Exit Property
        If VBA.IsNull(.Fields(Column + This.ColumnOffset).Value) And RHS = "" Then
            Exit Property
        End If
        .Fields(Column + This.ColumnOffset).Value = RHS
    End With
End Property
Private Property Get TableFile_Cell( _
                         ByVal Row As Long, _
                         ByVal Column As Long _
                     ) As String
    Dim TempVar As Variant
    With This.RecordSet
        If Not tryMoveCursorToRow(Row) Then Exit Property
        TempVar = .Fields(Column + This.ColumnOffset).Value
        If Not VBA.IsNull(TempVar) Then
            TableFile_Cell = TempVar
        End If
    End With
End Property

Private Sub TableFile_Dispose()
    With This
        If Not .RecordSet Is Nothing Then
            If .RecordSet.EditMode = adEditInProgress Then .RecordSet.CancelBatch
            .RecordSet.Close
            Set .RecordSet = Nothing
        End If
        If Not .Connection Is Nothing Then
            .Connection.Close
            Set .Connection = Nothing
        End If
    End With
End Sub

Private Sub TableFile_ForceLoad()
    ThrowNotImplemented "ForceLoad"
End Sub

Private Sub TableFile_ForceSave()
    With This
        If .RecordSet.EditMode = adEditInProgress Then .RecordSet.Update
    End With
End Sub

Private Property Get TableFile_MaxColumn() As Long
    TableFile_MaxColumn = This.RecordSet.Fields.Count
End Property

Private Property Get TableFile_MaxRow() As Long
    TableFile_MaxRow = This.RecordSet.RecordCount
End Property

Private Property Get TableFile_ReadOnly() As Boolean
    TableFile_ReadOnly = This.ReadOnly
End Property

Private Property Get TableFile_Self() As TableFile
    Set TableFile_Self = Me
End Property

'===============================================================================

Private Sub Connect()
    Set This.Connection = VBA.CreateObject("ADODB.Connection")
    With This.Connection
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Properties("Extended Properties").Value = "Excel 12.0;HDR=No"
        .Open This.File
        ValidateSheetName
    End With
    Set This.RecordSet = VBA.CreateObject("ADODB.RecordSet")
    With This.RecordSet
        Set .ActiveConnection = This.Connection
        .CursorType = adOpenStatic
        .CursorLocation = adUseClient
        If This.ReadOnly Then
            .LockType = adLockReadOnly
        Else
            .LockType = adLockOptimistic
        End If
        .Source = "SELECT * FROM [" & This.SheetName & "]"
        .Open
    End With
End Sub

Private Function ValidateSheetName()
    If This.SheetName = "" Then
        With This.Connection.OpenSchema(adSchemaTables)
            This.SheetName = .Fields("table_name").Value
            .Close
        End With
    Else
        This.SheetName = This.SheetName & "$"
    End If
End Function

Private Function tryMoveCursorToRow(ByVal Row As Long) As Boolean
    With This.RecordSet
        .MoveFirst
        .Move Row + This.RowOffset
        If Not .EOF Then tryMoveCursorToRow = True
    End With
End Function

Private Sub ThrowReadOnly()
    Throw "TableIsReadOnly"
End Sub

Private Sub ThrowNotImplemented(Optional ByVal FunctionName As String = "Function")
    Throw FunctionName & " not implemented"
End Sub

Private Property Get Self() As TableFile
    Set Self = Me
End Property
