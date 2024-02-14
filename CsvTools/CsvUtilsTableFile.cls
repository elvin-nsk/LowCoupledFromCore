VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CsvUtilsTableFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'===============================================================================
'   Модуль          : CsvUtilsTableFile
'   Версия          : 2024.02.13
'   Автор           : elvin-nsk (me@elvin.nsk.ru)
'   Сайт            : https://github.com/elvin-nsk/LowCoupledFromCore
'   Назначение      : Адаптер CsvUtils к TableFile
'   Зависимости     : Нет
'===============================================================================

'@PredeclaredId
Option Explicit

Implements ITableFile

'===============================================================================
' # Declarations

Private Type This
    Arr As Variant
    Delimiter As String
    File As String
    StartingRow As Long
    StartingColumn As Long
    ReadOnly As Boolean
End Type
Private This As This

'===============================================================================
' # Constructors

Public Function New_( _
                    ByVal FileToBind As String, _
                    Optional ByVal Delimiter As String = ";", _
                    Optional ByVal StartingRow As Long = 1, _
                    Optional ByVal StartingColumn As Long = 1 _
                ) As ITableFile
    With New CsvUtilsTableFile
        Set New_ = _
            .Init(FileToBind, Delimiter, StartingRow, StartingColumn, False)
    End With
End Function

Public Function NewReadOnly( _
                    ByVal FileToBind As String, _
                    Optional ByVal Delimiter As String = ";", _
                    Optional ByVal StartingRow As Long = 1, _
                    Optional ByVal StartingColumn As Long = 1 _
                ) As ITableFile
    With New CsvUtilsTableFile
        Set NewReadOnly = _
            .Init(FileToBind, Delimiter, StartingRow, StartingColumn, True)
    End With
End Function

Friend Function Init( _
                    ByVal FileToBind As String, _
                    ByVal Delimiter As String, _
                    ByVal StartingRow As Long, _
                    ByVal StartingColumn As Long, _
                    ByVal ReadOnly As Boolean _
                ) As ITableFile
    With This
        .Delimiter = Delimiter
        .File = FileToBind
        .StartingColumn = StartingColumn
        .StartingRow = StartingRow
        .ReadOnly = ReadOnly
        Load
    End With
    Set Init = Me
End Function

Private Sub Class_Terminate()
    If VBA.IsArray(This.Arr) And Not This.ReadOnly Then
        Save
    End If
End Sub

'===============================================================================
' # Interfaces

Private Property Get ITableFile_Cell( _
                         ByVal Row As Long, _
                         ByVal Column As Long _
                     ) As String
    If AbsRow(Row) > Rows Or AbsColumn(Column) > Columns Then Exit Property
    With This
        ITableFile_Cell = .Arr(AbsRow(Row), AbsColumn(Column))
    End With
End Property
Private Property Let ITableFile_Cell( _
                         ByVal Row As Long, _
                         ByVal Column As Long, _
                         RHS As String _
                     )
    'TODO сделать добавление новых строк и колонок
    If AbsRow(Row) > Rows Or AbsColumn(Column) > Columns Then Exit Property
    With This
        .Arr(AbsRow(Row), AbsColumn(Column)) = RHS
    End With
End Property

Private Sub ITableFile_ForceClose()
    'Dummy
End Sub

Private Sub Load()
    Dim CsvString As String
    With This
        CsvString = ReadFile(.File)
        .Arr = CsvUtils.Create(.Delimiter).ParseCsvToArray(CsvString, False)
    End With
End Sub

Private Sub ITableFile_ForceLoad()
    Load
End Sub

Private Sub Save()
    With This
        If .ReadOnly Then Throw "Файл открыт только для чтения"
        Dim CsvString As String
        CsvString = CsvUtils.Create(.Delimiter).ConvertArrayToCsv(.Arr)
        WriteFile CsvString, .File
    End With
End Sub

Private Sub ITableFile_ForceSave()
    Save
End Sub

Private Property Get Columns() As Long
    Columns = Count(This.Arr, 2)
End Property

Private Property Get ITableFile_MaxColumn() As Long
    ITableFile_MaxColumn = Columns
End Property

Private Property Get Rows() As Long
    Rows = Count(This.Arr, 1)
End Property

Private Property Get ITableFile_MaxRow() As Long
    ITableFile_MaxRow = Rows
End Property

Private Property Get ITableFile_ReadOnly() As Boolean
    ITableFile_ReadOnly = This.ReadOnly
End Property

Private Property Get ITableFile_Self() As ITableFile
    Set ITableFile_Self = Me
End Property

'===============================================================================
' # Helpers

Private Property Get AbsColumn(ByVal RelativeColumn As Long) As Long
    AbsColumn = RelativeColumn + This.StartingColumn - 1
End Property

Private Property Get AbsRow(ByVal RelativeRow As Long) As Long
    AbsRow = RelativeRow + This.StartingRow - 1
End Property
