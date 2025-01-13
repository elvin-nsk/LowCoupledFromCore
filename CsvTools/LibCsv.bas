Attribute VB_Name = "LibCsv"
'===============================================================================
'   Модуль          : LibCsv
'   Версия          : 2025.01.13
'   Автор           : elvin-nsk (me@elvin.nsk.ru)
'   Сайт            : https://github.com/elvin-nsk/LowCoupledFromCore
'   Назначение      : вспомогательные функции к классу CsvUtils
'                   : для парсинга csv-файлов
'   Зависимости     : CsvUtils, LibCore
'===============================================================================

Option Explicit

Public Function FileToKeyedColumns( _
                    ByVal CsvFile As String, _
                    Optional ByVal CharSet As String = "utf-8", _
                    Optional ByVal CsvSeparator As String = ";" _
                ) As Scripting.IDictionary
    Set FileToKeyedColumns = _
        TableToKeyedColumns( _
            GetTableFromFile(CsvFile, CharSet, CsvSeparator) _
        )
End Function

Public Function GetTableFromFile( _
                    ByVal CsvFile As String, _
                    Optional ByVal CharSet As String = "utf-8", _
                    Optional ByVal CsvSeparator As String = ";" _
                ) As String()
    Dim Str As String
    Str = ReadFileAD(CsvFile, CharSet)
    GetTableFromFile = _
        CsvUtils.New_(CsvSeparator) _
            .ParseCsvToArray(Str, False)
End Function

Public Function TableToKeyedColumns( _
                    ByRef Table() As String _
                ) As Scripting.IDictionary
    Dim Dic As New Scripting.Dictionary
    Dim Row As Long
    Dim Column As Long
    Dim Key As String
    For Column = 1 To UBound(Table, 2)
        Key = Table(1, Column)
        Dic.Add Key, New Collection
        For Row = 2 To UBound(Table, 1)
            Dic(Key).Add Table(Row, Column)
        Next Row
    Next Column
    Set TableToKeyedColumns = Dic
End Function
