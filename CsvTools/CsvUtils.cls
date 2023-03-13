VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CsvUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'===============================================================================
'
' This class based on
' VBA-CSV
'
' Copyright (C) 2017- sdkn104 ( https://github.com/sdkn104/VBA-CSV/ )
' License MIT (http://www.opensource.org/licenses/mit-license.php)
' Document: https://github.com/sdkn104/VBA-CSV/README.md
'
'===============================================================================

'@PredeclaredId
Option Explicit

'===============================================================================

' Field Quoting
'     Used for the argument 'Quoting' of ConvertArrayToCsv()
'     This argument controls what kind of fields to be quoted
Public Enum CsvUtilsQuote
    ' quote the fields that requires quotation
    ' (i.e., that includes comma, return code, quotation mark)
    CsvUtilsMinimal = 0
    ' quote all the fields
    CsvUtilsAll = 1
    ' quote non-numeric (Not IsNumeric()) fields
    CsvUtilsNonNumeric = 2
End Enum

' Variables used in FindNextSeparator()
Private NextSep1 As Long
Private NextSep2 As Long
Private NextSep3 As Long

Private Type typeThis
    Delimiter As String
    CsvUtilsAnyErrorIsFatal As Boolean  'default False
End Type
Private This As typeThis

'===============================================================================

Public Function Create(Optional ByVal Delimiter As String = ";") As CsvUtils
    With New CsvUtils
        Set Create = .Init(Delimiter)
    End With
End Function

Friend Function Init(ByVal Delimiter As String) As CsvUtils
    With This
        .Delimiter = Delimiter
    End With
    Set Init = Me
End Function

'===============================================================================

'----- ERROR HANDLER -----------------------------------------------------------

'
' Error function
'
Private Sub ErrorRaise(Code As Long, Src As String, Msg As String)
    ' raise only if this is the first error
    If VBA.Err.Number = 0 Then VBA.Err.Raise Code, Src, Msg
End Sub

'
' Setting error handling mode
'
'    False (default) --- When run-time error occurs, the parser function returns special value (Nothing,    Null, etc.),
'                        and the error information is set to properties of Err object.
'    True            --- Any run-time error that occurs is fatal (an error message is displayed and execution stops).
'
Public Sub SetCsvUtilsAnyErrorIsFatal(ByRef Value As Boolean)
    This.CsvUtilsAnyErrorIsFatal = Value
End Sub

'------ Public Function/Sub ----------------------------------------------------

'
' Parse CSV Text and retern Collection
'
'     Return a Collection of records; record is a Collection of fields
'     When error, return Nothing
'
Public Function ParseCsvToCollection( _
                    ByRef CsvText As String, _
                    Optional ByRef AllowVariableNumOfFields As Boolean = False, _
                    Optional ByRef HeaderOnly As Boolean = False _
                ) As Collection
    ' "On Error Resume Next" only if CsvUtilsAnyErrorIsFatal is True
    VBA.Err.Clear
    If This.CsvUtilsAnyErrorIsFatal Then GoTo Head
    On Error Resume Next
Head:
    Dim CsvPos As Long
    Dim FieldText As String
    Dim NextSep As Long, NextSepType As Long, QuoteCount As Long, FieldStart As Long, FieldLen As Long
    Dim Fields As Collection
    Dim CsvCollection As Collection
    Set CsvCollection = New Collection  'empty collection
        
    Set ParseCsvToCollection = CsvCollection
        
    'for empty Text
    If CsvText = "" Then Exit Function  'return empty collection
        
    ' Add trailing record separator if not
    If Right(CsvText, 1) <> "" & vbCr And Right(CsvText, 1) <> "" & vbLf Then
        CsvText = CsvText & vbCrLf
    End If
        
    'extract records and fields
    CsvPos = 1
    Set Fields = New Collection
    Call FindNextSeparatorInit(CsvText)
    Do While FindNextSeparator(CsvText, CsvPos, FieldStart, FieldLen, NextSepType, QuoteCount)
        FieldText = Mid(CsvText, FieldStart, FieldLen)
        If Err.Number <> 0 Then Exit Do
                
        If QuoteCount > 0 Then                                             ' the field includes " (double-quote)
            FieldText = TrimQuotes(FieldText)                    'get internal of ""
            If QuoteCount > 2 Then                                         'the field includes double-quote in internal of ""
                FieldText = VBA.Replace(FieldText, """""", """") 'un-escape double quote
                If FieldText Like "=""*""" Then FieldText = Mid(FieldText, 3, Len(FieldText) - 3) 'remove MS quote (="...")
            End If
        End If
        'add to collection
        Fields.Add FieldText
                        
        If NextSepType <> 1 Then                                         ' end of the record
            CsvCollection.Add Fields
            If HeaderOnly Then Exit Do
            If Not AllowVariableNumOfFields And CsvCollection.Item(1).Count <> Fields.Count Then
                ErrorRaise 10001, "ParseCsvToCollection", "Syntax Error in CSV: numbers of fields are different among records"
                GoTo ErrorExit
            End If
            Set Fields = New Collection
        End If
    Loop
    If Err.Number <> 0 Then GoTo ErrorExit
        
    Set ParseCsvToCollection = CsvCollection
    Exit Function

ErrorExit:
    Set ParseCsvToCollection = Nothing
End Function

'
' Parse CSV Text and return 2-dim array
'
'    Return 2-dim array --- String(1 TO recordCount, 1 TO fieldCount)
'    When CSV Text is "", return empty array --- String(0 TO -1)
'    When error, return Null
'
Public Function ParseCsvToArray( _
                    ByRef CsvText As String, _
                    Optional ByRef AllowVariableNumOfFields As Boolean = False _
                ) As Variant
    ' "On Error Resume Next" only if CsvUtilsAnyErrorIsFatal is True
    Err.Clear
    If This.CsvUtilsAnyErrorIsFatal Then GoTo Head
    On Error Resume Next
Head:
    Dim Csv As Collection
    Dim RowCount As Long, ColCount As Long
    Dim CsvArray() As String
    Dim Ri As Long, Fi As Long
    Dim SepIndex As Long
    Dim FieldStart As Long, FieldLen As Long, NextSepType As Long, QuoteCount As Long
    Dim FieldText As String
        
    ParseCsvToArray = Null                                                 'for error
    
    Dim SepArray1() As Long
    Dim SepArray2() As Long
    Dim SepArray3() As Long
    Dim SepArray4() As Long
    ReDim SepArray1(Len(CsvText) / 40 + 64)
    ReDim SepArray2(Len(CsvText) / 40 + 64)
    ReDim SepArray3(Len(CsvText) / 40 + 64)
    ReDim SepArray4(Len(CsvText) / 40 + 64)
    
    ' Parse CSV and get row/col count, sepArray1234
    Call ParseCsv(RowCount, ColCount, SepArray1, SepArray2, SepArray3, SepArray4, CsvText, AllowVariableNumOfFields)
    If Err.Number <> 0 Then                                                'error occur
        Exit Function
    End If
        
    ' empty
    If RowCount = 0 Then
        ParseCsvToArray = VBA.Split("", "/")                         'return empty(zero length) String array of bound 0 TO -1
        '(https://msdn.microsoft.com/ja-jp/library/office/gg278528.aspx)
        Exit Function
    End If
                
    ' allocate result array
    ReDim CsvArray(1 To RowCount, 1 To ColCount) As String
        
    ' fill result array
    SepIndex = 0
    Ri = 1
    Fi = 1
    Do
        FieldStart = SepArray1(SepIndex)
        If FieldStart = 0 Then Exit Do                             ' EOF
                
        FieldLen = SepArray2(SepIndex)
        NextSepType = SepArray3(SepIndex)
        QuoteCount = SepArray4(SepIndex)
        FieldText = VBA.Mid(CsvText, FieldStart, FieldLen)
        If QuoteCount > 0 Then                                             ' the field includes " (double-quote)
            FieldText = TrimQuotes(FieldText)                    'get internal of ""
            If QuoteCount > 2 Then                                         'the field includes double-quote in internal of ""
                FieldText = VBA.Replace(FieldText, """""", """") 'un-escape double quote
                If FieldText Like "=""*""" Then FieldText = Mid(FieldText, 3, Len(FieldText) - 3) 'remove MS quote (="...")
            End If
        End If
        CsvArray(Ri, Fi) = FieldText
        Fi = Fi + 1
        If NextSepType <> 1 Then                                         ' end of record
            Ri = Ri + 1
            Fi = 1
        End If
        SepIndex = SepIndex + 1
    Loop
    ParseCsvToArray = CsvArray
End Function

'
' Convert 2-dim array to CSV Text string
'
'    inArray : 2-dim array of arbitary size/range and type.
'    fmtDate : format used for conversion from type Date to type String
'    When error, return ""
'
Public Function ConvertArrayToCsv( _
                    inArray As Variant, _
                    Optional fmtDate As String = "yyyy/m/d", _
                    Optional ByVal Quoting As CsvUtilsQuote = CsvUtilsMinimal, _
                    Optional ByVal RecordSeparator As String = vbCrLf _
                ) As String
    ' "On Error Resume Next" only if CsvUtilsAnyErrorIsFatal is True
    Err.Clear
    If This.CsvUtilsAnyErrorIsFatal Then GoTo Head
    On Error Resume Next
Head:
    Dim Csv As String
    Dim r As Long, c As Long, ub2 As Long
    Dim v As Variant
    Dim Cell As String
    Dim ArrRecord As Variant, ArrField As Variant
        
    'error check
    If Not VBA.IsArray(inArray) Then
        ErrorRaise 10004, "ConvertArrayToCsv", "Input argument inArray is not array"
        GoTo ErrorExit
    End If
    ub2 = UBound(inArray, 2)
    If Err.Number <> 0 Then                                                'expecting Err.Number = 9, Err.Description = "Subscript out of range", for inArray is 1-dim
        GoTo ErrorExit
    End If

    Dim rc As Long, cc As Long
    ReDim ArrRecord(LBound(inArray, 1) To UBound(inArray, 1)) As String 'temporary array
    ReDim ArrField(LBound(inArray, 2) To UBound(inArray, 2)) As String 'temporary array
        
    For r = LBound(inArray, 1) To UBound(inArray, 1)
        For c = LBound(inArray, 2) To UBound(inArray, 2)
            v = inArray(r, c)
            'formatting
            Cell = VBA.IIf(IsNull(v), "", v)
            If VBA.TypeName(v) = "Date" Then Cell = VBA.Format(v, fmtDate)
            'quote and escape
            If Quoting = CsvUtilsQuote.CsvUtilsAll Or _
                 (Quoting = CsvUtilsQuote.CsvUtilsNonNumeric And Not VBA.IsNumeric(v)) Or _
                 VBA.InStr(Cell, This.Delimiter) > 0 Or VBA.InStr(Cell, """") > 0 Or VBA.InStr(Cell, vbCr) > 0 Or VBA.InStr(Cell, vbLf) > 0 Then
                Cell = VBA.Replace(Cell, """", """""")
                Cell = """" & Cell & """"
            End If
            'add to array
            ArrField(c) = Cell
        Next
        ArrRecord(r) = VBA.Join(ArrField, This.Delimiter) & RecordSeparator
    Next
    If Err.Number <> 0 Then GoTo ErrorExit                 'unexpected error
        
    ConvertArrayToCsv = VBA.Join(ArrRecord, "")
    Exit Function
ErrorExit:
    ConvertArrayToCsv = ""
End Function

'
' ParseCsvToDictionary
'        return Dictionary whose key is value of KeyColumn and whose value is a Collection of fields in the record
'
Public Function ParseCsvToDictionary( _
                    ByRef CsvText As String, _
                    Optional ByRef KeyColumn As Long = 1, _
                    Optional ByRef AllowVariableNumOfFields As Boolean = False _
                ) As Object
    Dim Collection As Collection
    Dim Dic As Object
    Dim r As Long
    Set ParseCsvToDictionary = Nothing                         'for error
    Set Collection = ParseCsvToCollection(CsvText, AllowVariableNumOfFields)
    If Collection Is Nothing Then Exit Function        ' error
    Set Dic = CreateObject("Scripting.Dictionary")
    For r = 1 To Collection.Count                                    'include header row
        Set Dic(Collection(r)(KeyColumn)) = Collection(r)
    Next
    Set ParseCsvToDictionary = Dic
End Function

'
' GetFieldDictionary
'        return Dictionary whose key is field name and whose value is column number (1,2,3,...)    of the field
'
Public Function GetFieldDictionary(ByRef CsvText As String) As Object
    Dim Collection As Collection
    Dim c As Long
    Dim v
    Set Collection = ParseCsvToCollection(CsvText, True, True)            'parse header only
    Set GetFieldDictionary = Nothing                                                                ' for error
    If Collection Is Nothing Then Exit Function                                         ' Error
    Set GetFieldDictionary = CreateObject("Scripting.Dictionary")
    If Collection.Count = 0 Then Exit Function                                            ' no field (empty)
    For c = 1 To Collection(1).Count
        v = Collection(1)(c)
        GetFieldDictionary.Item(v) = c
    Next
End Function

' ------------- Private function/sub -------------------------------------------

'
' find all separators in csvText
' - RowCount, ColCount = size of array in Csv
' - sepArray1234 = array of field info.    Their size => number of fields + 1.    Index Start with 0, sepArray1234(number of fields) = 0
'                                    SepArray1 = Start pos of field, SepArray2 = field length, SepArray3 = NextSepType, SepArray4 = number of double quotes in field
Private Sub ParseCsv( _
                ByRef RowCount As Long, _
                ByRef ColCount As Long, _
                ByRef SepArray1() As Long, _
                ByRef SepArray2() As Long, _
                ByRef SepArray3() As Long, _
                ByRef SepArray4() As Long, _
                ByRef CsvText As String, _
                Optional ByRef AllowVariableNumOfFields As Boolean = False _
            )
    ' "On Error Resume Next" only if CsvUtilsAnyErrorIsFatal is True
    Err.Clear
    If This.CsvUtilsAnyErrorIsFatal Then GoTo Head
    On Error Resume Next
Head:
    Dim CsvPos As Long
    Dim FieldText As String
    Dim NextSep As Long, NextSepType As Long, QuoteCount As Long, FieldStart As Long, FieldLen As Long
    Dim ColCountTmp As Long
    Dim SepIndex As Long, sepSize As Long
        
    sepSize = UBound(SepArray1)
        
    RowCount = 0
    ColCount = 0                                                                     'max of colomn counts
    ColCountTmp = 0                                                                'current column count
    SepIndex = 0
        
    'for empty Text
    If CsvText = "" Then Exit Sub                                    'return empty collection
        
    ' Add trailing record separator if not
    If VBA.Right(CsvText, 1) <> "" & vbCr And VBA.Right(CsvText, 1) <> "" & vbLf Then
        CsvText = CsvText & vbCrLf
    End If
        
    'extract records and fields
    CsvPos = 1
    Call FindNextSeparatorInit(CsvText)
    Do While FindNextSeparator(CsvText, CsvPos, FieldStart, FieldLen, NextSepType, QuoteCount)
        If Err.Number <> 0 Then Exit Do
                
        ' enhance array size if it is short
        If SepIndex + 1 > sepSize Then
            sepSize = sepSize * 2
            ReDim Preserve SepArray1(sepSize)                    'new elements is initialized by 0
            ReDim Preserve SepArray2(sepSize)
            ReDim Preserve SepArray3(sepSize)
            ReDim Preserve SepArray4(sepSize)
        End If
        SepArray1(SepIndex) = FieldStart
        SepArray2(SepIndex) = FieldLen
        SepArray3(SepIndex) = NextSepType
        SepArray4(SepIndex) = QuoteCount
        SepIndex = SepIndex + 1
                
        ColCountTmp = ColCountTmp + 1
                
        If NextSepType <> 1 Then                                         ' next sep is record separator
            RowCount = RowCount + 1
            If ColCount = 0 Then ColCount = ColCountTmp ' at initial row
            If Not AllowVariableNumOfFields And ColCount <> ColCountTmp Then
                ErrorRaise 10001, "ParseCsvToCollection", "Syntax Error in CSV: numbers of fields are different among records"
                Exit Sub
            End If
            If ColCountTmp > ColCount Then ColCount = ColCountTmp
            ColCountTmp = 0
        End If
    Loop
End Sub

' Find next separator (comma, CR, LF, CRLF) in inText starting with the position "Start"
'     FieldStart = Start position of found field
'     FieldLen     = length of found field
'     Start = found separator + 1 (Start of next field)
'     NextSepType = found separator type (1=comma, 2=CR or CRLF, 3=LF)
'     QuoteCount = double quotation count in found field
'     return False if there is no next separator
'     * found field includes double quote (not yet parsing quotation syntax)
'     * assuming CR or LF exists at EOF
Private Sub FindNextSeparatorInit(ByRef inText As String)
    Dim LenText As Long
    LenText = Len(inText)
    NextSep1 = VBA.InStr(1, inText, This.Delimiter)
    If NextSep1 = 0 Then NextSep1 = LenText + 1        'EOF
    NextSep2 = VBA.InStr(1, inText, "" & vbCr)
    If NextSep2 = 0 Then NextSep2 = LenText + 1        'EOF
    NextSep3 = VBA.InStr(1, inText, "" & vbLf)
    If NextSep3 = 0 Then NextSep3 = LenText + 1        'EOF
End Sub

Private Function FindNextSeparator( _
                     ByRef inText As String, _
                     ByRef Start As Long, _
                     ByRef FieldStart As Long, _
                     ByRef FieldLen As Long, _
                     NextSepType As Long, ByRef QuoteCount As Long _
                 ) As Boolean
    Dim InitStart As Long, LenText As Long
    Dim NextSep As Long, NextStart As Long
        
    FindNextSeparator = False
        
    LenText = Len(inText)
                
    If Start > LenText Then Exit Function                    'over run (no separator found in previous call)
                
    QuoteCount = 0
    FieldStart = Start
        
    Do While Start <= LenText
        ' update nextSep(min of nextSep123), NextSepType, NextStart(next pos of next separator), nextSep123
        If NextSep1 < NextSep2 Then
            If NextSep1 < NextSep3 Then                                ' nextSep1 is smallest
                NextSep = NextSep1
                NextSepType = 1
                NextStart = NextSep + 1
                NextSep1 = InStr(NextStart, inText, This.Delimiter)
                If NextSep1 = 0 Then NextSep1 = LenText + 1 'EOF
            Else                                                                             ' nextSep3 is smallest
                NextSep = NextSep3
                NextSepType = 3
                NextStart = NextSep + 1
                NextSep3 = InStr(NextStart, inText, "" & vbLf)
                If NextSep3 = 0 Then NextSep3 = LenText + 1 'EOF
            End If
        Else
            If NextSep2 < NextSep3 Then                                ' nextSep2 is smallest
                NextSep = NextSep2
                NextSepType = 2
                NextStart = NextSep + 1
                If NextSep3 = NextSep2 + 1 Then                    ' CRLF
                    NextStart = NextStart + 1
                    NextSep3 = InStr(NextStart, inText, "" & vbLf)
                    If NextSep3 = 0 Then NextSep3 = LenText + 1 'EOF
                End If
                NextSep2 = InStr(NextStart, inText, "" & vbCr)
                If NextSep2 = 0 Then NextSep2 = LenText + 1 'EOF
            Else                                                                             ' nextSep3 is smallest
                NextSep = NextSep3
                NextSepType = 3
                NextStart = NextSep + 1
                NextSep3 = InStr(NextStart, inText, "" & vbLf)
                If NextSep3 = 0 Then NextSep3 = LenText + 1 'EOF
            End If
        End If
                
        If NextSep > LenText Then                                        ' separator not found
            Exit Function
        End If
                
        Call StrCount(inText, Start - 1, NextSep - 1, QuoteCount) 'update number of double quates in [FieldStart, nextSep-1]
        Start = NextStart
                
        If QuoteCount Mod 2 = 0 Then                                 'if the number of double-quates is even, then the separator is not fake
            FindNextSeparator = True
            FieldLen = NextSep - FieldStart
            Exit Function
        End If
    Loop
        
    ErrorRaise 10002, "ParseCsvToCollection", "Syntax Error in CSV: illegal double-quote code"
End Function

'
' add number of double quotes in [n+1, p1] of Source to QuoteCount
'
Private Sub StrCount( _
                Source As String, _
                N As Long, _
                p1 As Long, _
                ByRef QuoteCount As Long _
            )
    Dim ss As String
    Dim nn As Long
    Do
        ss = Mid(Source, N + 1, p1 - N)                            ' to avoid from feeding long string to InStr().
        nn = InStr(1, ss, """")
        If nn = 0 Then Exit Do
        N = N + nn
        QuoteCount = QuoteCount + 1
    Loop
End Sub

'
' Trim all before and after doube-quote
'     * Text MUST include two or more double-quotes (")
Private Function TrimQuotes(ByRef Text As String) As String
    'If InStr(Text, """") = 0 Then Err.Raise 9999, "", "program error"
    Dim p0 As Long, p1 As Long
    Dim s As String
        
    'trim tail
    For p1 = Len(Text) To 1 Step -1
        s = Mid(Text, p1, 1)
        If (s = """") Then Exit For
    Next
    'trim head
    For p0 = 1 To p1
        s = Mid(Text, p0, 1)
        If (s = """") Then Exit For
    Next
    'return
    TrimQuotes = Mid(Text, p0 + 1, p1 - p0 - 1)
End Function

