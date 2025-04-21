VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'===============================================================================
'   Модуль          : FileSpec
'   Версия          : 2025.02.28
'   Автор           : elvin-nsk (me@elvin.nsk.ru)
'   Сайт            : https://github.com/elvin-nsk/LowCoupledFromCore
'   Назначение      : Хранит имя файла и путь
'   Зависимости     : Нет
'===============================================================================

'@Interface
'@PredeclaredId
Option Explicit

'===============================================================================
' # Declarations

Private Type This
    Path As String
    BaseName As String
    Ext As String
End Type
Private This As This

'===============================================================================
' # Constructors

Friend Function Build( _
                    Optional ByVal FullFileName As String, _
                    Optional ByVal Path As String, _
                    Optional ByVal Name As String, _
                    Optional ByVal Ext As String _
                ) As FileSpec
    Set Build = New FileSpec
    If Not FullFileName = vbNullString Then Build.Inject FullFileName
    If Not Path = vbNullString Then Build.Path = Path
    If Not Name = vbNullString Then Build.Name = Name
    If Not Ext = vbNullString Then Build.Ext = Ext
End Function

Friend Function New_(Optional ByVal FullFileName As String) As FileSpec
    Set New_ = New FileSpec
    If Not FullFileName = vbNullString Then New_.Inject FullFileName
End Function

Friend Sub Inject( _
               Optional ByVal FullFileName As String _
           )
    With This
        RefreshFilePath FullFileName
        RefreshBaseName FileName(FullFileName)
        RefreshFileExt FileName(FullFileName)
    End With
End Sub

'===============================================================================
' # Public

Public Property Get Ext() As String
    Ext = This.Ext
End Property
Public Property Let Ext(ByVal RHS As String)
    This.Ext = RHS
End Property

Public Property Get GetCopy() As FileSpec
    With New FileSpec
        .Path = This.Path
        .BaseName = This.BaseName
        .Ext = This.Ext
        Set GetCopy = .Self
    End With
End Property

'@DefaultMember
Public Function ToString() As String
Attribute ToString.VB_UserMemId = 0
    With This
        ToString = .Path & .BaseName & AddDotIfNotEmpty(.Ext)
    End With
End Function

Public Property Get Name() As String
    With This
        Name = .BaseName & AddDotIfNotEmpty(.Ext)
    End With
End Property
Public Property Let Name(ByVal RHS As String)
    With This
        RefreshBaseName RHS
        RefreshFileExt RHS
    End With
End Property

Public Property Get BaseName() As String
    BaseName = This.BaseName
End Property
Public Property Let BaseName(ByVal RHS As String)
    This.BaseName = RHS
End Property

Public Property Get Path() As String
    Path = This.Path
End Property
Public Property Let Path(ByVal RHS As String)
    This.Path = AddProperEndingToPath(RHS)
End Property

Public Property Get RightFolder() As String
    Dim xs() As String: xs = Split(This.Path, "\")
    RightFolder = xs(UBound(xs) - 1)
End Property

Public Property Get Self() As FileSpec
    Set Self = Me
End Property

'===============================================================================
' # Helpers

Private Sub RefreshFileExt(ByVal FullFileName As String)
    This.Ext = _
        VBA.Right$( _
            FullFileName, _
            VBA.Switch( _
                VBA.InStr(FullFileName, ".") = 0, _
                    0, _
                VBA.InStr(FullFileName, ".") > 0, _
                    VBA.Len(FullFileName) - VBA.InStrRev(FullFileName, ".")) _
        )
End Sub

Private Sub RefreshBaseName(ByVal FullFileName As String)
    This.BaseName = _
        VBA.Left$( _
            FullFileName, _
            VBA.Switch( _
                VBA.InStr(FullFileName, ".") = 0, _
                    VBA.Len(FullFileName), _
                VBA.InStr(FullFileName, ".") > 0, _
                    VBA.InStrRev(FullFileName, ".") - 1) _
        )
End Sub

Private Sub RefreshFilePath(ByVal FullFileName As String)
    This.Path = VBA.Left(FullFileName, VBA.InStrRev(FullFileName, "\"))
End Sub

Private Function AddProperEndingToPath(ByVal Path As String) As String
    If Not VBA.Right(Path, 1) = "\" Then AddProperEndingToPath = Path & "\" _
    Else AddProperEndingToPath = Path
End Function

Private Function AddDotIfNotEmpty(ByVal Ext As String) As String
    If Ext = vbNullString Then
        AddDotIfNotEmpty = vbNullString
    Else
        AddDotIfNotEmpty = "." & Ext
    End If
End Function

Private Function FileName(ByVal FullFileName As String) As String
    FileName = _
        VBA.Right$( _
            FullFileName, _
            Len(FullFileName) - VBA.InStrRev(FullFileName, "\") _
        )
End Function
