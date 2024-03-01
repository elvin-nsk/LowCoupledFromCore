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
'   Версия          : 2024.02.12
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

Friend Function New_(Optional ByVal File As String) As FileSpec
    Set New_ = New FileSpec
    New_.Inject File
End Function

Friend Sub Inject( _
               Optional ByVal File As String _
           )
    With This
        RefreshFilePath File
        RefreshBaseName GetFileName(File)
        RefreshFileExt GetFileName(File)
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

Public Property Get Self() As FileSpec
    Set Self = Me
End Property

'===============================================================================
' # Helpers

Private Sub RefreshFileExt(ByVal FileName As String)
    This.Ext = _
        VBA.Right$( _
            FileName, _
            VBA.Switch( _
                VBA.InStr(FileName, ".") = 0, _
                    0, _
                VBA.InStr(FileName, ".") > 0, _
                    VBA.Len(FileName) - VBA.InStrRev(FileName, ".")) _
        )
End Sub

Private Sub RefreshBaseName(ByVal FileName As String)
    This.BaseName = _
        VBA.Left$( _
            FileName, _
            VBA.Switch( _
                VBA.InStr(FileName, ".") = 0, _
                    VBA.Len(FileName), _
                VBA.InStr(FileName, ".") > 0, _
                    VBA.InStrRev(FileName, ".") - 1) _
        )
End Sub

Private Sub RefreshFilePath(ByVal FullFileName As String)
    This.Path = VBA.Left$(FullFileName, VBA.InStrRev(FullFileName, "\"))
End Sub

Private Function AddProperEndingToPath(ByVal Path As String) As String
    If Not VBA.Right$(Path, 1) = "\" Then AddProperEndingToPath = Path & "\" _
    Else AddProperEndingToPath = Path
End Function

Private Function AddDotIfNotEmpty(ByVal Ext As String) As String
    If Ext = vbNullString Then
        AddDotIfNotEmpty = vbNullString
    Else
        AddDotIfNotEmpty = "." & Ext
    End If
End Function

Private Function GetFileName(ByVal FullFileName As String) As String
    GetFileName = _
        VBA.Right$( _
            FullFileName, _
            VBA.Len(FullFileName) - VBA.InStrRev(FullFileName, "\") _
        )
End Function
