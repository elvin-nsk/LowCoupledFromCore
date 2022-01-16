VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
Option Explicit

Implements IFileSpec

'===============================================================================

Private Type typeThis
  Path As String
  NameWithoutExt As String
  Ext As String
End Type
Private This As typeThis

'===============================================================================

Public Function Create(Optional ByVal File As String) As IFileSpec
  With New FileSpec
    Set Create = .Init(File)
  End With
End Function

Friend Function Init(ByVal File As String) As IFileSpec
  With This
    .Path = GetFilePath(File)
    .NameWithoutExt = GetFileNameWithoutExt(GetFileName(File))
    .Ext = GetFileExt(GetFileName(File))
  End With
  Set Init = Me
End Function

'===============================================================================

Private Property Get IFileSpec_Ext() As String
  IFileSpec_Ext = This.Ext
End Property
Private Property Let IFileSpec_Ext(ByVal RHS As String)
  This.Ext = RHS
End Property

Private Function IFileSpec_ToString() As String
  With This
    IFileSpec_ToString = .Path & .NameWithoutExt & AddDotIfNotEmpty(.Ext)
  End With
End Function

Private Property Get IFileSpec_Name() As String
  With This
    IFileSpec_Name = .NameWithoutExt & AddDotIfNotEmpty(.Ext)
  End With
End Property
Private Property Let IFileSpec_Name(ByVal RHS As String)
  With This
    .NameWithoutExt = GetFileNameWithoutExt(RHS)
    .Ext = GetFileExt(RHS)
  End With
End Property

Private Property Get IFileSpec_NameWithoutExt() As String
  IFileSpec_NameWithoutExt = This.NameWithoutExt
End Property
Private Property Let IFileSpec_NameWithoutExt(ByVal RHS As String)
  This.NameWithoutExt = RHS
End Property

Private Property Get IFileSpec_Path() As String
  IFileSpec_Path = This.Path
End Property
Private Property Let IFileSpec_Path(ByVal RHS As String)
  This.Path = AddProperEndingToPath(RHS)
End Property

Private Property Get IFileSpec_Self() As IFileSpec
  Set IFileSpec_Self = Me
End Property

'===============================================================================

Private Function AddProperEndingToPath(ByVal Path As String) As String
  If Not VBA.Right$(Path, 1) = "\" Then AddProperEndingToPath = Path & "\" _
  Else AddProperEndingToPath = Path
End Function

Private Function GetFileExt(FileName As String) As String
  GetFileExt = VBA.Right$(FileName, _
    VBA.Switch _
      (VBA.InStr(FileName, ".") = 0, _
        0, _
      VBA.InStr(FileName, ".") > 0, _
        VBA.Len(FileName) - VBA.InStrRev(FileName, ".")))
End Function

Private Function AddDotIfNotEmpty(ByVal Ext As String) As String
  If Ext = "" Then AddDotIfNotEmpty = "" Else AddDotIfNotEmpty = "." & Ext
End Function

Private Function GetFileName(ByVal FullFileName As String) As String
  GetFileName = VBA.Right$(FullFileName, VBA.Len(FullFileName) - VBA.InStrRev(FullFileName, "\"))
End Function

Private Function GetFileNameWithoutExt(ByVal FileName As String) As String
  GetFileNameWithoutExt = VBA.Left$(FileName, _
    VBA.Switch _
      (VBA.InStr(FileName, ".") = 0, _
        VBA.Len(FileName), _
      VBA.InStr(FileName, ".") > 0, _
        VBA.InStrRev(FileName, ".") - 1))
End Function

Private Function GetFilePath(ByVal FullFileName As String)
  GetFilePath = VBA.Left$(FullFileName, VBA.InStrRev(FullFileName, "\"))
End Function
