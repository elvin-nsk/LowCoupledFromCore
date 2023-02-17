VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
Option Explicit

'===============================================================================

Private Const OFN_ALLOWMULTISELECT = &H200&
Private Const OFN_EXPLORER = &H80000
Private Const OFN_FILEMUSTEXIST = &H1000&
Private Const OFN_HIDEREADONLY = &H4&
Private Const OFN_PATHMUSTEXIST = &H800&

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As LongPtr
    hInstance As LongPtr
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As LongPtr
    lpTemplateName As String
End Type
Private OpenFile As OPENFILENAME

#If VBA7 Then
    Private Declare PtrSafe Function GetOpenFileName Lib "comdlg32.dll" Alias _
        "GetOpenFileNameA" (pOpenFileName As OPENFILENAME) As Long
#Else
    Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias _
        "GetOpenFileNameA" (pOpenFileName As typeOpenFile) As Long
#End If

Private Type typeThis
    Filter As String
    InitialDir As String
    MultiSelect As Boolean
    Title As String
End Type
Private This As typeThis

'===============================================================================

Private Sub Class_Initialize()
    With This
        .Filter = "CorelDRAW (*.cdr)" & Chr(0) & "*.cdr" _
                & Chr(0) & "All Files (*.*)" & Chr(0) & "*.*"
        .InitialDir = "%USERPROFILE%\Desktop"
        .MultiSelect = True
        .Title = "Открыть..."
    End With
End Sub

'===============================================================================

Public Property Get Filter() As String
    Filter = This.Filter
End Property
Public Property Let Filter(ByVal RHS As String)
    This.Filter = RHS
End Property

Public Property Get InitialDir() As String
    InitialDir = This.InitialDir
End Property
Public Property Let InitialDir(ByVal RHS As String)
    This.InitialDir = RHS
End Property

Public Property Get MultiSelect() As Boolean
    MultiSelect = This.MultiSelect
End Property
Public Property Let MultiSelect(ByVal RHS As Boolean)
    This.MultiSelect = RHS
End Property

Public Property Get Title() As String
    Title = This.Title
End Property
Public Property Let Title(ByVal RHS As String)
    This.Title = RHS
End Property

Public Function ShowFileOpenDialog() As Collection

    Set ShowFileOpenDialog = New Collection
    
    Dim lReturn As Long
    Dim FileDir As String
    Dim FilePos As Long
    Dim PrevFilePos As Long
    
    With OpenFile
        .lStructSize = LenB(OpenFile) 'VBA.LenB вызывает в этом месте ошибку (!)
        .hwndOwner = 0
        .hInstance = 0
        'The Filter starts with a FileType Name then the File Extension with Chr(0) in between.
        'Multiple FileType Names & Extensions have a Chr(0) between each.
        'Mutiple extensions for the same FileType must have semi-colons between each extension.
        '.lpstrFilter = "CorelDRAW (*.cdr)" & Chr(0) & "*.cdr"
        '.lpstrFilter = "Images (jpg, png, gif, pcx)" & Chr(0) & "*.jpg;*.png;*.gif;*.pcx"
        .lpstrFilter = This.Filter
        .nFilterIndex = 1
        .lpstrFile = String(4096, 0)
        .nMaxFile = VBA.Len(.lpstrFile) - 1
        .lpstrFileTitle = .lpstrFile
        .nMaxFileTitle = .nMaxFile
        'Insert the initial directory where you want the File Dialog box to open.
        .lpstrInitialDir = This.InitialDir
        .lpstrTitle = This.Title
        'If multiple selection of files is not allowed do not add the value of OFN_ALLOWMULTISELECT to .flags
        .Flags = OFN_HIDEREADONLY + _
                 OFN_PATHMUSTEXIST + _
                 OFN_FILEMUSTEXIST + _
                 OFN_EXPLORER + _
                 VBA.Switch(This.MultiSelect, OFN_ALLOWMULTISELECT, Not This.MultiSelect, 0)
        lReturn = GetOpenFileName(OpenFile)
        
        If lReturn = 0 Then Exit Function
        FilePos = VBA.InStr(1, .lpstrFile, VBA.Chr(0))
        If VBA.Mid(.lpstrFile, FilePos + 1, 1) = VBA.Chr(0) Then
            ShowFileOpenDialog.Add .lpstrFile
        Else
            FileDir = VBA.Mid(.lpstrFile, 1, FilePos - 1)
            Do While True
                PrevFilePos = FilePos
                FilePos = InStr(PrevFilePos + 1, .lpstrFile, VBA.Chr(0))
                If FilePos - PrevFilePos <= 1 Then Exit Do
                ShowFileOpenDialog.Add FileDir + "\" + _
                    VBA.Mid(.lpstrFile, PrevFilePos + 1, _
                        FilePos - PrevFilePos - 1)
                Loop
        End If
    End With
End Function

Public Property Get Self() As FileBrowser
    Set Self = Me
End Property

'===============================================================================

