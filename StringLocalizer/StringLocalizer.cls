VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringLocalizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
Option Explicit

Implements IStringLocalizer
Implements IStringLocalizerBuilder

'===============================================================================

Private Const CustomError = vbObjectError Or 32

Private Type typeThis
    StringPacks As Scripting.IDictionary
    CurrentLCID As String
    FallbackLCID As String
End Type
Private This As typeThis

'===============================================================================

Public Function Builder( _
                    ByVal FallbackLCID As Variant, _
                    ByVal FallbackStrings As ILocalizedStrings _
                ) As IStringLocalizerBuilder
    With New StringLocalizer
        Set Builder = .Init(VBA.CStr(FallbackLCID), FallbackStrings)
    End With
End Function

Friend Function Init( _
                    ByVal FallbackLCID As Variant, _
                    ByVal FallbackStrings As ILocalizedStrings _
                ) As IStringLocalizerBuilder
    With This
        .FallbackLCID = FallbackLCID
        Set .StringPacks = New Scripting.Dictionary
        .StringPacks.Add FallbackLCID, FallbackStrings
    End With
    Set Init = Me
End Function

'-------------------------------------------------------------------------------

Private Function IStringLocalizerBuilder_Build() As IStringLocalizer
    Set IStringLocalizerBuilder_Build = Me
    trySetCurrentLCID
End Function

Private Function IStringLocalizerBuilder_WithLocale( _
                     ByVal LCID As Variant, _
                     ByVal Strings As ILocalizedStrings _
                 ) As IStringLocalizerBuilder
    This.StringPacks.Add VBA.CStr(LCID), Strings
End Function

'===============================================================================

Private Property Get IStringLocalizer_CurrentLocale() As String
    IStringLocalizer_CurrentLocale = This.CurrentLCID
End Property

Private Sub IStringLocalizer_ForceLocale(ByVal LCID As Variant)
    Dim StrLCID As String
    StrLCID = VBA.CStr(LCID)
    If This.StringPacks.Exists(StrLCID) Then
        This.CurrentLCID = StrLCID
    Else
        ThrowNoStringsForLCID StrLCID
    End If
End Sub

Private Property Get IStringLocalizer_Item( _
                         ByVal Key As String, _
                         ParamArray Args() As Variant _
                     ) As String
    IStringLocalizer_Item = GetString(Key)
    Dim ArgumentsToPass As Variant
    If Not (LBound(Args) = 0 And UBound(Args) = -1) Then
        ArgumentsToPass = Args
        IStringLocalizer_Item = _
            StringFormat(IStringLocalizer_Item, ArgumentsToPass)
    End If
End Property

Private Function GetString(ByVal Key As String) As String
    With This
        If .StringPacks(.CurrentLCID).Strings.Exists(Key) Then
            GetString = This.StringPacks(.CurrentLCID).Strings(Key)
        Else
            GetString = This.StringPacks(.FallbackLCID).Strings(Key)
        End If
    End With
End Function


Private Property Get IStringLocalizer_Self() As IStringLocalizer
    Set IStringLocalizer_Self = Me
End Property

'===============================================================================

Private Function trySetCurrentLCID() As Boolean
    Dim ContextLCID As String
    ContextLCID = GetContextLCID
    If This.StringPacks.Exists(ContextLCID) Then
        This.CurrentLCID = ContextLCID
        trySetCurrentLCID = True
    Else
        This.CurrentLCID = This.FallbackLCID
    End If
End Function

Private Function GetContextLCID() As String
    GetContextLCID = VBA.CStr(UILanguage)
End Function

Private Function StringFormat( _
                     ByVal Text As String, _
                     ByRef Args As Variant _
                 ) As String
'    codekabinett.com -
'    You may use, modify, copy, distribute this code as long as this line remains
    
    Dim i As Long
    Dim StrRetVal As String
    Dim StartPos As Long
    Dim EndPos As Long
    Dim FormatString As String
    Dim ArgValueLen As Long
    StrRetVal = Text
    
    For i = LBound(Args) To UBound(Args)
        ArgValueLen = VBA.Len(VBA.CStr(i))
        StartPos = VBA.InStr(StrRetVal, "{" & VBA.CStr(i) & ":")
        If StartPos > 0 Then
            EndPos = VBA.InStr(StartPos + 1, StrRetVal, "}")
            FormatString = VBA.Mid( _
                                   StrRetVal, _
                                   StartPos + 2 + ArgValueLen, _
                                   EndPos - (StartPos + 2 + ArgValueLen) _
                               )
            StrRetVal = VBA.Mid( _
                                StrRetVal, 1, StartPos - 1) & _
                                VBA.Format(Nz(Args(i), ""), FormatString) & _
                                VBA.Mid(StrRetVal, EndPos + 1 _
                            )
        Else
            StrRetVal = VBA.Replace(StrRetVal, "{" & VBA.CStr(i) & "}", Nz(Args(i), ""))
        End If
    Next i
    
    StringFormat = StrRetVal

End Function

Private Function Nz( _
                     ByRef Value As Variant, _
                     Optional ByRef ValueIfNull As Variant = Empty _
                 ) As Variant
    If VBA.IsNull(Value) Or VBA.IsEmpty(Value) Then
        AssignUnknown Nz, ValueIfNull
    Else
        AssignUnknown Nz, Value
    End If
End Function

Private Sub AssignUnknown( _
                ByRef Destination As Variant, _
                ByRef Value As Variant _
            )
    If VBA.IsObject(Value) Then
        Set Destination = Value
    Else
        Destination = Value
    End If
End Sub

Private Sub ThrowNoStringsForLCID(ByVal LCID As String)
    VBA.Err.Raise CustomError, "StringLocalizer", "NoStringsForLCID " & LCID
End Sub
