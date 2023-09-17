VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IStringLocalizerBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Interface
Option Explicit

'===============================================================================

Public Function WithLocale( _
                    ByVal LCID As Variant, _
                    ByVal Strings As ILocalizedStrings _
                ) As IStringLocalizerBuilder
End Function

'===============================================================================

Public Function Build() As IStringLocalizer
End Function
