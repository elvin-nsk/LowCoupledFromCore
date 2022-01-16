VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@ModuleDescription "Low-coupled version from Core framework"
'@Interface
Option Explicit

'===============================================================================

Public Sub Add(ByRef Item As Variant, _
               Optional ByVal Before As Long, _
               Optional ByVal After As Long)
End Sub

Public Sub Append(ByVal List As IList, _
                  Optional ByVal Before As Long, _
                  Optional ByVal After As Long)
End Sub

Public Sub AppendCollection(ByVal Collection As VBA.Collection, _
                            Optional ByVal Before As Long, _
                            Optional ByVal After As Long)
End Sub

Public Sub Remove(ByVal Index As Long)
End Sub

Public Sub RemoveList(ByVal List As IList)
End Sub

Public Function IndexOf(ByRef Item As Variant) As Long
End Function

Public Function Intersect(ByVal List As IList) As IList
End Function

Public Function GetCopy() As IList
End Function

Public Function Contains(ByRef Item As Variant) As Boolean
End Function

Public Property Get Count() As Long
End Property

'@DefaultMember
Public Property Get Item(ByVal Index As Long) As Variant
Attribute Item.VB_UserMemId = 0
End Property
Public Property Let Item(ByVal Index As Long, ByRef RHS As Variant)
End Property
Public Property Set Item(ByVal Index As Long, ByVal RHS As Variant)
End Property

Public Property Get NewEnum() As EnumHelper
End Property

'===============================================================================

Public Property Get Self() As IList
End Property
