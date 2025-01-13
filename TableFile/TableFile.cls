VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Interface
Option Explicit

'===============================================================================

'@DefaultMember
Public Property Get Cell( _
                        ByVal Row As Long, _
                        ByVal Column As Long _
                    ) As String
Attribute Cell.VB_UserMemId = 0
End Property
Public Property Let Cell( _
                        ByVal Row As Long, _
                        ByVal Column As Long, _
                        RHS As String _
                    )
End Property

Public Sub Dispose()
End Sub

Public Sub ForceLoad()
End Sub

Public Sub ForceSave()
End Sub

Public Property Get MaxRow() As Long
End Property

Public Property Get MaxColumn() As Long
End Property

Public Property Get ReadOnly() As Boolean
End Property

'===============================================================================

Public Property Get Self() As TableFile
End Property
