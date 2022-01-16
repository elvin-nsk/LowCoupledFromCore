VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Stacker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
Option Explicit

'===============================================================================

Private Type typeThis
  Stackables As Collection
  StartingPoint As IPoint
  MaxPlacesInWidth As Long
  MaxWidth As Double
  HorizontalSpace As Double
  VerticalSpace As Double
  Cursor As IPoint
  CurrentRowItemsInWidth As Long
  CurrentRowWidth As Long
  CurrentRowHeight As Double
End Type
Private This As typeThis

'===============================================================================

Public Function CreateAndStack( _
                  ByVal Stackables As Collection, _
                  Optional ByVal StartingPoint As IPoint, _
                  Optional ByVal MaxPlacesInWidth As Long, _
                  Optional ByVal MaxWidth As Double, _
                  Optional ByVal HorizontalSpace As Double, _
                  Optional ByVal VerticalSpace As Double _
                ) As Stacker
  With New Stacker
    Set CreateAndStack = .Init(Stackables, StartingPoint, _
                               MaxPlacesInWidth, MaxWidth, _
                               HorizontalSpace, VerticalSpace)
  End With
End Function

Friend Function Init( _
                  ByVal Stackables As Collection, _
                  ByVal StartingPoint As IPoint, _
                  ByVal MaxPlacesInWidth As Long, _
                  ByVal MaxWidth As Double, _
                  ByVal HorizontalSpace As Double, _
                  ByVal VerticalSpace As Double _
                ) As Stacker
  With This
    Set .Stackables = Stackables
    .MaxPlacesInWidth = NumberOrMaxLong(MaxPlacesInWidth)
    .MaxWidth = NumberOrMaxLong(MaxWidth)
    .VerticalSpace = VerticalSpace
    .HorizontalSpace = HorizontalSpace
    If StartingPoint Is Nothing Then
      Set .StartingPoint = FreePoint.Create(0, 0)
    Else
      Set .StartingPoint = StartingPoint
    End If
    Set .Cursor = .StartingPoint.GetCopy
  End With
  Stack
  Set Init = Me
End Function

'===============================================================================

Private Sub Stack()
  
  With This
  
    Dim Item As Stackable
    For Each Item In .Stackables
    
      BeginNextRowIfNeeded Item
    
      Item.PivotX = .Cursor.X
      Item.PivotY = .Cursor.Y
      
      .CurrentRowWidth = .CurrentRowWidth + .HorizontalSpace + Item.Width
      If .CurrentRowHeight < Item.Height Then _
        .CurrentRowHeight = Item.Height
      .CurrentRowItemsInWidth = .CurrentRowItemsInWidth + 1
      .Cursor.X = .Cursor.X + .HorizontalSpace + Item.Width
    
    Next Item
    
  End With

End Sub

Private Sub BeginNextRowIfNeeded(ByVal Item As Stackable)
  With This
    If (.CurrentRowWidth + Item.Width > .MaxWidth) Or _
       (.CurrentRowItemsInWidth = .MaxPlacesInWidth) Then
      .Cursor.X = .StartingPoint.X
      .Cursor.Y = .Cursor.Y - .CurrentRowHeight - .VerticalSpace
      .CurrentRowWidth = 0
      .CurrentRowHeight = 0
      .CurrentRowItemsInWidth = 0
    End If
  End With
End Sub

Private Function NumberOrMaxLong(ByVal Number As Long) As Long
  If Number <= 0 Then
    NumberOrMaxLong = 2147483647
  Else
    NumberOrMaxLong = Number
  End If
End Function

Private Function NumberOrMaxDouble(ByVal Number As Double) As Double
  If Number <= 0 Then
    NumberOrMaxDouble = 1.79769313486231E+308
  Else
    NumberOrMaxDouble = Number
  End If
End Function
