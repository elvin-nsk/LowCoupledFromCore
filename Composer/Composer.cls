VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Composer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
Option Explicit

'===============================================================================

Private Type typeThis
  Elements As Collection
  ComposedElements As Collection
  RemainingElements As Collection
  StartingPoint As IPoint
  MaxPlacesInWidth As Long
  MaxPlacesInHeight As Long
  MaxWidth As Double
  MaxHeight As Double
  HorizontalSpace As Double
  VerticalSpace As Double
  Cursor As IPoint
  CurrentRowItemsInWidth As Long
  CurrentRowWidth As Long
  CurrentRowHeight As Double
  CurrentRowNumber As Long
End Type
Private This As typeThis

'===============================================================================

Public Function CreateAndCompose( _
                  ByVal Elements As Collection, _
                  Optional ByVal StartingPoint As IPoint, _
                  Optional ByVal MaxPlacesInWidth As Long, _
                  Optional ByVal MaxPlacesInHeight As Long, _
                  Optional ByVal MaxWidth As Double, _
                  Optional ByVal MaxHeight As Double, _
                  Optional ByVal HorizontalSpace As Double, _
                  Optional ByVal VerticalSpace As Double _
                ) As Composer
  With New Composer
    Set CreateAndCompose = .Init(Elements, StartingPoint, _
                                 MaxPlacesInWidth, MaxPlacesInHeight, _
                                 MaxWidth, MaxHeight, _
                                 HorizontalSpace, VerticalSpace)
  End With
End Function

Friend Function Init( _
                  ByVal Elements As Collection, _
                  ByVal StartingPoint As IPoint, _
                  ByVal MaxPlacesInWidth As Long, _
                  ByVal MaxPlacesInHeight As Long, _
                  ByVal MaxWidth As Double, _
                  ByVal MaxHeight As Double, _
                  ByVal HorizontalSpace As Double, _
                  ByVal VerticalSpace As Double _
                ) As Composer
  With This
    Set .Elements = Elements
    Set .ComposedElements = New Collection
    Set .RemainingElements = New Collection
    .MaxPlacesInWidth = NumberOrMaxLong(MaxPlacesInWidth)
    .MaxPlacesInHeight = NumberOrMaxLong(MaxPlacesInHeight)
    .MaxWidth = NumberOrMaxDouble(MaxWidth)
    .MaxHeight = NumberOrMaxDouble(MaxHeight)
    .VerticalSpace = VerticalSpace
    .HorizontalSpace = HorizontalSpace
    If StartingPoint Is Nothing Then
      Set .StartingPoint = FreePoint.Create(0, 0)
    Else
      Set .StartingPoint = StartingPoint
    End If
    Set .Cursor = .StartingPoint.GetCopy
  End With
  Compose
  Set Init = Me
End Function

'===============================================================================

Public Property Get ComposedElements() As Collection
  Set ComposedElements = This.ComposedElements
End Property

Public Property Get RemainingElements() As Collection
  Set RemainingElements = This.RemainingElements
End Property

'===============================================================================

Private Sub Compose()
  
  With This
  
    Dim Item As ComposerElement
    .CurrentRowNumber = 1
    For Each Item In .Elements
    
      BeginNextRowIfNeeded Item
      
      If IsHeightExceeded(Item) Then
        .RemainingElements.Add Item
      Else
        PlaceElement Item
        .ComposedElements.Add Item
      End If
    
    Next Item
    
  End With

End Sub

Private Sub PlaceElement(ByVal Item As ComposerElement)
  With This
  
    Item.PivotX = .Cursor.X
    Item.PivotY = .Cursor.Y
    
    .CurrentRowWidth = .CurrentRowWidth + .HorizontalSpace + Item.Width
    If .CurrentRowHeight < Item.Height Then _
      .CurrentRowHeight = Item.Height
    .CurrentRowItemsInWidth = .CurrentRowItemsInWidth + 1
    .Cursor.X = .Cursor.X + .HorizontalSpace + Item.Width
    
  End With
End Sub

Private Function IsHeightExceeded(ByVal Item As ComposerElement) As Boolean
  With This
    If (VBA.Abs(.StartingPoint.Y - .Cursor.Y) + Item.Height >= .MaxHeight) Or _
       (.CurrentRowNumber > .MaxPlacesInHeight) Then _
      IsHeightExceeded = True
  End With
End Function

Private Sub BeginNextRowIfNeeded(ByVal Item As ComposerElement)
  With This
    If (.CurrentRowWidth + Item.Width > .MaxWidth) Or _
       (.CurrentRowItemsInWidth = .MaxPlacesInWidth) Then
      .Cursor.X = .StartingPoint.X
      .Cursor.Y = .Cursor.Y - .CurrentRowHeight - .VerticalSpace
      .CurrentRowWidth = 0
      .CurrentRowHeight = 0
      .CurrentRowItemsInWidth = 0
      .CurrentRowNumber = .CurrentRowNumber + 1
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
