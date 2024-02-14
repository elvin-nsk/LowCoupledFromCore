VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MarksSetter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'===============================================================================
'   Модуль          : MarksSetter
'   Версия          : 2024.02.13
'   Автор           : elvin-nsk (me@elvin.nsk.ru)
'   Сайт            : https://github.com/elvin-nsk/LowCoupledFromCore
'   Назначение      : Расставляет метки реза
'   Зависимости     : LibCore
'===============================================================================

'@PredeclaredId
Option Explicit

'===============================================================================
' # Declarations

Private Type This
    Bleeds As Double
    MarksLayerColor As Color
    MarksLayerName As String
    MarkColor As Color
    MarkThickness As Double
    PlaceholderGroupPrefix As String
    PlaceWhiteUndermark As Boolean
    Shapes As ShapeRange
    Size As Double
    UndermarkColor As Color
    UndermarkThickness As Double
End Type
Private This As This

'===============================================================================
' # Constructors

Public Function New_( _
                    ByVal Shapes As ShapeRange _
                ) As MarksSetter
    Set New_ = New MarksSetter
    New_.Inject Shapes
End Function

Friend Sub Inject( _
               ByVal Shapes As ShapeRange _
           )
    With This
        .Bleeds = 0
        Set .MarksLayerColor = CreateCMYKColor(60, 0, 20, 0)
        .MarksLayerName = "метки"
        Set .MarkColor = CreateRegistrationColor
        .MarkThickness = 0.1
        .PlaceWhiteUndermark = False
        .PlaceholderGroupPrefix = "место"
        Set .Shapes = Shapes
        .Size = 3
        Set .UndermarkColor = CreateRegistrationColor
        .UndermarkColor.Tint = 0
        .UndermarkThickness = 1.5
    End With
End Sub

'===============================================================================
' # Public

Public Property Get Bleeds() As Double
    Bleeds = This.Bleeds
End Property
Public Property Let Bleeds(ByVal RHS As Double)
    This.Bleeds = RHS
End Property

Public Property Get MarkThickness() As Double
    MarkThickness = This.MarkThickness
End Property
Public Property Let MarkThickness(ByVal RHS As Double)
    This.MarkThickness = RHS
End Property

Public Property Get PlaceholderGroupPrefix() As String
    PlaceholderGroupPrefix = This.PlaceholderGroupPrefix
End Property
Public Property Let PlaceholderGroupPrefix(ByVal RHS As String)
    This.PlaceholderGroupPrefix = RHS
End Property

Public Property Get PlaceWhiteUndermark() As Boolean
    PlaceWhiteUndermark = This.PlaceWhiteUndermark
End Property
Public Property Let PlaceWhiteUndermark(ByVal RHS As Boolean)
    This.PlaceWhiteUndermark = RHS
End Property

Public Function SetMarksBelow() As ShapeRange
    Set SetMarksBelow = SetMarks(This.Shapes)
    SetMarksBelow.OrderBackOf GetBottomOrderShape(This.Shapes)
End Function

Public Function SetMarksOnNewLayer() As ShapeRange
    With ActivePage.CreateLayer(This.MarksLayerName)
        .MoveBelow ShapesLayer(This.Shapes)
        .Color.CopyAssign This.MarksLayerColor
        .Activate
    End With
    Set SetMarksOnNewLayer = SetMarks(This.Shapes)
End Function

Public Property Get Size() As Double
    Size = This.MarkThickness
End Property
Public Property Let Size(ByVal RHS As Double)
    This.Size = RHS
End Property

Public Property Get UndermarkThickness() As Double
    UndermarkThickness = This.UndermarkThickness
End Property
Public Property Let UndermarkThickness(ByVal RHS As Double)
    This.UndermarkThickness = RHS
End Property

Public Property Get Self() As MarksSetter
    Set Self = Me
End Property

'===============================================================================
' # Helpers

Private Function SetMarks( _
                    ByVal Shapes As ShapeRange _
                ) As ShapeRange
    ActiveDocument.Unit = cdrMillimeter
    Set SetMarks = CreateShapeRange
    Dim Box As Rect
    Dim Shape As Shape
    For Each Shape In Shapes
        SetMarks.Add DrawPlace(SpaceBox(Shape, -This.Bleeds))
    Next Shape
End Function

Private Function DrawPlace( _
                    ByVal TrimBox As Rect _
                ) As Shape
    Dim Place As New ShapeRange
    Place.Add DrawTrimBox(TrimBox)
    Dim Mark As Shape
    Set Mark = DrawCutMark(TrimBox.Left, TrimBox.Top)
    Place.Add Mark
    Set Mark = DuplicateAndSetAngle(Mark, 0, 0, 90)
    Place.Add Mark
    Set Mark = DuplicateAndSetAngle(Mark, TrimBox.Width, 0, 0)
    Place.Add Mark
    Set Mark = DuplicateAndSetAngle(Mark, 0, 0, -90)
    Place.Add Mark
    Set Mark = DuplicateAndSetAngle(Mark, 0, -TrimBox.Height, 180)
    Place.Add Mark
    Set Mark = DuplicateAndSetAngle(Mark, 0, 0, -90)
    Place.Add Mark
    Set Mark = DuplicateAndSetAngle(Mark, -TrimBox.Width, 0, 180)
    Place.Add Mark
    Set Mark = DuplicateAndSetAngle(Mark, 0, 0, 90)
    Place.Add Mark
    Set DrawPlace = Place.Group
    DrawPlace.Name = This.PlaceholderGroupPrefix & " " _
                   & FixSize(TrimBox.Width) & "x" _
                   & FixSize(TrimBox.Height)
End Function

Private Function DrawTrimBox( _
                     ByVal TrimBox As Rect _
                 ) As Shape
    Set DrawTrimBox = ActiveLayer.CreateRectangleRect(TrimBox)
    DrawTrimBox.ConvertToCurves
    DrawTrimBox.Outline.SetNoOutline
    DrawTrimBox.Fill.ApplyNoFill
End Function

Private Function DuplicateAndSetAngle( _
                     ByVal Shape As Shape, _
                     ByVal OffsetX As Double, _
                     ByVal OffsetY As Double, _
                     ByVal Angle As Double _
                 ) As Shape
    Set DuplicateAndSetAngle = Shape.Duplicate(OffsetX, OffsetY)
    DuplicateAndSetAngle.RotationAngle = Angle
End Function

Private Function DrawCutMark( _
                     ByVal PivotX As Double, _
                     ByVal PivotY As Double, _
                     Optional ByVal Angle As Double = 0 _
                 ) As Shape
    Dim Mark As New ShapeRange
    Dim Part As Shape
    With ActiveLayer
        Set Part = _
            .CreateLineSegment( _
                PivotX, PivotY, _
                PivotX, PivotY + This.Bleeds _
            )
        Part.Outline.SetNoOutline
        Mark.Add Part
        Set Part = _
            .CreateLineSegment( _
                PivotX, PivotY + This.Bleeds, _
                PivotX, PivotY + This.Bleeds + This.Size _
            )
        Part.Outline.Color.CopyAssign This.MarkColor
        Part.Outline.SetProperties Width:=This.MarkThickness
        Mark.Add Part
        If This.PlaceWhiteUndermark Then
            Set Part = Part.Duplicate
            Part.OrderBackOne
            Part.Outline.Color.CopyAssign This.UndermarkColor
            Part.Outline.SetProperties Width:=This.UndermarkThickness
            Mark.Add Part
        End If
        Set DrawCutMark = Mark.Group
        DrawCutMark.SetRotationCenter PivotX, PivotY
        DrawCutMark.RotationAngle = Angle
    End With
End Function

Private Property Get ShapesLayer(ByVal Shapes As ShapeRange) As Layer
    Set ShapesLayer = Shapes.FirstShape.Layer
End Property

Private Property Get FixSize(ByVal Size As Double) As Double
    FixSize = VBA.Round(Size, 2)
End Property
