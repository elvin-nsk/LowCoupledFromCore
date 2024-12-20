Attribute VB_Name = "LibCore"
'===============================================================================
'   Модуль          : LibCore
'   Версия          : 2024.12.04
'   Автор           : elvin-nsk (me@elvin.nsk.ru)
'   Использован код : dizzy (из макроса CtC), Alex Vakulenko
'                     и др.
'   Описание        : библиотека функций для макросов
'   Использование   :
'   Зависимости     : самодостаточный
'===============================================================================

Option Explicit
Option Base 1

'===============================================================================
' # приватные переменные модуля

Private Type LayerProperties
    Visible As Boolean
    Printable As Boolean
    Editable As Boolean
End Type

Private StartTime As Double

'===============================================================================
' # публичные переменные

Public Type typeMatrix
    d11 As Double
    d12 As Double
    d21 As Double
    d22 As Double
    tx As Double
    ty As Double
End Type

Public Enum ErrorCodes
    ErrorInvalidArgument = 5
    ErrorTypeMismatch = 13
End Enum

Public Const CUSTOM_ERROR = vbObjectError Or 32

'===============================================================================
' # мнимые константы

Public Property Get None() As Variant
End Property
Public Property Let None(RHS As Variant)
End Property

Public Property Get Cyan() As Color
    Set Cyan = CreateCMYKColor(100, 0, 0, 0)
End Property

Public Property Get Magenta() As Color
    Set Magenta = CreateCMYKColor(0, 100, 0, 0)
End Property

Public Property Get Yellow() As Color
    Set Yellow = CreateCMYKColor(0, 0, 100, 0)
End Property

Public Property Get Black() As Color
    Set Black = CreateCMYKColor(0, 0, 0, 100)
End Property

Public Property Get GrayBlack() As Color
    Set GrayBlack = CreateGrayColor(0)
End Property

Public Property Get GrayWhite() As Color
    Set GrayWhite = CreateGrayColor(255)
End Property

'===============================================================================
' # функции поиска и получения информации об объектах корела

Public Property Get AllNodesInside( _
                        ByVal Nodes As NodeRange, _
                        ByVal Curve As Curve _
                    ) As Boolean
    Dim Node As Node
    For Each Node In Nodes
        If Not Curve.IsPointInside(Node.PositionX, Node.PositionY) Then _
            Exit Property
    Next Node
    AllNodesInside = True
End Property

'возвращает среднее сторон шейпа/рэйнджа/страницы
Public Property Get AverageDim(ByVal ShapeOrRangeOrPage As Object) As Double
    If Not TypeOf ShapeOrRangeOrPage Is Shape _
   And Not TypeOf ShapeOrRangeOrPage Is ShapeRange _
   And Not TypeOf ShapeOrRangeOrPage Is Page Then
        Err.Raise _
            13, Source:="AverageDim", _
            Description:="Type mismatch: ShapeOrRangeOrPage должен быть Shape, ShapeRange или Page"
        Exit Property
    End If
    AverageDim = (ShapeOrRangeOrPage.SizeWidth + ShapeOrRangeOrPage.SizeHeight) _
               / 2
End Property

'находит все шейпы, включая шейпы в поверклипах, с рекурсией,
'опционально исключая шейпы-поверклипы и шейпы-группы
Public Property Get FindAllShapes( _
                        ByVal Shapes As ShapeRange, _
                        Optional ExcludeGroupShapes As Boolean = False, _
                        Optional ExcludePowerClipShapes As Boolean = False _
                    ) As ShapeRange
    Dim Shape As Shape
    Set FindAllShapes = CreateShapeRange
    FindAllShapes.AddRange Shapes.Shapes.FindShapes
    For Each Shape In FindPowerClips(Shapes)
        FindAllShapes.AddRange FindAllShapes(Shape.PowerClip.Shapes.All)
    Next Shape
    If ExcludeGroupShapes Then
        FindAllShapes.RemoveRange _
            FindAllShapes.Shapes.FindShapes(Type:=cdrGroupShape)
    End If
    If ExcludePowerClipShapes Then
        FindAllShapes.RemoveRange _
            FindPowerClips(FindAllShapes)
    End If
End Property

Public Property Get FindColor( _
                        ByVal PaletteName As String, _
                        ByVal ColorName As String _
                    ) As Color
    Dim Palette As Palette
    Set Palette = PaletteManager.GetPalette(PaletteName)
    If Palette Is Nothing Then Exit Property
    Dim Index As Long
    Index = Palette.FindColor(ColorName)
    If Index = 0 Then Exit Property
    Set FindColor = Palette.Color(Index)
End Property

Public Property Get FindShapesByFillColor( _
                        ByVal Shapes As ShapeRange, _
                        ByVal Color As Color _
                    ) As ShapeRange
    Dim Result As New ShapeRange
    Dim Shape As Shape
    For Each Shape In FindAllShapes(Shapes)
        If Shape.Fill.Type = cdrUniformFill Then
            If Shape.Fill.UniformColor.IsSame(Color) Then
                Result.Add Shape
            End If
        End If
    Next Shape
    Set FindShapesByFillColor = Result
End Property

'находит все шейпы с данным именем, включая шейпы в поверклипах, с рекурсией
Public Property Get FindShapesByName( _
                        ByVal Shapes As ShapeRange, _
                        ByVal Name As String _
                    ) As ShapeRange
    Set FindShapesByName = _
        FindAllShapes(Shapes).Shapes.FindShapes( _
            Name:=Name, Recursive:=False _
        )
End Property

'находит все шейпы с данными именами, включая шейпы в поверклипах, с рекурсией
'и помещает их в как рейнджи в словарь, где ключ = имя
Public Property Get FindShapesByNames( _
                        ByVal Shapes As ShapeRange, _
                        ByRef NamesSequence As Variant _
                    ) As Scripting.IDictionary
    Dim Dic As New Scripting.Dictionary
    Dim Name As Variant
    For Each Name In NamesSequence
        Dic.Add Name, FindShapesByName(Shapes, Name)
    Next Name
    Set FindShapesByNames = Dic
End Property

'находит все шейпы, часть имени которых совпадает с NamePart,
'включая шейпы в поверклипах, с рекурсией
Public Property Get FindShapesByNamePart( _
                        ByVal Shapes As ShapeRange, _
                        ByVal NamePart As String _
                    ) As ShapeRange
    Set FindShapesByNamePart = _
        FindAllShapes(Shapes).Shapes.FindShapes( _
            Query:="@Name.Contains('" & NamePart & "')" _
        )
End Property

Public Property Get FindShapesByOutlineColor( _
                        ByVal Shapes As ShapeRange, _
                        ByVal Color As Color _
                    ) As ShapeRange
    Dim Result As New ShapeRange
    Dim Shape As Shape
    For Each Shape In FindAllShapes(Shapes)
        If Shape.Outline.Type = cdrOutline Then
            If Shape.Outline.Color.IsSame(Color) Then
                Result.Add Shape
            End If
        End If
    Next Shape
    Set FindShapesByOutlineColor = Result
End Property

'находит поверклипы, без рекурсии
Public Property Get FindPowerClips(ByVal Shapes As ShapeRange) As ShapeRange
    Set FindPowerClips = CreateShapeRange
    Dim Shape As Shape
    For Each Shape In Shapes
        If IsValidShape(Shape) Then _
            If Not Shape.PowerClip Is Nothing Then FindPowerClips.Add Shape
    Next Shape
End Property

'находит содержимое поверклипов, без рекурсии
Public Property Get FindShapesInPowerClips( _
                        ByVal Shapes As ShapeRange _
                    ) As ShapeRange
    Dim Shape As Shape
    Set FindShapesInPowerClips = CreateShapeRange
    For Each Shape In FindPowerClips(Shapes)
        FindShapesInPowerClips.AddRange Shape.PowerClip.Shapes.All
    Next Shape
End Property

'отсюда: https://community.coreldraw.com/talk/coreldraw_graphics_suite_x4/f/coreldraw-graphics-suite-x4/57576/macro-list-fonts-within-a-text-file
Public Sub FindFontsInRange( _
                ByVal TextRange As TextRange, _
                ByVal MutFonts As Collection _
            )
    Dim FontName As String
    Dim Before As TextRange, After As TextRange
    FontName = TextRange.Font
    If FontName = "" Then
        ' There are more than one font in the range
        ' Divide the range in two and look into each half separately
        ' to see if any of them has the same font. Repeat recursively
        Set Before = TextRange.Duplicate
        Before.End = (Before.Start + Before.End) \ 2
        Set After = TextRange.Duplicate
        After.Start = Before.End
        FindFontsInRange Before, MutFonts
        FindFontsInRange After, MutFonts
    Else
        AddFontToCollection FontName, MutFonts
    End If
End Sub
'+++
Private Sub AddFontToCollection( _
                ByVal FontName As String, _
                ByVal MutFonts As Collection _
            )
    Dim Font As Variant
    Dim Found As Boolean
    Found = False
    For Each Font In MutFonts
        If Font = FontName Then
            Found = True
            Exit For
        End If
    Next Font
    If Not Found Then MutFonts.Add FontName
End Sub

Public Property Get DiffWithinTolerance( _
                        ByVal Number1 As Variant, _
                        ByVal Number2 As Variant, _
                        ByVal Tolerance As Variant _
                    ) As Boolean
    DiffWithinTolerance = Abs(Number1 - Number2) <= Tolerance
End Property

'возвращает все шейпы на всех слоях текущей страницы, по умолчанию - без мастер-слоёв и без гайдов
Public Property Get FindShapesActivePageLayers( _
                        Optional ByVal GuidesLayers As Boolean, _
                        Optional ByVal MasterLayers As Boolean _
                    ) As ShapeRange
    Dim tLayer As Layer
    Set FindShapesActivePageLayers = CreateShapeRange
    For Each tLayer In ActivePage.Layers
        If Not (tLayer.IsGuidesLayer And (GuidesLayers = False)) Then _
            FindShapesActivePageLayers.AddRange tLayer.Shapes.All
    Next
    If MasterLayers Then
        For Each tLayer In ActiveDocument.MasterPage.Layers
            If Not (tLayer.IsGuidesLayer And (GuidesLayers = False)) Then _
                FindShapesActivePageLayers.AddRange tLayer.Shapes.All
    Next
    End If
End Property

Public Property Get FindShapesNotInside( _
                        ByVal Shapes As ShapeRange _
                    ) As ShapeRange
    Set FindShapesNotInside = CreateShapeRange
    Dim Shape As Shape
    For Each Shape In Shapes
        If Not ShapeInsideAny(Shape, Shapes) Then FindShapesNotInside.Add Shape
    Next Shape
End Property

Public Property Get FindShapesWithText( _
                        ByVal Source As ShapeRange, _
                        ByVal Text As String _
                    ) As ShapeRange
    Dim TextShapes As ShapeRange
    Set TextShapes = Source.Shapes.FindShapes(Type:=cdrTextShape)
    Set FindShapesWithText = CreateShapeRange
    Dim Shape As Shape
    For Each Shape In TextShapes
        If VBA.InStr(1, Shape.Text.Story.Text, Text, vbTextCompare) > 0 Then _
            FindShapesWithText.Add Shape
    Next Shape
End Property

'возвращает коллекцию слоёв с текущей страницы, имена которых включают NamePart
Public Property Get FindLayersActivePageByNamePart( _
                        ByVal NamePart As String, _
                        Optional ByVal SearchMasters = True _
                    ) As Collection
    Dim tLayer As Layer
    Dim tLayers As Layers
    If SearchMasters Then
        Set tLayers = ActivePage.AllLayers
    Else
        Set tLayers = ActivePage.Layers
    End If
    Set FindLayersActivePageByNamePart = New Collection
    For Each tLayer In tLayers
        If InStr(tLayer.Name, NamePart) > 0 Then _
            FindLayersActivePageByNamePart.Add tLayer
    Next
End Property

'найти дубликат слоя по ряду параметров (достовернее, чем поиск по имени)
Public Property Get FindLayerDuplicate( _
                        ByVal PageToSearch As Page, _
                        ByVal SrcLayer As Layer _
                    ) As Layer
    For Each FindLayerDuplicate In PageToSearch.AllLayers
        With FindLayerDuplicate
            If (.Name = SrcLayer.Name) And _
                 (.IsDesktopLayer = SrcLayer.IsDesktopLayer) And _
                 (.Master = SrcLayer.Master) And _
                 (.Color.IsSame(SrcLayer.Color)) Then _
                 Exit Property
        End With
    Next
    Set FindLayerDuplicate = Nothing
End Property

Public Property Get GetAverageColor(ByVal Colors As Collection) As Color
    If Colors.Count = 0 Then
        Throw "No colors in colors collection"
        Exit Property
    End If
    If Colors.Count = 1 Then
        Set GetAverageColor = Colors(1).GetCopy
        Exit Property
    End If
    Dim Index As Long
    For Index = 1 To Colors.Count
        Set GetAverageColor = _
            GetMixedColor( _
                GetAverageColor, _
                Colors(Index), _
                100 - (100 / Index) _
            )
    Next Index
End Property

Public Property Get GetAverageColorFromShapes( _
                        ByVal Shapes As ShapeRange, _
                        Optional ByVal Fills As Boolean = True, _
                        Optional ByVal Outlines As Boolean = True _
                    ) As Color
    On Error GoTo NoColor
    Set GetAverageColorFromShapes = GetAverageColor( _
        GetBoundColors( _
            Shapes:=Shapes, _
            Fills:=Fills, _
            Outlines:=Outlines _
        ) _
    )
NoColor:
End Property

Public Property Get GetBoundColors( _
                        ByVal Shapes As ShapeRange, _
                        Optional ByVal Fills As Boolean = True, _
                        Optional ByVal Outlines As Boolean = True _
                    ) As Collection
    Set GetBoundColors = New Collection
    Dim Shape As Shape
    For Each Shape In Shapes
        Shape.CreateSelection
        If Fills Then _
            AppendCollection GetBoundColors, GetBoundColorsFromFill(Shape)
        If Outlines Then _
            If ShapeHasOutline(Shape) Then _
                GetBoundColors.Add Shape.Outline.Color
    Next Shape
End Property

Public Property Get GetBoundColorsFromFill( _
                        ByVal Shape As Shape _
                    ) As Collection
    Set GetBoundColorsFromFill = New Collection
    With Shape.Fill
        If Shape.Fill.Type = cdrUniformFill Then
            GetBoundColorsFromFill.Add Shape.Fill.UniformColor
        ElseIf Shape.Fill.Type = cdrFountainFill Then
            AppendCollection GetBoundColorsFromFill, _
                             GetBoundColorsFromFountain(Shape)
        ElseIf Shape.Fill.Type = cdrPatternFill Then
            If Shape.Fill.Pattern.Type = cdrTwoColorPattern Then
                AppendCollection GetBoundColorsFromFill, _
                                 GetBoundColorsFromTwoColorPattern(Shape)
            End If
        End If
    End With
End Property

Public Property Get GetBoundColorsFromFountain( _
                        ByVal Shape As Shape _
                    ) As Collection
    Set GetBoundColorsFromFountain = New Collection
    Dim FColor As FountainColor
    For Each FColor In Shape.Fill.Fountain.Colors
        GetBoundColorsFromFountain.Add FColor.Color
    Next FColor
End Property

Public Property Get GetBoundColorsFromTwoColorPattern( _
                        ByVal Shape As Shape _
                    ) As Collection
    Set GetBoundColorsFromTwoColorPattern = New Collection
    GetBoundColorsFromTwoColorPattern.Add Shape.Fill.Pattern.FrontColor
    GetBoundColorsFromTwoColorPattern.Add Shape.Fill.Pattern.BackColor
End Property

Public Property Get GetBottomOrderShape(ByVal Shapes As ShapeRange) As Shape
    If Shapes.Count = 0 Then Exit Property
    Set GetBottomOrderShape = Shapes(1)
    If Shapes.Count = 1 Then Exit Property
    Dim Index As Long
    For Index = 2 To Shapes.Count
        If Shapes(Index).ZOrder > GetBottomOrderShape.ZOrder Then
            Set GetBottomOrderShape = Shapes(Index)
        End If
    Next Index
End Property

Public Property Get GetColorLightness(ByVal Color As Color) As Long
    Dim GrayScale As Color
    Set GrayScale = Color.GetCopy
    GrayScale.ConvertToGray
    GetColorLightness = GrayScale.Gray
End Property

Public Property Get GetHeightKeepProportions( _
                        ByVal Rect As Rect, _
                        ByVal Width As Double _
                    ) As Double
    Dim WidthToHeight As Double
    WidthToHeight = Rect.Width / Rect.Height
    GetHeightKeepProportions = Width / WidthToHeight
End Property

Public Property Get GetMixedColor( _
                        ByVal MaybeColor1 As Variant, _
                        ByVal MaybeColor2 As Variant, _
                        Optional ByVal MixRatio As Long = 50 _
                    ) As Color
    If Not (IsColor(MaybeColor1) Or IsColor(MaybeColor2)) Then Exit Property
    If Not IsColor(MaybeColor1) Then
        Set GetMixedColor = MaybeColor2.GetCopy
        Exit Property
    ElseIf Not IsColor(MaybeColor2) Then
        Set GetMixedColor = MaybeColor1.GetCopy
        Exit Property
    End If
    Set GetMixedColor = MaybeColor1.GetCopy
    GetMixedColor.BlendWith MaybeColor2, MixRatio
End Property

Public Property Get GetRotatedRect(ByVal Rect As Rect) As Rect
    Dim HalfDifference As Double
    HalfDifference = (Rect.Width - Rect.Height) / 2
    Set GetRotatedRect = Rect.GetCopy
    GetRotatedRect.Inflate _
        -HalfDifference, HalfDifference, -HalfDifference, HalfDifference
End Property

Public Property Get GetTopOrderShape(ByVal Shapes As ShapeRange) As Shape
    If Shapes.Count = 0 Then Exit Property
    Set GetTopOrderShape = Shapes(1)
    If Shapes.Count = 1 Then Exit Property
    Dim Index As Long
    For Index = 2 To Shapes.Count
        If Shapes(Index).ZOrder < GetTopOrderShape.ZOrder Then
            Set GetTopOrderShape = Shapes(Index)
        End If
    Next Index
End Property

Public Property Get GetWidthKeepProportions( _
                        ByVal Rect As Rect, _
                        ByVal Height As Double _
                    ) As Double
    Dim WidthToHeight As Double
    WidthToHeight = Rect.Width / Rect.Height
    GetWidthKeepProportions = Height * WidthToHeight
End Property

'возвращает бОльшую сторону шейпа/рэйнджа/страницы
Public Property Get GreaterDim(ByVal ShapeOrRangeOrPage As Object) As Double
    If Not TypeOf ShapeOrRangeOrPage Is Shape _
   And Not TypeOf ShapeOrRangeOrPage Is ShapeRange _
   And Not TypeOf ShapeOrRangeOrPage Is Page Then
        Err.Raise 13, Source:="GreaterDim", _
                  Description:="Type mismatch: ShapeOrRangeOrPage должен быть Shape, ShapeRange или Page"
        Exit Property
    End If
    If ShapeOrRangeOrPage.SizeWidth > ShapeOrRangeOrPage.SizeHeight Then
        GreaterDim = ShapeOrRangeOrPage.SizeWidth
    Else
        GreaterDim = ShapeOrRangeOrPage.SizeHeight
    End If
End Property

Public Property Get HasSize(ByRef MaybeSome As Variant) As Boolean
    If Not ObjectAssigned(MaybeSome) Then Exit Property
    If TypeOf MaybeSome Is Shape Then GoTo Success
    If TypeOf MaybeSome Is ShapeRange Then GoTo Success
    If TypeOf MaybeSome Is Page Then GoTo Success
    If TypeOf MaybeSome Is Rect Then GoTo Success
    Exit Property
Success:
    HasSize = True
End Property

Public Property Get HasPosition(ByRef MaybeSome As Variant) As Boolean
    If Not ObjectAssigned(MaybeSome) Then Exit Property
    If TypeOf MaybeSome Is Shape Then GoTo Success
    If TypeOf MaybeSome Is ShapeRange Then GoTo Success
    If TypeOf MaybeSome Is Page Then GoTo Success
    If TypeOf MaybeSome Is Rect Then GoTo Success
    If TypeOf MaybeSome Is Node Then GoTo Success
    If TypeOf MaybeSome Is SubPath Then GoTo Success
    Exit Property
Success:
    HasPosition = True
End Property

Public Property Get IsColor(ByRef MaybeColor As Variant) As Boolean
    If Not ObjectAssigned(MaybeColor) Then Exit Property
    IsColor = TypeOf MaybeColor Is Color
End Property

Public Property Get IsCurve(ByRef MaybeCurve As Variant) As Boolean
    If Not ObjectAssigned(MaybeCurve) Then Exit Property
    IsCurve = TypeOf MaybeCurve Is Curve
End Property

Public Property Get IsDark(ByVal Color As Color) As Boolean
    Dim HSB As Color: Set HSB = Color.GetCopy
    HSB.ConvertToHSB
    IsDark = HSB.HSBBrightness < 128
End Property

Public Property Get IsDocument(ByRef MaybeDocument As Variant) As Boolean
    If Not ObjectAssigned(MaybeDocument) Then Exit Property
    IsDocument = TypeOf MaybeDocument Is Document
End Property

'True, если Value - значение или присвоенный объект (не пустота, не ошибка...)
Public Property Get IsJust(ByRef Value As Variant) As Boolean
    IsJust = Not (VBA.IsError(Value) Or IsNone(Value))
End Property

'является ли шейп/рэйндж/страница альбомным
Public Property Get IsLandscape(ByVal ShapeOrRangeOrPage As Object) As Boolean
    If Not TypeOf ShapeOrRangeOrPage Is Shape _
   And Not TypeOf ShapeOrRangeOrPage Is ShapeRange _
   And Not TypeOf ShapeOrRangeOrPage Is Page Then
        Err.Raise 13, Source:="IsLandscape", _
                  Description:="Type mismatch: ShapeOrRangeOrPage должен быть Shape, ShapeRange или Page"
        Exit Property
    End If
    If ShapeOrRangeOrPage.SizeWidth > ShapeOrRangeOrPage.SizeHeight Then
        IsLandscape = True
    Else
        IsLandscape = False
    End If
End Property

Public Property Get IsLayer(ByRef MaybeLayer As Variant) As Boolean
    If Not ObjectAssigned(MaybeLayer) Then Exit Property
    IsLayer = TypeOf MaybeLayer Is Layer
End Property

'todo: ПРОВЕРИТЬ КАК СЛЕДУЕТ
Public Property Get IsOverlap( _
                        ByVal FirstShape As Shape, _
                        ByVal SecondShape As Shape _
                    ) As Boolean
    
    Dim tIS As Shape
    Dim tShape1 As Shape, tShape2 As Shape
    Dim tBound1 As Shape, tBound2 As Shape
    Dim tProps As LayerProperties
    
    If FirstShape.Type = cdrConnectorShape _
    Or SecondShape.Type = cdrConnectorShape Then _
        Exit Property
    
    'запоминаем какой слой был активным
    Dim tLayer As Layer: Set tLayer = ActiveLayer
    'запоминаем состояние первого слоя
    FirstShape.Layer.Activate
    LayerPropsPreserveAndReset FirstShape.Layer, tProps
    
    If IsIntersectReady(FirstShape) Then
        Set tShape1 = FirstShape
    Else
        Set tShape1 = CreateBoundary(FirstShape)
        Set tBound1 = tShape1
    End If
    
    If IsIntersectReady(SecondShape) Then
        Set tShape2 = SecondShape
    Else
        Set tShape2 = CreateBoundary(SecondShape)
        Set tBound2 = tShape2
    End If
    
    Set tIS = tShape1.Intersect(tShape2)
    If tIS Is Nothing Then
        IsOverlap = False
    Else
        tIS.Delete
        IsOverlap = True
    End If
    
    On Error Resume Next
        tBound1.Delete
        tBound2.Delete
    On Error GoTo 0
    
    'возвращаем всё на место
    LayerPropsRestore FirstShape.Layer, tProps
    tLayer.Activate

End Property

'IsOverlap здорового человека - меряет по габаритам,
'но зато стабильно работает и в большинстве случаев его достаточно
Public Property Get IsOverlapBox( _
                        ByVal FirstShape As Shape, _
                        ByVal SecondShape As Shape _
                    ) As Boolean
    Dim tShape As Shape
    Dim tProps As LayerProperties
    'запоминаем какой слой был активным
    Dim tLayer As Layer: Set tLayer = ActiveLayer
    'запоминаем состояние первого слоя
    FirstShape.Layer.Activate
    LayerPropsPreserveAndReset FirstShape.Layer, tProps
    Dim tRect As Rect
    Set tRect = FirstShape.BoundingBox.Intersect(SecondShape.BoundingBox)
    If tRect.Width = 0 And tRect.Height = 0 Then
        IsOverlapBox = False
    Else
        IsOverlapBox = True
    End If
    'возвращаем всё на место
    LayerPropsRestore FirstShape.Layer, tProps
    tLayer.Activate
End Property

Public Property Get IsPage(ByRef MaybePage As Variant) As Boolean
    If Not ObjectAssigned(MaybePage) Then Exit Property
    IsPage = TypeOf MaybePage Is Page
End Property

Public Property Get IsRect(ByRef MaybeRect As Variant) As Boolean
    If Not ObjectAssigned(MaybeRect) Then Exit Property
    IsRect = TypeOf MaybeRect Is Rect
End Property

Public Property Get IsSameColor( _
                        ByVal MaybeColor1 As Variant, _
                        ByVal MaybeColor2 As Variant _
                    ) As Boolean
    If VBA.IsEmpty(MaybeColor1) Or VBA.IsEmpty(MaybeColor2) Then Exit Property
    IsSameColor = MaybeColor1.IsSame(MaybeColor2)
End Property

'являются ли кривые дубликатами, находящимися друг над другом в одном месте
'(underlying dubs)
Public Property Get IsSameCurves( _
                        ByVal Curve1 As Curve, _
                        ByVal Curve2 As Curve _
                    ) As Boolean
    Dim tNode As Node
    Dim Tolerance As Double
    'допуск = 0.001 мм
    Tolerance = ConvertUnits(0.001, cdrMillimeter, ActiveDocument.Unit)
    IsSameCurves = False
    If Not Curve1.Nodes.Count = Curve2.Nodes.Count Then Exit Property
    If Abs(Curve1.Length - Curve2.Length) > Tolerance Then Exit Property
    For Each tNode In Curve1.Nodes
        If Curve2.FindNodeAtPoint( _
               tNode.PositionX, _
               tNode.PositionY, _
               Tolerance * 2 _
           ) Is Nothing Then Exit Property
    Next
    IsSameCurves = True
End Property

Public Property Get IsNode(ByRef MaybeNode As Variant) As Boolean
    If Not ObjectAssigned(MaybeNode) Then Exit Property
    IsNode = TypeOf MaybeNode Is Node
End Property

Public Property Get IsSegment(ByRef MaybeSegment As Variant) As Boolean
    If Not ObjectAssigned(MaybeSegment) Then Exit Property
    IsSegment = TypeOf MaybeSegment Is Segment
End Property

Public Property Get IsShape(ByRef MaybeShape As Variant) As Boolean
    If Not ObjectAssigned(MaybeShape) Then Exit Property
    IsShape = TypeOf MaybeShape Is Shape
End Property

Public Property Get IsShapeRange(ByRef MaybeShapeRange As Variant) As Boolean
    If Not ObjectAssigned(MaybeShapeRange) Then Exit Property
    IsShapeRange = TypeOf MaybeShapeRange Is ShapeRange
End Property

Public Property Get IsShapeType( _
                        ByVal MaybeShape As Variant, _
                        ByVal ShapeType As cdrShapeType _
                    ) As Boolean
    If Not IsShape(MaybeShape) Then Exit Property
    IsShapeType = (MaybeShape.Type = ShapeType)
End Property

Public Property Get IsSubPath(ByRef MaybeSubPath As Variant) As Boolean
    If Not ObjectAssigned(MaybeSubPath) Then Exit Property
    IsSubPath = TypeOf MaybeSubPath Is SubPath
End Property

Public Property Get IsValidCurve(ByVal MaybeCurve As Variant) As Boolean
    If Not IsCurve(MaybeCurve) Then GoTo Fail
    Dim Temp As Double
    On Error GoTo Fail
    Temp = MaybeCurve.Length
    On Error GoTo 0
    IsValidCurve = True
Fail:
End Property

Public Property Get IsValidDocument(ByRef MaybeDocument As Variant) As Boolean
    If Not IsDocument(MaybeDocument) Then GoTo Fail
    Dim Temp As String
    On Error GoTo Fail
    Temp = MaybeDocument.Name
    On Error GoTo 0
    IsValidDocument = True
Fail:
End Property

Public Property Get IsValidLayer(ByVal MaybeLayer As Variant) As Boolean
    If Not IsLayer(MaybeLayer) Then GoTo Fail
    Dim Temp As String
    On Error GoTo Fail
    Temp = MaybeLayer.Name
    On Error GoTo 0
    IsValidLayer = True
Fail:
End Property

Public Property Get IsValidPage(ByVal MaybePage As Variant) As Boolean
    If Not IsPage(MaybePage) Then GoTo Fail
    Dim Temp As String
    On Error GoTo Fail
    Temp = MaybePage.Name
    On Error GoTo 0
    IsValidPage = True
Fail:
End Property

Public Property Get IsValidSegment(ByVal MaybeSegment As Variant) As Boolean
    If Not IsSegment(MaybeSegment) Then GoTo Fail
    Dim Temp As Long
    On Error GoTo Fail
    Temp = MaybeSegment.AbsoluteIndex
    On Error GoTo 0
    IsValidSegment = True
Fail:
End Property

Public Property Get IsValidSubPath(ByVal MaybeSubPath As Variant) As Boolean
    If Not IsSubPath(MaybeSubPath) Then GoTo Fail
    Dim Temp As Boolean
    On Error GoTo Fail
    Temp = MaybeSubPath.Closed
    On Error GoTo 0
    IsValidSubPath = True
Fail:
End Property

Public Property Get GetCombinedCurve(ByVal Shapes As ShapeRange) As Curve
    Set GetCombinedCurve = CreateCurve(Shapes.FirstShape.Page.Parent.Parent)
    Dim Shape As Shape
    For Each Shape In Shapes
        GetCombinedCurve.AppendCurve GetCurve(Shape)
    Next Shape
End Property

Public Property Get GetCurve(ByVal MaybeShape As Variant) As Curve
    If Not IsShape(MaybeShape) Then GoTo Fail
    Dim Temp As Double
    On Error GoTo Fail
    Set GetCurve = MaybeShape.Curve
    On Error GoTo 0
Fail:
End Property

Public Property Get HasDiplayCurve(ByVal MaybeShape As Variant) As Boolean
    If Not IsShape(MaybeShape) Then GoTo Fail
    Dim Temp As Double
    On Error GoTo Fail
    Temp = MaybeShape.DisplayCurve.Length
    On Error GoTo 0
    HasDiplayCurve = Temp > 0
Fail:
End Property

Public Property Get HasCurve(ByVal MaybeShape As Variant) As Boolean
    If Not IsShape(MaybeShape) Then GoTo Fail
    Dim Temp As Double
    On Error GoTo Fail
    Temp = MaybeShape.Curve.Length
    On Error GoTo 0
    HasCurve = Temp > 0
Fail:
End Property

Public Property Get IsValidNode(ByVal MaybeNode As Variant) As Boolean
    If Not IsNode(MaybeNode) Then GoTo Fail
    Dim Temp As Long
    On Error GoTo Fail
    Temp = MaybeNode.AbsoluteIndex
    On Error GoTo 0
    IsValidNode = True
Fail:
End Property

Public Property Get IsValidShape(ByVal MaybeShape As Variant) As Boolean
    If Not IsShape(MaybeShape) Then GoTo Fail
    Dim Temp As String
    On Error GoTo Fail
    Temp = MaybeShape.Name
    On Error GoTo 0
    IsValidShape = True
Fail:
End Property

Public Property Get IsValidShapeRange(ByVal MaybeShapeRange As Variant) As Boolean
    If Not IsShapeRange(MaybeShapeRange) Then GoTo Fail
    Dim Temp As Long
    On Error GoTo Fail
    Temp = MaybeShapeRange.Count
    On Error GoTo 0
    IsValidShapeRange = True
Fail:
End Property

'возвращает меньшую сторону шейпа/рэйнджа/страницы
Public Property Get LesserDim(ByVal ShapeOrRangeOrPage As Object) As Double
    If Not TypeOf ShapeOrRangeOrPage Is Shape _
   And Not TypeOf ShapeOrRangeOrPage Is ShapeRange _
   And Not TypeOf ShapeOrRangeOrPage Is Page Then
        Err.Raise 13, Source:="LesserDim", _
                  Description:="Type mismatch: ShapeOrRangeOrPage должен быть Shape, ShapeRange или Page"
        Exit Property
    End If
    If ShapeOrRangeOrPage.SizeWidth < ShapeOrRangeOrPage.SizeHeight Then
        LesserDim = ShapeOrRangeOrPage.SizeWidth
    Else
        LesserDim = ShapeOrRangeOrPage.SizeHeight
    End If
End Property

'количество объектов BoxToFit, которое поместится как есть на площади Area
Public Property Get NumberToFitArea( _
                        ByVal BoxToFit As Rect, _
                        ByVal Area As Rect _
                    ) As Long
    NumberToFitArea = Fix(Area.Width / BoxToFit.Width) _
                    * Fix(Area.Height / BoxToFit.Height)
End Property

Public Property Get PixelsToDocUnits(ByVal SizeInPixels As Long) As Double
    PixelsToDocUnits = ConvertUnits(SizeInPixels, cdrPixel, ActiveDocument.Unit)
End Property

Public Property Get RectInsideRect( _
                        ByVal Rect1 As Rect, _
                        ByVal Rect2 As Rect _
                    ) As Boolean
    RectInsideRect = _
        (Rect1.Left > Rect2.Left) _
    And (Rect1.Right < Rect2.Right) _
    And (Rect1.Top < Rect2.Top) _
    And (Rect1.Bottom > Rect2.Bottom)
End Property

Public Property Get ShapeHasOutline(ByVal Shape As Shape) As Boolean
    On Error GoTo Fail
    ShapeHasOutline = Not (Shape.Outline.Type = cdrNoOutline)
Fail:
End Property

Public Property Get ShapeHasUniformFill(ByVal Shape As Shape) As Boolean
    On Error GoTo Fail
    ShapeHasUniformFill = (Shape.Fill.Type = cdrUniformFill)
Fail:
End Property

Public Property Get ShapeInsideAny( _
                        ByVal Shape As Shape, _
                        ByVal Shapes As ShapeRange _
                    ) As Boolean
    Dim Curve As Curve
    Dim CurrentShape As Shape
    For Each CurrentShape In Shapes
        If Not CurrentShape Is Shape Then
            If ShapeInsideShape(Shape, CurrentShape) Then
                ShapeInsideAny = True
                Exit Property
            End If
        End If
    Next CurrentShape
End Property

Public Property Get ShapeInsideShape( _
                        ByVal Shape1 As Shape, _
                        ByVal Shape2 As Shape _
                    ) As Boolean
    Dim Curve1 As Curve
    If HasCurve(Shape1) Then
        Set Curve1 = Shape1.Curve
    ElseIf HasDiplayCurve(Shape1) Then
        Set Curve1 = Shape1.DisplayCurve
    End If
    
    Dim Curve2 As Curve
    If HasCurve(Shape2) Then
        Set Curve2 = Shape2.Curve
    ElseIf HasDiplayCurve(Shape2) Then
        Set Curve2 = Shape2.DisplayCurve
    End If
    
    If IsNone(Curve1) Or IsNone(Curve2) Then
        ShapeInsideShape = _
            RectInsideRect(Shape1.BoundingBox, Shape2.BoundingBox)
    Else
        ShapeInsideShape = AllNodesInside(Curve1.Nodes.All, Curve2)
    End If
End Property

Public Property Get ShapeIsInGroup(ByVal Shape As Shape) As Boolean
    On Error GoTo Fail
    ShapeIsInGroup = Not (Shape.ParentGroup Is Nothing)
Fail:
End Property

'возвращает коллекцию слоёв, на которых лежат шейпы из ренджа
Public Property Get ShapeRangeLayers( _
                        ByVal ShapeRange As ShapeRange _
                    ) As Collection
    
    Dim tShape As Shape
    Dim tLayer As Layer
    Dim inCol As Boolean
    
    If ShapeRange.Count = 0 Then Exit Property
    Set ShapeRangeLayers = New Collection
    If ShapeRange.Count = 1 Then
        ShapeRangeLayers.Add ShapeRange(1).Layer
        Exit Property
    End If
    
    For Each tShape In ShapeRange
        inCol = False
        For Each tLayer In ShapeRangeLayers
            If tLayer Is tShape.Layer Then
                inCol = True
                Exit For
            End If
        Next tLayer
        If inCol = False Then ShapeRangeLayers.Add tShape.Layer
    Next tShape

End Property

'возвращает Rect, равный габаритам объекта плюс Space со всех сторон
Public Property Get SpaceBox( _
                        ByVal MaybeHasSize As Variant, _
                        ByVal Space As Double _
                    ) As Rect
    If Not ObjectAssigned(MaybeHasSize) Then Exit Property
    If TypeOf MaybeHasSize Is Shape Then
        Set SpaceBox = MaybeHasSize.BoundingBox.GetCopy
    ElseIf TypeOf MaybeHasSize Is ShapeRange Then
        Set SpaceBox = MaybeHasSize.BoundingBox.GetCopy
    ElseIf TypeOf MaybeHasSize Is Page Then
        Set SpaceBox = MaybeHasSize.BoundingBox.GetCopy
    ElseIf TypeOf MaybeHasSize Is Rect Then
        Set SpaceBox = MaybeHasSize.GetCopy
    Else
        Exit Property
    End If
    SpaceBox.Inflate Space, Space, Space, Space
End Property

'возвращает Outline если или Empty
Public Property Get TryGetOutline(ByVal Shape As Shape) As Variant
    If ShapeHasOutline(Shape) Then Set TryGetOutline = Shape.Outline
End Property

'===============================================================================
' # функции манипуляций с объектами корела

Public Function AddPage( _
                    ByVal MaybeAfterPageOrIndex As Variant _
                ) As Page
    Dim Index As Long
    If IsPage(MaybeAfterPageOrIndex) Then
        Index = MaybeAfterPageOrIndex.Index
        MaybeAfterPageOrIndex.Parent.Parent.Activate
    End If
    If VBA.IsNumeric(MaybeAfterPageOrIndex) Then
        Index = MaybeAfterPageOrIndex
    End If
    If Index < 1 Then Exit Function
    Set AddPage = ActiveDocument.AddPages(1)
    AddPage.MoveTo Index + 1
End Function

Public Function AddPages( _
                    ByVal MaybeAfterPageOrIndex As Variant, _
                    ByVal Quantity As Long _
                ) As Collection
    Set AddPages = New Collection
    If Quantity < 1 Then Exit Function
    Dim LastPage As Page
    Set LastPage = AddPage(MaybeAfterPageOrIndex)
    AddPages.Add LastPage
    If Quantity = 1 Then Exit Function
    Dim Index As Long
    For Index = 2 To Quantity
        Set LastPage = AddPage(LastPage)
        AddPages.Add LastPage
    Next Index
End Function

Public Sub Align( _
               ByVal ShapesBeingAligned As ShapeRange, _
               ByVal RelativeTo As Rect, _
               ByVal ReferencePoint As cdrReferencePoint _
           )
    With ShapesBeingAligned
        Select Case ReferencePoint
            Case cdrTopRight
                .TopY = RelativeTo.Top
                .RightX = RelativeTo.Right
            Case cdrTopMiddle
                .TopY = RelativeTo.Top
                .CenterX = RelativeTo.CenterX
            Case cdrTopLeft
                .TopY = RelativeTo.Top
                .LeftX = RelativeTo.Left
            Case cdrMiddleLeft
                .CenterY = RelativeTo.CenterY
                .LeftX = RelativeTo.Left
            Case cdrBottomLeft
                .BottomY = RelativeTo.Bottom
                .LeftX = RelativeTo.Left
            Case cdrBottomMiddle
                .BottomY = RelativeTo.Bottom
                .CenterX = RelativeTo.CenterX
            Case cdrBottomRight
                .BottomY = RelativeTo.Bottom
                .RightX = RelativeTo.Right
            Case cdrMiddleRight
                .CenterY = RelativeTo.CenterY
                .RightX = RelativeTo.Right
            Case cdrCenter
                .CenterY = RelativeTo.CenterY
                .CenterX = RelativeTo.CenterX
        End Select
    End With
End Sub

Public Function BreakApart(ByVal Shape As Shape) As ShapeRange
    If Shape.Curve.SubPaths.Count < 2 Then
        Set BreakApart = CreateShapeRange
        BreakApart.Add Shape
        Exit Function
    End If
    Set BreakApart = Shape.BreakApartEx
    If BreakApart.Count > 1 Then Exit Function
    Set BreakApart = CreateShapeRange
    Dim RemainingShape As Shape
    Dim ExtractedShape As Shape
    'RemainingShape и ExtractedShape в Extract
    'на самом деле наоборот, чем в спеках
    Set RemainingShape = Shape.Curve.SubPaths.First.Extract(ExtractedShape)
    BreakApart.Add ExtractedShape
    BreakApart.AddRange BreakApart(RemainingShape)
End Function

'разбивает шейп по узлам, т. е. получаются шейпы с двумя узлами
Public Function BreakByNodes(ByVal Shape As Shape) As ShapeRange
    Set BreakByNodes = CreateShapeRange
    If Shape.Curve.Nodes.Count <= 2 Then
        BreakByNodes.Add Shape
        Exit Function
    End If
    Dim Node As Node
    For Each Node In Shape.Curve.Nodes.All
        Node.BreakApart
    Next Node
    Dim TempShape As Shape
    For Each TempShape In BreakApart(Shape)
        BreakByNodes.AddRange BreakByNodes(TempShape)
    Next TempShape
End Function

'перекрашивает объект в чёрный или белый в серой шкале,
'в зависимости от исходного цвета
'ДОРАБОТАТЬ
Public Function ContrastShape(ByVal Shape As Shape) As Shape
    With Shape.Fill
        Select Case .Type
            Case cdrUniformFill
                .UniformColor.ConvertToGray
                If .UniformColor.Gray < 128 Then
                    .UniformColor.GrayAssign 0
                Else
                    .UniformColor.GrayAssign 255
                End If
            Case cdrFountainFill
                'todo
        End Select
    End With
    With Shape.Outline
        If Not .Type = cdrNoOutline Then
            .Color.ConvertToGray
            If .Color.Gray < 128 Then _
                .Color.GrayAssign 0 Else .Color.GrayAssign 255
        End If
    End With
    Set ContrastShape = Shape
End Function

'правильно копирует Shape или ShapeRange на другой слой
Public Function CopyToLayer( _
                    ByVal ShapeOrRange As Object, _
                    ByVal Layer As Layer _
                ) As Object

    If Not TypeOf ShapeOrRange Is Shape And Not TypeOf ShapeOrRange Is ShapeRange Then
        Err.Raise 13, Source:="CopyToLayer", _
                  Description:="Type mismatch: ShapeOrRange должен быть Shape или ShapeRange"
        Exit Function
    End If
    
    Set CopyToLayer = ShapeOrRange.Duplicate
    MoveToLayer CopyToLayer, Layer

End Function

'инструмент Boundary
Public Function CreateBoundary(ByVal ShapeOrRange As Object) As Shape
    On Error GoTo Catch
    Dim tShape As Shape, tRange As ShapeRange
    'просто объект не ест, надо конкретный тип
    If TypeOf ShapeOrRange Is Shape Then
        Set tShape = ShapeOrRange
        Set CreateBoundary = tShape.CustomCommand("Boundary", "CreateBoundary")
    ElseIf TypeOf ShapeOrRange Is ShapeRange Then
        Set tRange = ShapeOrRange
        Set CreateBoundary = tRange.CustomCommand("Boundary", "CreateBoundary")
    Else
        Err.Raise 13, Source:="CreateBoundary", _
            Description:="Type mismatch: ShapeOrRange должен быть Shape или ShapeRange"
        Exit Function
    End If
    Exit Function
Catch:
    Debug.Print Err.Number
End Function

'создаёт слой, если такой слой есть - возвращает этот слой
Public Function CreateOrFindLayer( _
                    ByVal Page As Page, _
                    ByVal Name As String _
                ) As Layer
    Set CreateOrFindLayer = Page.Layers.Find(Name)
    If CreateOrFindLayer Is Nothing Then
        Set CreateOrFindLayer = Page.CreateLayer(Name)
    End If
End Function

'инструмент Crop Tool
Public Function CropTool( _
                    ByVal ShapeOrRangeOrPage As Object, _
                    ByVal x1 As Double, ByVal y1 As Double, _
                    ByVal x2 As Double, ByVal y2 As Double, _
                    Optional ByVal Angle = 0 _
                ) As ShapeRange
    If TypeOf ShapeOrRangeOrPage Is Shape Or _
         TypeOf ShapeOrRangeOrPage Is ShapeRange Or _
         TypeOf ShapeOrRangeOrPage Is Page Then
        Set CropTool = ShapeOrRangeOrPage.CustomCommand("Crop", "CropRectArea", x1, y1, x2, y2, Angle)
    Else
        Err.Raise 13, Source:="CropTool", _
            Description:="Type mismatch: ShapeOrRangeOrPage должен быть Shape, ShapeRange или Page"
        Exit Function
    End If
End Function

'отрезать кусок от Shape по контуру Knife, возвращает отрезанный кусок
Public Function Dissect(ByRef Shape As Shape, ByRef Knife As Shape) As Shape
    Set Dissect = Intersect(Knife, Shape, True, True)
    Set Shape = Knife.Trim(Shape, True, False)
End Function

'дублировать активную страницу со всеми слоями и объектами
Public Function DuplicateActivePage( _
                    ByVal NumberOfPages As Long, _
                    Optional ByVal ExcludeLayerName As String = "" _
                ) As Page
    Dim tRange As ShapeRange
    Dim tShape As Shape, sDuplicate As Shape
    Dim tProps As LayerProperties
    Dim i&
    For i = 1 To NumberOfPages
        Set tRange = FindShapesActivePageLayers
        Set DuplicateActivePage = _
            ActiveDocument.InsertPages(1, False, ActivePage.Index)
        DuplicateActivePage.SizeHeight = ActivePage.SizeHeight
        DuplicateActivePage.SizeWidth = ActivePage.SizeWidth
        For Each tShape In tRange.ReverseRange
            If tShape.Layer.Name <> ExcludeLayerName Then
                LayerPropsPreserveAndReset tShape.Layer, tProps
                Set sDuplicate = tShape.Duplicate
                sDuplicate.MoveToLayer _
                    FindLayerDuplicate(DuplicateActivePage, tShape.Layer)
                LayerPropsRestore tShape.Layer, tProps
            End If
        Next tShape
    Next i
End Function

Public Sub FillInside( _
               ByVal MaybeShapeOrRange As Variant, _
               ByVal MaybeTargetRect As Variant _
           )
    If VBA.IsEmpty(MaybeShapeOrRange) _
    Or VBA.IsEmpty(MaybeTargetRect) Then Exit Sub
    ThrowIfNotShapeOrRange MaybeShapeOrRange
    
    If GetHeightKeepProportions( _
           MaybeShapeOrRange.BoundingBox, _
           MaybeTargetRect.Width _
       ) > MaybeTargetRect.Height Then
        MaybeShapeOrRange.SetSize MaybeTargetRect.Width
    Else
        MaybeShapeOrRange.SetSize , MaybeTargetRect.Height
    End If
    MaybeShapeOrRange.CenterX = MaybeTargetRect.CenterX
    MaybeShapeOrRange.CenterY = MaybeTargetRect.CenterY
End Sub

Public Sub FitInside( _
               ByVal ShapeToFit As Shape, _
               ByVal TargetRect As Rect _
           )
    If GetHeightKeepProportions(ShapeToFit.BoundingBox, TargetRect.Width) _
     > TargetRect.Height Then
        ShapeToFit.SetSize , TargetRect.Height
    Else
        ShapeToFit.SetSize TargetRect.Width
    End If
    ShapeToFit.CenterX = TargetRect.CenterX
    ShapeToFit.CenterY = TargetRect.CenterY
End Sub

'все объекты на всех страницах, включая мастер-страницу - на один слой
'все страницы прибиваются, все объекты на слоях guides прибиваются
Public Function FlattenPagesToLayer(ByVal LayerName As String) As Layer

    Dim DL As Layer: Set DL = ActiveDocument.MasterPage.DesktopLayer
    Dim DLstate As Boolean: DLstate = DL.Editable
    Dim p As Page
    Dim L As Layer
    
    DL.Editable = False
    
    For Each p In ActiveDocument.Pages
        For Each L In p.Layers
            If L.IsSpecialLayer Then
                L.Shapes.All.Delete
            Else
                L.Activate
                L.Editable = True
                With L.Shapes.All
                    .MoveToLayer DL
                    .OrderToBack
                End With
                L.Delete
            End If
        Next
        If p.Index <> 1 Then p.Delete
    Next
    
    Set FlattenPagesToLayer = ActiveDocument.Pages.First.CreateLayer(LayerName)
    FlattenPagesToLayer.MoveBelow ActiveDocument.Pages.First.GuidesLayer
    
    For Each L In ActiveDocument.MasterPage.Layers
        If Not L.IsSpecialLayer Or L.IsDesktopLayer Then
            L.Activate
            L.Editable = True
            With L.Shapes.All
                .MoveToLayer FlattenPagesToLayer
                .OrderToBack
            End With
            If Not L.IsSpecialLayer Then L.Delete
        Else
            L.Shapes.All.Delete
        End If
    Next
    
    FlattenPagesToLayer.Activate
    DL.Editable = DLstate

End Function

Public Function Group( _
                    ByVal Shapes As ShapeRange, Optional ByVal Name As String _
                ) As Shape
    Shapes.FirstShape.Page.Activate
    Set Group = Shapes.Group
    If Not Name = vbNullString Then Group.Name = Name
End Function

'правильный интерсект
Public Function Intersect( _
                    ByVal SourceShape As Shape, _
                    ByVal TargetShape As Shape, _
                    Optional ByVal LeaveSource As Boolean = True, _
                    Optional ByVal LeaveTarget As Boolean = True _
                ) As Shape
                                     
    Dim tPropsSource As LayerProperties
    Dim tPropsTarget As LayerProperties
    
    If Not SourceShape.Layer Is TargetShape.Layer Then _
        LayerPropsPreserveAndReset SourceShape.Layer, tPropsSource
    LayerPropsPreserveAndReset TargetShape.Layer, tPropsTarget
    
    Set Intersect = SourceShape.Intersect(TargetShape)
    
    If Not SourceShape.Layer Is TargetShape.Layer Then _
        LayerPropsRestore SourceShape.Layer, tPropsSource
    LayerPropsRestore TargetShape.Layer, tPropsTarget
    
    If Intersect Is Nothing Then Exit Function
    
    Intersect.OrderFrontOf TargetShape
    If Not LeaveSource Then SourceShape.Delete
    If Not LeaveTarget Then TargetShape.Delete

End Function

'инструмент Join Curves
Public Function JoinCurves( _
                    ByVal ShapeOrShapes As Variant, _
                    ByVal Tolerance As Double _
                )
    ShapeOrShapes.CustomCommand "ConvertTo", "JoinCurves", Tolerance
End Function

'ПРОВЕРИТЬ
Public Function MakeContour( _
                    ByRef Shape As Shape, _
                    ByVal OFFSET As Double, _
                    Optional ByVal CornerType As cdrContourCornerType _
                ) As Shape
    Dim Direction As cdrContourDirection
    If OFFSET > 0 Then
        Direction = cdrContourOutside
    ElseIf OFFSET < 0 Then
        Direction = cdrContourInside
    Else
        Exit Function
    End If
    Dim Contour As ShapeRange
    With Shape.CreateContour( _
            Direction:=Direction, _
            OFFSET:=Abs(OFFSET), _
            Steps:=1 _
        )
        .Contour.CornerType = CornerType
        Set Contour = .Separate
    End With
    Set MakeContour = Contour(1)
    Set Shape = Contour(2)
End Function

Public Function MakeCircle( _
                    ByVal x As Double, _
                    ByVal y As Double, _
                    ByVal Radius As Double, _
                    Optional ByVal FillColor As Color, _
                    Optional ByVal OutlineColor As Color _
                ) As Shape
    Set MakeCircle = ActiveLayer.CreateEllipse2(x, y, Radius)
    With MakeCircle
        If IsSome(FillColor) Then .Fill.ApplyUniformFill FillColor
        If IsSome(OutlineColor) Then .Outline.Color = OutlineColor
    End With
End Function

Public Function MakeLinearDimension( _
                    ByVal Point1 As SnapPoint, _
                    ByVal Point2 As SnapPoint, _
                    Optional ByVal TextCentered As Boolean = True, _
                    Optional ByVal TextOffset As Double, _
                    Optional ByVal DimensionStyle As Long = 0, _
                    Optional ByVal Precision As Long = 2, _
                    Optional ByVal ShowUnits As Boolean = True, _
                    Optional ByVal UnitsStyle As Long = 3, _
                    Optional ByVal PlacementStyle As Long = 2, _
                    Optional ByVal HorizontalText As Boolean = False, _
                    Optional ByVal BoxedText As Boolean = False, _
                    Optional ByVal LeadingZero As Boolean = True, _
                    Optional ByVal Prefix As String, _
                    Optional ByVal Suffix As String, _
                    Optional ByVal OutlineWidth As Double = -1, _
                    Optional ByVal Arrows As ArrowHead, _
                    Optional ByVal OutlineColor As Color, _
                    Optional ByVal TextFont As String, _
                    Optional ByVal TextSize As Double, _
                    Optional ByVal TextColor As Color _
                ) As Shape
    Dim DimType As cdrLinearDimensionType
    Dim TextX As Double, TextY As Double
    
    If Point1.PositionX = Point2.PositionX Then 'вертикальный
        DimType = cdrDimensionVertical
        TextX = Point1.PositionX + TextOffset
    ElseIf Point1.PositionY = Point2.PositionY Then 'горизонтальный
        DimType = cdrDimensionHorizontal
        TextY = Point1.PositionY + TextOffset
    Else 'наклонный
        DimType = cdrDimensionSlanted
        'TODO TextX = ?: TextY = ?
    End If
    
    Set MakeLinearDimension = _
    ActiveLayer.CreateLinearDimension( _
        Type:=DimType, _
        Point1:=Point1, _
        Point2:=Point2, _
        TextX:=TextX, TextY:=TextY, _
        OutlineWidth:=OutlineWidth, _
        Arrows:=Arrows, _
        OutlineColor:=OutlineColor, _
        TextFont:=TextFont, _
        TextSize:=TextSize, _
        TextColor:=TextColor _
    )
    
    Dim DimStyle As Style: Set DimStyle = _
        MakeLinearDimension.Style.GetProperty("dimension")
    With DimStyle
        .SetProperty "centerText", TextCentered
        .SetProperty "textStyle", DimensionStyle
        .SetProperty "precision", Precision
        .SetProperty "showUnits", ShowUnits
        .SetProperty "units", UnitsStyle
        .SetProperty "textPlacement", PlacementStyle
        .SetProperty "horizontalText", HorizontalText
        .SetProperty "boxAroundText", BoxedText
        .SetProperty "showLeadingZero", LeadingZero
        .SetProperty "prefix", Prefix
        .SetProperty "suffix", Suffix
        
        'TODO сделать сразу единой json-струтурой, см. Show DimStyle.ToString
    End With
    
End Function

Public Function MakeRect( _
                    ByVal x As Double, _
                    ByVal y As Double, _
                    ByVal Width As Double, _
                    ByVal Height As Double, _
                    Optional ByVal FillColor As Color, _
                    Optional ByVal OutlineColor As Color _
                ) As Shape
    Set MakeRect = _
        ActiveLayer.CreateRectangle2( _
            x - Width / 2, y - Height / 2, Width, Height _
        )
    With MakeRect
        If IsSome(FillColor) Then .Fill.ApplyUniformFill FillColor
        If IsSome(OutlineColor) Then .Outline.Color = OutlineColor
    End With
End Function

'не работает с поверклипом
Public Sub MatrixCopy(ByVal SourceShape As Shape, ByVal TargetShape As Shape)
    Dim tMatrix As typeMatrix
    With tMatrix
        SourceShape.GetMatrix .d11, .d12, .d21, .d22, .tx, .ty
        TargetShape.SetMatrix .d11, .d12, .d21, .d22, .tx, .ty
    End With
End Sub

'правильно перемещает Shape или ShapeRange на другой слой
Public Function MoveToLayer( _
                    ByVal MaybeShapeOrRange As Variant, _
                    ByVal MaybeLayer As Layer _
                )
    If VBA.IsEmpty(MaybeShapeOrRange) _
    Or VBA.IsEmpty(MaybeLayer) Then Exit Function
    ThrowIfNotShapeOrRange MaybeShapeOrRange
    
    Dim tSrcLayer() As Layer
    Dim tProps() As LayerProperties
    Dim tLayersCol As Collection
    Dim i&
    
    If TypeOf MaybeShapeOrRange Is Shape Then
    
        Set tLayersCol = New Collection
        tLayersCol.Add MaybeShapeOrRange.Layer
        
    ElseIf TypeOf MaybeShapeOrRange Is ShapeRange Then
        
        If MaybeShapeOrRange.Count < 1 Then Exit Function
        Set tLayersCol = ShapeRangeLayers(MaybeShapeOrRange)
        
    Else
    
        Throw "Type mismatch: MaybeShapeOrRange должен быть Shape или ShapeRange"
        Exit Function
    
    End If
    
    ReDim tSrcLayer(1 To tLayersCol.Count)
    ReDim tProps(1 To tLayersCol.Count)
    For i = 1 To tLayersCol.Count
        Set tSrcLayer(i) = tLayersCol(i)
        LayerPropsPreserveAndReset tSrcLayer(i), tProps(i)
    Next i
    MaybeShapeOrRange.MoveToLayer MaybeLayer
    For i = 1 To tLayersCol.Count
        LayerPropsRestore tSrcLayer(i), tProps(i)
    Next i

End Function

Public Sub NameShapes(ByVal Shapes As ShapeRange, ByVal Name As String)
    Dim Shape As Shape
    For Each Shape In Shapes
        Shape.Name = Name
    Next Shape
End Sub

Public Sub ResizeImageToDocumentResolution(ByVal ImageShape As Shape)
    With ImageShape.Bitmap
        ImageShape.SetSize _
            PixelsToDocUnits(.SizeWidth), PixelsToDocUnits(.SizeHeight)
    End With
End Sub

Public Sub ResizePageToShapes( _
               Optional ByVal SideMult As Double = 1, _
               Optional ByVal SideAdd As Double = 0 _
            )
    With ActivePage
        .SetSize .Shapes.All.SizeWidth * SideMult + SideAdd, _
                 .Shapes.All.SizeHeight * SideMult + SideAdd
        .Shapes.All.SetPositionEx cdrCenter, .CenterX, .CenterY
    End With
End Sub

'удаление сегмента
'автор: Alex Vakulenko http://www.oberonplace.com/vba/drawmacros/delsegment.htm
Public Sub SegmentDelete(ByVal Segment As Segment)
    If Not Segment.EndNode.IsEnding Then
        Segment.EndNode.BreakApart
        Set Segment = Segment.SubPath.LastSegment
    End If
    Segment.EndNode.Delete
End Sub

Public Sub Separate(ByRef Shapes As ShapeRange)
    Dim Shape As Shape
    Dim Result As New ShapeRange
    Dim i As Long
    For Each Shape In Shapes
        If Shape.Effects.Count > 0 Then
            For i = 1 To Shape.Effects.Count
                Result.AddRange Shape.Effects(i).Separate
            Next i
        Else
            Result.Add Shape
        End If
    Next
    Set Shapes = Result
End Sub

Public Sub SetDimensionPrecision( _
               ByVal DimensionShape As Shape, _
               ByVal Precision As Long _
           )
    SetDimensionProperty DimensionShape, "precision", Precision
End Sub

Public Sub SetDimensionProperty( _
               ByVal DimensionShape As Shape, _
               ByVal PropertyName As String, _
               ByVal Value As Variant _
           )
    DimensionShape.Style.GetProperty("dimension") _
        .SetProperty PropertyName, Value
End Sub

Public Sub SetDimensionShowUnits( _
               ByVal DimensionShape As Shape, _
               ByVal ShowUnits As Boolean _
           )
    SetDimensionProperty DimensionShape, "showUnits", ShowUnits
End Sub

Public Sub SetNoOutline( _
               ByVal Shapes As ShapeRange _
           )
    Dim Shape As Shape
    For Each Shape In Shapes
        Shape.Outline.SetNoOutline
    Next Shape
End Sub

'присвоить цвет абриса ренджу
Public Sub SetOutlineColor( _
               ByVal Shapes As ShapeRange, _
               ByVal Color As Color _
           )
    Dim Shape As Shape
    For Each Shape In Shapes
        Shape.Outline.Color.CopyAssign Color
    Next Shape
End Sub

Public Sub Simplify(ByRef Shapes As ShapeRange)
    Set Shapes = FindAllShapes(Shapes)
    Set Shapes = Shapes.UngroupAllEx
    Separate Shapes
    Set Shapes = Shapes.UngroupAllEx
    Shapes.ConvertToCurves
    Set Shapes = Shapes.UngroupAllEx
End Sub

Public Sub SwapOrientation(ByVal Page As Page)
    Dim x As Variant
    x = Page.SizeHeight
    Page.SizeHeight = Page.SizeWidth
    Page.SizeWidth = x
End Sub

Public Sub Trim( _
                ByVal TrimmerShape As Shape, _
                ByRef TargetShape As Shape, _
                Optional ByVal DeleteTrimmer As Boolean _
           )
    Set TargetShape = TrimmerShape.Trim(TargetShape, Not DeleteTrimmer)
End Sub

'обрезать битмап по CropEnvelopeShape, но по-умному,
'сначала кропнув на EXPANDBY пикселей побольше
Public Function TrimBitmap( _
                    ByVal BitmapShape As Shape, _
                    ByVal CropEnvelopeShape As Shape, _
                    Optional ByVal LeaveCropEnvelope As Boolean = True _
                ) As Shape

    Const EXPANDBY& = 2 'px
    
    Dim tCrop As Shape
    Dim tPxW#, tPxH#
    Dim tSaveUnit As cdrUnit

    If Not BitmapShape.Type = cdrBitmapShape Then Exit Function
    
    'save
    tSaveUnit = ActiveDocument.Unit
    
    ActiveDocument.Unit = cdrInch
    tPxW = 1 / BitmapShape.Bitmap.ResolutionX
    tPxH = 1 / BitmapShape.Bitmap.ResolutionY
    BitmapShape.Bitmap.ResetCropEnvelope
    Set tCrop = BitmapShape.Layer.CreateRectangle( _
                    CropEnvelopeShape.LeftX - tPxW * EXPANDBY, _
                    CropEnvelopeShape.TopY + tPxH * EXPANDBY, _
                    CropEnvelopeShape.RightX + tPxW * EXPANDBY, _
                    CropEnvelopeShape.BottomY - tPxH * EXPANDBY _
                )
    Set TrimBitmap = Intersect(tCrop, BitmapShape, False, False)
    If TrimBitmap Is Nothing Then
        tCrop.Delete
        GoTo Finally
    End If
    TrimBitmap.Bitmap.Crop
    Set TrimBitmap = _
        Intersect(CropEnvelopeShape, TrimBitmap, LeaveCropEnvelope, False)
    
Finally:
    'restore
    ActiveDocument.Unit = tSaveUnit
    
End Function

Public Function Weld(ByVal Shapes As ShapeRange) As Shape
    Set Weld = Shapes.FirstShape
    Do Until Shapes.Count = 1
        Shapes(1).CreateSelection
        Shapes(2).AddToSelection
        Shapes.Remove 1
        Shapes.Remove 1
        With ActiveSelectionRange
            Set Weld = .FirstShape.Weld(.LastShape)
        End With
        Shapes.Add Weld
    Loop
End Function

'===============================================================================
' # функции работы с файлами

Public Function AddProperEndingToPath(ByVal Path As String) As String
    If Not VBA.Right$(Path, 1) = "\" Then AddProperEndingToPath = Path & "\" _
    Else: AddProperEndingToPath = Path
End Function

'существует ли файл или папка (папка должна заканчиваться на "\")
Public Property Get FileExists(ByVal File As String) As Boolean
    If File = "" Then Exit Property
    FileExists = VBA.Len(VBA.Dir(File)) > 0
End Property

Public Property Get FindFileInGMSFolders(ByVal FileName As String) As String
    FindFileInGMSFolders = GMSManager.UserGMSPath & FileName
    If Not FileExists(FindFileInGMSFolders) Then _
        FindFileInGMSFolders = GMSManager.GMSPath & FileName
    If Not FileExists(FindFileInGMSFolders) Then _
        FindFileInGMSFolders = ""
End Property

Public Property Get FSO() As Scripting.FileSystemObject
    Set FSO = New Scripting.FileSystemObject
End Property

'возвращает имя файла без расширения
Public Property Get GetFileNameNoExt(ByVal FileName As String) As String
    If VBA.Right(FileName, 1) <> "\" And VBA.Len(FileName) > 0 Then
        GetFileNameNoExt = VBA.Left(FileName, _
            Switch _
                (VBA.InStr(FileName, ".") = 0, _
                    Len(FileName), _
                VBA.InStr(FileName, ".") > 0, _
                    VBA.InStrRev(FileName, ".") - 1))
    End If
End Property

'полное имя временного файла
Public Property Get GetTempFile() As String
    GetTempFile = GetTempFolder & GetTempFileName
End Property

'имя временного файла
Public Property Get GetTempFileName() As String
    GetTempFileName = "elvin_" & CreateGUID & ".tmp"
End Property

'находит временную папку
Public Property Get GetTempFolder() As String
    GetTempFolder = AddProperEndingToPath(VBA.Environ$("TEMP"))
    If FileExists(GetTempFolder) Then Exit Property
    GetTempFolder = AddProperEndingToPath(VBA.Environ$("TMP"))
    If FileExists(GetTempFolder) Then Exit Property
    GetTempFolder = "c:\temp\"
    If FileExists(GetTempFolder) Then Exit Property
    GetTempFolder = "c:\windows\temp\"
    If FileExists(GetTempFolder) Then Exit Property
End Property

Public Property Get GetFileName(ByVal File As String) As String
    GetFileName = VBA.Right(File, VBA.Len(File) - VBA.InStrRev(File, "\"))
End Property

Public Property Get GetFilePath(ByVal File As String) As String
    GetFilePath = VBA.Left(File, VBA.InStrRev(File, "\"))
End Property

Public Property Get GetFilesFromFolder(ByVal Path As String) As VBA.Collection
    Set GetFilesFromFolder = SequenceToCollection(FSO.GetFolder(Path).Files)
End Property

Public Property Get GetRandomFilesFromFolder( _
                        ByVal Path As String, _
                        Optional ByVal NumberOfFiles As Long = 1 _
                    ) As VBA.Collection
    Set GetRandomFilesFromFolder = New Collection
    Dim Files As Collection: Set Files = GetFilesFromFolder(Path)
    If Files.Count = 0 Then Exit Property
    Do While GetRandomFilesFromFolder.Count < NumberOfFiles
        GetRandomFilesFromFolder.Add Files(RndLong(1, Files.Count))
    Loop
End Property

'создаёт папку, если не было
'возвращает Path обратно (для inline-использования)
Public Function MakeDir(ByVal Path As String) As String
    Dim FS As New FileSystemObject
    If Not FS.FolderExists(Path) Then FS.CreateFolder Path
    MakeDir = Path
End Function

'создаёт путь, если не было
'возвращает Path обратно (для inline-использования)
Public Function MakePath(ByVal Path As String) As String
    If Not FSO.FolderExists(Path) Then MakeNestedFolders Path
    MakePath = Path
End Function
Private Sub MakeNestedFolders(ByVal Path As String)
    Dim SubPath As Variant
    Dim StrCheckPath As String
    For Each SubPath In VBA.Split(Path, "\")
        StrCheckPath = StrCheckPath & SubPath & "\"
        If VBA.Dir(StrCheckPath, vbDirectory) = vbNullString Then
            VBA.MkDir StrCheckPath
        End If
    Next
End Sub

'загружает файл в строку
Public Function ReadFile(ByVal File As String) As String
    Dim tFileNum As Long
    tFileNum = FreeFile
    Open File For Input As #tFileNum
    ReadFile = Input(LOF(tFileNum), tFileNum)
    Close #tFileNum
End Function

'загружает файл в строку через ADODB, можно задать кодировку
Public Function ReadFileAD( _
                    ByVal File As String, _
                    Optional ByVal CharSet As String = "utf-8" _
                ) As String
    Dim ADODB As Object
    Set ADODB = VBA.CreateObject("ADODB.Stream")
    ADODB.CharSet = CharSet
    ADODB.Open
    ADODB.LoadFromFile File
    ReadFileAD = ADODB.ReadText()
    ADODB.Close
End Function

'заменяет расширение файлу на заданное
Public Function SetFileExt( _
                    ByVal SourceFile As String, _
                    ByVal NewExt As String _
                ) As String
    If Right(SourceFile, 1) <> "\" And Len(SourceFile) > 0 Then
        SetFileExt = GetFileNameNoExt(SourceFile$) & "." & NewExt
    End If
End Function

'сохраняет строку Content в файл, перезаписывая, делая в процессе temp файл,
'и оставляя бэкап, если необходимо
Public Sub WriteFile( _
               ByVal Content As String, _
               ByVal File As String, _
               Optional ByVal KeepBak As Boolean = False _
           )

    Dim tFileNum As Long
    tFileNum = FreeFile
    Dim tBak As String
    tBak = SetFileExt(File, "bak")
    Dim tTemp As String
    
    If KeepBak Then
        If FileExists(File) Then FileCopy File, tBak
    Else
        If FileExists(File) Then
            tTemp = GetFilePath(File) & GetTempFileName
            FileCopy File, tTemp
        End If
    End If
        
    Open File For Output Access Write As #tFileNum
    Print #tFileNum, Content
    Close #tFileNum
    
    On Error Resume Next
        If Not KeepBak Then Kill tTemp
    On Error GoTo 0

End Sub

'сохраняет строку Content в файл через ADODB, можно задать кодировку
Public Sub WriteFileAD( _
               ByVal File As String, _
               ByVal Content As String, _
               Optional CharSet As String = "utf-8" _
           )
    Dim ADODB As Object
    Set ADODB = VBA.CreateObject("ADODB.Stream")
    ADODB.CharSet = CharSet
    ADODB.Open
    ADODB.WriteText Content
    ADODB.SaveToFile File
    ADODB.Close
End Sub

'===============================================================================
' # прочие функции

Public Sub AppendCollection( _
               ByVal Destination As Collection, _
               ByVal SourceToAdd As Collection _
           )
    Dim Item As Variant
    For Each Item In SourceToAdd
        Destination.Add Item
    Next Item
End Sub

Public Function AskForLong( _
                    ByVal Message As String, _
                    ByRef Num As Long, _
                    Optional ByVal Title As String _
                ) As Boolean
    Dim Out As Variant
    Out = VBA.InputBox(Message, Title, Num)
    If Not IsLong(Out) Then Exit Function
    AskForLong = True
    Num = Out
End Function

Public Function AskYesNo( _
                    ByVal Message As String, _
                    Optional ByVal Title As String _
                ) As Boolean
    AskYesNo = (VBA.MsgBox(Message, vbYesNo, Title) = 6)
End Function

Public Sub Assign(ByRef Destination As Variant, ByVal x As Variant)
    If VBA.IsObject(x) Then
        Set Destination = x
    Else
        Destination = x
    End If
End Sub

Public Sub BoostStart( _
               Optional ByVal UndoGroupName As String = "" _
           )
    If Not UndoGroupName = vbNullString _
   And Not ActiveDocument Is Nothing Then _
        ActiveDocument.BeginCommandGroup UndoGroupName
    #If Not DebugMode = 1 Then
    If Not Optimization Then Optimization = True
    #End If
    If EventsEnabled Then EventsEnabled = False
    If Not ActiveDocument Is Nothing Then
        With ActiveDocument
            .SaveSettings
            .PreserveSelection = False
            .Unit = cdrMillimeter
            .WorldScale = 1
            .ReferencePoint = cdrCenter
        End With
    End If
End Sub
Public Sub BoostFinish(Optional ByVal EndUndoGroup As Boolean = True)
    If Not EventsEnabled Then EventsEnabled = True
    If Optimization Then Optimization = False
    If Not ActiveDocument Is Nothing Then
        With ActiveDocument
            .RestoreSettings
            .PreserveSelection = True
            If EndUndoGroup Then .EndCommandGroup
        End With
        ActiveWindow.Refresh
    End If
    Application.Refresh
    Application.Windows.Refresh
End Sub

'находит ближайшее к Value число, которое делится на Divisor без остатка
Public Property Get ClosestDividend( _
                        ByVal Value As Double, _
                        ByVal Divisor As Double _
                    ) As Double
    Dim q As Long: q = Fix(Value / Divisor)
    Dim n1 As Double: n1 = Divisor * q

    Dim n2 As Double
    If (Value * Divisor) > 0 Then
        n2 = Divisor * (q + 1)
    Else
        n2 = Divisor * (q - 1)
    End If

    If Abs(Value - n1) < Abs(Value - n2) Then
        ClosestDividend = n1
    Else
        ClosestDividend = n2
    End If
End Property

Public Property Get Collection( _
                        ParamArray Elements() As Variant _
                    ) As VBA.Collection
    Set Collection = New VBA.Collection
    Dim Element As Variant
    For Each Element In Elements
        Collection.Add Element
    Next Element
End Property

Public Property Get Contains( _
                        ByRef ContainerSeq As Variant, _
                        ByRef Item As Variant _
                    ) As Boolean
    Dim Element As Variant
    For Each Element In ContainerSeq
        If Same(Item, Element) Then
            Contains = True
            Exit Property
        End If
    Next Element
End Property

Public Property Get ContainsAll( _
                        ByRef ContainerSeq As Variant, _
                        ByRef ItemsSeq As Variant _
                    ) As Boolean
    Dim Item As Variant
    For Each Item In ItemsSeq
        If Not Contains(ContainerSeq, Item) Then Exit Property
    Next Item
    ContainsAll = True
End Property

Public Property Get Count( _
                        ByRef Arr As Variant, _
                        Optional ByVal Dimension As Long = 1 _
                    ) As Long
    Count = UBound(Arr, Dimension) - LBound(Arr, Dimension) + 1
End Property

Public Property Get CountCopy( _
                        ByVal Arr As Variant, _
                        Optional ByVal Dimension As Long = 1 _
                    ) As Long
    CountCopy = Count(Arr, Dimension)
End Property

'https://www.codegrepper.com/code-examples/vb/excel+vba+generate+guid+uuid
Public Function CreateGUID( _
                    Optional ByVal Lowercase As Boolean, _
                    Optional ByVal Parens As Boolean _
                ) As String
    Dim k As Long, H As String
    CreateGUID = VBA.Space(36)
    For k = 1 To VBA.Len(CreateGUID)
        VBA.Randomize
        Select Case k
            Case 9, 14, 19, 24:         H = "-"
            Case 15:                    H = "4"
            Case 20:                    H = VBA.Hex(VBA.Rnd * 3 + 8)
            Case Else:                  H = VBA.Hex(VBA.Rnd * 15)
        End Select
        Mid(CreateGUID, k, 1) = H
    Next
    If Lowercase Then CreateGUID = VBA.LCase$(CreateGUID)
    If Parens Then CreateGUID = "{" & CreateGUID & "}"
End Function

Public Function DebugOut( _
               ByVal Context As Variant, _
               ParamArray Output() As Variant _
            )
    #If DebugMode = 1 Then
    If Not VBA.VarType(Context) = vbString Then
        Context = VBA.TypeName(Context)
    End If
    Debug.Print "<" & Context & "> " & VBA.Join(Output, " ")
    #End If
End Function

Public Property Get Deduplicate(ByVal Sequence As Variant) As VBA.Collection
    Set Deduplicate = New Collection
    Dim Item As Variant
    For Each Item In Sequence
        If Not Contains(Deduplicate, Item) Then Deduplicate.Add Item
    Next Item
End Property

'первое число в строке, всегда положительное целое
'dfhfgh-072.88fdfg12 -> 72
Public Property Get FindFirstInteger(ByVal Str As String) As Variant
    Dim i As Long
    Dim TempNumber As String
    For i = 1 To Len(Str)
        'проверка, является ли символ цифрой
        'Like чуть быстрее, чем IsNumeric
        If Mid(Str, i, 1) Like "#" Then
            TempNumber = TempNumber & Mid(Str, i, 1)
        'если уже найдено число и встречен нецифровой символ
        ElseIf Not TempNumber = vbNullString Then
            Exit For
        End If
    Next i
    If Not TempNumber = vbNullString Then FindFirstInteger = CLng(TempNumber)
End Property

Public Property Get FindMaxNumberIndex(ByVal Numbers As Collection) As Long
    FindMaxNumberIndex = 1
    Dim i As Long
    For i = 1 To Numbers.Count
        If VBA.IsNumeric(Numbers(i)) Then
            If Numbers(i) > Numbers(FindMaxNumberIndex) Then _
                FindMaxNumberIndex = i
        End If
    Next i
End Property

Public Property Get FindMinNumberIndex(ByVal Numbers As Collection) As Long
    FindMinNumberIndex = 1
    Dim i As Long
    For i = 1 To Numbers.Count
        If VBA.IsNumeric(Numbers(i)) Then
            If Numbers(i) < Numbers(FindMinNumberIndex) Then _
                FindMinNumberIndex = i
        End If
    Next i
End Property

Public Function GetCollectionCopy(ByVal Source As Collection) As Collection
    Set GetCollectionCopy = New Collection
    Dim Item As Variant
    For Each Item In Source
        GetCollectionCopy.Add Item
    Next Item
End Function

Public Function GetCollectionFromDictionary( _
                    ByVal Dictionary As Scripting.IDictionary _
                ) As Collection
    Set GetCollectionFromDictionary = New Collection
    Dim Item As Variant
    For Each Item In Dictionary.Items
        GetCollectionFromDictionary.Add Item
    Next Item
End Function

Public Function GetDictionaryCopy( _
                    ByVal Source As Scripting.IDictionary _
                ) As Scripting.Dictionary
    Set GetDictionaryCopy = New Scripting.Dictionary
    Dim Key As Variant
    For Each Key In Source.Keys
        GetDictionaryCopy.Add Key, Source.Item(Key)
    Next Key
End Function

'является ли число чётным :) Что такое Even и Odd запоминать лень...
Public Property Get IsChet(ByVal x As Variant) As Boolean
    If x Mod 2 = 0 Then IsChet = True Else IsChet = False
End Property

'делится ли Number на Divider нацело
Public Property Get IsDivider( _
                    ByVal Number As Long, _
                    ByVal Divider As Long _
                ) As Boolean
    If Number Mod Divider = 0 Then IsDivider = True Else IsDivider = False
End Property

Public Property Get IsLong(ByVal x As Variant) As Boolean
    If Not VBA.IsNumeric(x) Then Exit Property
    If CLng(x) <> VBA.Val(x) Then Exit Property
    IsLong = True
End Property

Public Property Get IsLowerCase(ByVal Str As String) As Boolean
    If VBA.LCase(Str) = Str Then IsLowerCase = True
End Property

Public Property Get IsNone(ByRef Unknown As Variant) As Boolean
    If VBA.IsNull(Unknown) _
    Or VBA.IsEmpty(Unknown) _
    Or VBA.IsMissing(Unknown) Then
        IsNone = True
        Exit Property
    End If
    If VBA.IsObject(Unknown) Then
        If Unknown Is Nothing Then
            IsNone = True
            Exit Property
        End If
    End If
End Property

Public Property Get IsUpperCase(ByVal Str As String) As Boolean
    If VBA.UCase(Str) = Str Then IsUpperCase = True
End Property

Public Property Get IsSome(ByRef Unknown As Variant) As Boolean
    IsSome = Not IsNone(Unknown)
End Property

Public Property Get MatchAll( _
                        ByVal Reference As Variant, _
                        ByVal SamplesSeq As Variant _
                    ) As Boolean
    Dim Sample As Variant
    For Each Sample In SamplesSeq
        If Not Same(Sample, Reference) Then Exit Property
    Next Sample
    MatchAll = True
End Property

Public Property Get MatchAllOf( _
                        ByVal Reference As Variant, _
                        ParamArray Samples() As Variant _
                    ) As Boolean
    MatchAllOf = MatchAll(Reference, Samples)
End Property

Public Property Get MatchAny( _
                        ByVal Reference As Variant, _
                        ByVal SamplesSeq As Variant _
                    ) As Boolean
    Dim Sample As Variant
    For Each Sample In SamplesSeq
        If Same(Sample, Reference) Then
            MatchAny = True
            Exit Property
        End If
    Next Sample
End Property

Public Property Get MatchAnyOf( _
                        ByVal Reference As Variant, _
                        ParamArray Samples() As Variant _
                    ) As Boolean
    MatchAnyOf = MatchAny(Reference, Samples)
End Property

Public Property Get Max(ByRef Sequence As Variant) As Variant
    Dim Item As Variant
    For Each Item In Sequence
        If VBA.IsNumeric(Item) Then
            If Item > Max Then Max = Item
        End If
    Next Item
End Property

Public Property Get MaxOfTwo( _
                        ByVal Value1 As Variant, _
                        ByVal Value2 As Variant _
                    ) As Variant
    If Value1 > Value2 Then MaxOfTwo = Value1 Else MaxOfTwo = Value2
End Property

Public Function MeasureStart()
    StartTime = Timer
End Function
Public Function MeasureFinish(Optional ByVal Message As String = "")
    Debug.Print Message & CStr(Round(Timer - StartTime, 3)) & " секунд"
End Function

Public Property Get Min(ByRef Sequence As Variant) As Variant
    Dim Item As Variant
    For Each Item In Sequence
        If VBA.IsNumeric(Item) Then
            If Item < Min Then
                Min = Item
            End If
        End If
    Next Item
End Property

Public Property Get MinOfTwo( _
                        ByVal Value1 As Variant, _
                        ByVal Value2 As Variant _
                    ) As Variant
    If Value1 < Value2 Then MinOfTwo = Value1 Else MinOfTwo = Value2
End Property

Public Sub Notify(ByVal Message As String, Optional ByVal Title As String)
    VBA.MsgBox Message, vbInformation, Title
End Sub

'возвращает True, если Value - это объект и при этом не Nothing
Public Property Get ObjectAssigned(ByRef Variable As Variant) As Boolean
    If Not VBA.IsObject(Variable) Then Exit Property
    ObjectAssigned = Not Variable Is Nothing
End Property

Public Property Get Pack(ParamArray Items() As Variant) As Variant
    Dim Length As Long
    Length = UBound(Items) - LBound(Items) + 1
    If Length = 0 Then
        Pack = Array()
        Exit Property
    End If
    ReDim Result(1 To Length) As Variant
    Dim Index As Long
    Dim Item As Variant
    For Each Item In Items
        Index = Index + 1
        Assign Result(Index), Item
    Next Item
    Pack = Result
End Property

'создаёт ShapeRange из Shape/Shapes/ShapeRange
Public Function PackShapes(ParamArray Shapes() As Variant) As ShapeRange
    Set PackShapes = CreateShapeRange
    Dim Item As Variant
    For Each Item In Shapes
        If TypeOf Item Is Shape Then
            PackShapes.Add Item
        ElseIf TypeOf Item Is ShapeRange Then
            PackShapes.AddRange Item
        ElseIf TypeOf Item Is Shapes Then
            PackShapes.AddRange Item.All
        Else
            Throw "Не является шейпом"
        End If
    Next Item
End Function

Public Sub RemoveElementFromCollection( _
               ByVal Collection As Collection, _
               ByVal Element As Variant _
           )
    If Collection.Count = 0 Then Exit Sub
    Dim i As Long
    For i = 1 To Collection.Count
        If Same(Element, Collection(i)) Then
            Collection.Remove i
            Exit Sub
        End If
    Next i
End Sub

Private Sub Resize(ByRef Arr As Variant, ByVal Length As Long)
    ReDim Preserve Arr(LBound(Arr) To LBound(Arr) + Length - 1)
End Sub

'случайное число от LowerBound до UpperBound
Public Function RndDouble(LowerBound As Double, UpperBound As Double) As Double
    RndDouble = (UpperBound - LowerBound + 1) * VBA.Rnd + LowerBound
End Function

'случайное целое от LowerBound до UpperBound
Public Function RndLong(LowerBound As Long, UpperBound As Long) As Long
    RndLong = VBA.Int((UpperBound - LowerBound + 1) * VBA.Rnd + LowerBound)
End Function

'выводит информацию о переменной / её значение в окно immediate
Public Sub Show(ByVal Some As Variant)
    Debug.Print ToShowable(Some)
End Sub

Public Property Get Same( _
                        ByRef x As Variant, _
                        ByRef y As Variant _
                    ) As Boolean
    If VBA.IsObject(x) And VBA.IsObject(y) Then
        Same = x Is y
    ElseIf Not VBA.IsObject(x) And Not VBA.IsObject(y) Then
        Same = (x = y)
    End If
End Property

Public Property Get SequenceToCollection( _
                        ByVal Sequence As Variant _
                    ) As VBA.Collection
    Set SequenceToCollection = New Collection
    Dim Item As Variant
    For Each Item In Sequence
        SequenceToCollection.Add Item
    Next Item
End Property

Public Property Get SequenceToShowable(ByVal Sequence As Variant) As String
    Dim Result As String
    Dim Item As Variant
    For Each Item In Sequence
        Result = Result & ToShowable(Item) & ", "
    Next Item
    If VBA.Len(Result) > 2 Then Result = VBA.Left(Result, VBA.Len(Result) - 2)
    SequenceToShowable = "[" & Result & "]"
End Property

'bubble sort
'https://stackoverflow.com/a/3588073/3700481
Public Sub SortCollection(ByVal Collection As Collection)
    If Collection.Count < 2 Then Exit Sub
    Dim i As Long, j As Long
    Dim Temp As Variant
    'Two loops to bubble sort
    For i = 1 To Collection.Count - 1
        For j = i + 1 To Collection.Count
            If Collection(i) > Collection(j) Then
                'store the lesser item
                Temp = Collection(j)
                'remove the lesser item
                Collection.Remove j
                're-add the lesser item before the
                'greater Item
                Collection.Add Item:=Temp, Before:=i
            End If
        Next j
    Next i
End Sub

Public Sub Swap(ByRef x As Variant, ByRef y As Variant)
    Dim z As Variant
    z = x
    x = y
    y = z
End Sub

Public Sub Throw(Optional ByVal Message As String = "Неизвестная ошибка")
    VBA.Err.Raise CUSTOM_ERROR, , Message
End Sub

Public Property Get ToShowable(ByVal Some As Variant) As String
    If VBA.IsObject(Some) Then
        If Some Is Nothing Then
            ToShowable = "[Nothing]"
        ElseIf TypeOf Some Is VBA.Collection Then
            ToShowable = SequenceToShowable(Some)
        ElseIf TypeOf Some Is Scripting.Dictionary Then
            ToShowable = SequenceToShowable(Some.Items)
        ElseIf TypeOf Some Is Scripting.File Then
            ToShowable = Some.Name
        ElseIf TypeOf Some Is Scripting.Files Then
            ToShowable = SequenceToShowable(Some)
        Else
            ToShowable = "[Object:" & VBA.TypeName(Some) & "]"
        End If
    ElseIf VBA.IsMissing(Some) Then
        ToShowable = "[Missing]"
    ElseIf VBA.IsArray(Some) Then
        ToShowable = SequenceToShowable(Some)
    Else
        Select Case VBA.VarType(Some)
            Case vbEmpty: ToShowable = "[Empty]"
            Case vbNull: ToShowable = "[Null]"
            Case vbError: ToShowable = "[Error]"
            Case Else: ToShowable = Some
        End Select
    End If
End Property

Public Property Get ToStr( _
                        ByVal Some As Variant, _
                        Optional ByVal DecimalSeparator As String = "," _
                    ) As String
    If VBA.VarType(Some) = vbString Then
        ToStr = Some
    Else
        If DecimalSeparator = "," Then
            ToStr = CStr(Some)
        Else
            ToStr = VBA.Replace(CStr(Some), ",", DecimalSeparator)
        End If
    End If
End Property

Public Sub Warn(ByVal Message As String, Optional ByVal Title As String)
    VBA.MsgBox Message, vbExclamation, Title
End Sub

'===============================================================================
' # приватные функции модуля

'для IsOverlap
Private Function IsIntersectReady(ByVal Shape As Shape) As Boolean
    With Shape
        If .Type = cdrCustomShape Or _
             .Type = cdrBlendGroupShape Or _
             .Type = cdrOLEObjectShape Or _
             .Type = cdrExtrudeGroupShape Or _
             .Type = cdrContourGroupShape Or _
             .Type = cdrBevelGroupShape Or _
             .Type = cdrConnectorShape Or _
             .Type = cdrMeshFillShape Or _
             .Type = cdrTextShape Then
            IsIntersectReady = False
        Else
            IsIntersectReady = True
        End If
    End With
End Function

Private Sub LayerPropsPreserve(ByVal L As Layer, ByRef Props As LayerProperties)
    With Props
        .Visible = L.Visible
        .Printable = L.Printable
        .Editable = L.Editable
    End With
End Sub
Private Sub LayerPropsReset(ByVal L As Layer)
    With L
        If Not .Visible Then .Visible = True
        If Not .Printable Then .Printable = True
        If Not .Editable Then .Editable = True
    End With
End Sub
Private Sub LayerPropsRestore(ByVal L As Layer, ByRef Props As LayerProperties)
    With Props
        If L.Visible <> .Visible Then L.Visible = .Visible
        If L.Printable <> .Printable Then L.Printable = .Printable
        If L.Editable <> .Editable Then L.Editable = .Editable
    End With
End Sub
Private Sub LayerPropsPreserveAndReset( _
                ByVal L As Layer, _
                ByRef Props As LayerProperties _
            )
    LayerPropsPreserve L, Props
    LayerPropsReset L
End Sub

Private Sub ThrowIfNotShapeOrRange( _
                ByVal MaybeShapeOrRange As Variant _
            )
    If VBA.IsObject(MaybeShapeOrRange) Then
        If Not MaybeShapeOrRange Is Nothing Then
            If TypeOf MaybeShapeOrRange Is Shape _
            Or TypeOf MaybeShapeOrRange Is ShapeRange Then _
                Exit Sub
        End If
    End If
    Throw "Тип должен быть Shape или ShapeRange"
End Sub

Private Sub ThrowIfNotCollectionOrArray(ByRef CollectionOrArray As Variant)
    If VBA.IsObject(CollectionOrArray) Then _
        If TypeOf CollectionOrArray Is Collection Then Exit Sub
    If VBA.IsArray(CollectionOrArray) Then Exit Sub
    VBA.Err.Raise _
        13, Source:="LibCore", _
        Description:="Type mismatch: CollectionOrArray должен быть Collection или Array"
End Sub
