Attribute VB_Name = "LibCore"
'===============================================================================
'   ������          : LibCore
'   ������          : 2023.03.12
'   �����           : elvin-nsk (me@elvin.nsk.ru)
'   ����������� ��� : dizzy (�� ������� CtC), Alex Vakulenko
'                     � ��.
'   ��������        : ���������� ������� ��� ��������
'   �������������   :
'   �����������     : ���������������
'===============================================================================

Option Explicit

'===============================================================================
' # ��������� ���������� ������

Private Type typeLayerProps
    Visible As Boolean
    Printable As Boolean
    Editable As Boolean
End Type

Private StartTime As Double

'===============================================================================
' # ��������� ����������

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

Public Const CustomError = vbObjectError Or 32

'===============================================================================
' # ������� ������ � ��������� ���������� �� �������� ������

'���������� ������� ������ �����/�������/��������
Public Property Get AverageDim(ByVal ShapeOrRangeOrPage As Object) As Double
    If Not TypeOf ShapeOrRangeOrPage Is Shape _
   And Not TypeOf ShapeOrRangeOrPage Is ShapeRange _
   And Not TypeOf ShapeOrRangeOrPage Is Page Then
        Err.Raise 13, Source:="AverageDim", _
                  Description:="Type mismatch: ShapeOrRangeOrPage ������ ���� Shape, ShapeRange ��� Page"
        Exit Property
    End If
    AverageDim = (ShapeOrRangeOrPage.SizeWidth + ShapeOrRangeOrPage.SizeHeight) _
               / 2
End Property

'������� ��� �����, ������� ����� � �����������, � ���������,
'����������� �������� �����-���������� � �����-������
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

'������� ��� ����� � ������ ������, ������� ����� � �����������, � ���������
Public Property Get FindShapesByName( _
                        ByVal Shapes As ShapeRange, _
                        ByVal Name As String _
                    ) As ShapeRange
    Set FindShapesByName = _
        FindAllShapes(Shapes).Shapes.FindShapes( _
            Name:=Name, Recursive:=False _
        )
End Property

'������� ��� ����� � ������� �������, ������� ����� � �����������, � ���������
'� �������� �� � ��� ������� � �������, ��� ���� = ���
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

'������� ��� �����, ����� ����� ������� ��������� � NamePart,
'������� ����� � �����������, � ���������
Public Property Get FindShapesByNamePart( _
                        ByVal Shapes As ShapeRange, _
                        ByVal NamePart As String _
                    ) As ShapeRange
    Set FindShapesByNamePart = FindAllShapes(Shapes).Shapes.FindShapes( _
                                   Query:="@Name.Contains('" & NamePart & "')" _
                               )
End Property

'������� ����������, ��� ��������
Public Property Get FindPowerClips(ByVal Shapes As ShapeRange) As ShapeRange
    Set FindPowerClips = CreateShapeRange
    Dim Shape As Shape
    For Each Shape In Shapes
        If IsValidShape(Shape) Then _
            If Not Shape.PowerClip Is Nothing Then FindPowerClips.Add Shape
    Next Shape
End Property

'������� ���������� �����������, ��� ��������
Public Property Get FindShapesInPowerClips( _
                        ByVal Shapes As ShapeRange _
                    ) As ShapeRange
    Dim Shape As Shape
    Set FindShapesInPowerClips = CreateShapeRange
    For Each Shape In FindPowerClips(Shapes)
        FindShapesInPowerClips.AddRange Shape.PowerClip.Shapes.All
    Next Shape
End Property

'������: https://community.coreldraw.com/talk/coreldraw_graphics_suite_x4/f/coreldraw-graphics-suite-x4/57576/macro-list-fonts-within-a-text-file
Public Sub FindFontsInRange( _
                ByVal TextRange As TextRange, _
                ByVal ioFonts As Collection _
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
        FindFontsInRange Before, ioFonts
        FindFontsInRange After, ioFonts
    Else
        AddFontToCollection FontName, ioFonts
    End If
End Sub
'+++
Private Sub AddFontToCollection( _
                ByVal FontName As String, _
                ByVal ioFonts As Collection _
            )
    Dim Font As Variant
    Dim Found As Boolean
    Found = False
    For Each Font In ioFonts
        If Font = FontName Then
            Found = True
            Exit For
        End If
    Next Font
    If Not Found Then ioFonts.Add FontName
End Sub

Public Property Get DiffWithinTolerance( _
                     ByVal Number1 As Variant, _
                     ByVal Number2 As Variant, _
                     ByVal Tolerance As Variant _
                 ) As Boolean
    DiffWithinTolerance = VBA.Abs(Number1 - Number2) < Tolerance
End Property

'���������� ��� ����� �� ���� ����� ������� ��������, �� ��������� - ��� ������-���� � ��� ������
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

'���������� ��������� ���� � ������� ��������, ����� ������� �������� NamePart
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

'����� �������� ���� �� ���� ���������� (�����������, ��� ����� �� �����)
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

'���������� ������� ������� �����/�������/��������
Public Property Get GreaterDim(ByVal ShapeOrRangeOrPage As Object) As Double
    If Not TypeOf ShapeOrRangeOrPage Is Shape And Not TypeOf ShapeOrRangeOrPage Is ShapeRange And Not TypeOf ShapeOrRangeOrPage Is Page Then
        Err.Raise 13, Source:="GreaterDim", _
                  Description:="Type mismatch: ShapeOrRangeOrPage ������ ���� Shape, ShapeRange ��� Page"
        Exit Property
    End If
    If ShapeOrRangeOrPage.SizeWidth > ShapeOrRangeOrPage.SizeHeight Then
        GreaterDim = ShapeOrRangeOrPage.SizeWidth
    Else
        GreaterDim = ShapeOrRangeOrPage.SizeHeight
    End If
End Property

Public Property Get IsColor(ByRef MaybeColor As Variant) As Boolean
    If Not AssignedObject(MaybeColor) Then Exit Property
    IsColor = TypeOf MaybeColor Is Color
End Property

Public Property Get IsCurve(ByRef MaybeCurve As Variant) As Boolean
    If Not AssignedObject(MaybeCurve) Then Exit Property
    IsCurve = TypeOf MaybeCurve Is Curve
End Property

Public Property Get IsDocument(ByRef MaybeDocument As Variant) As Boolean
    If Not AssignedObject(MaybeDocument) Then Exit Property
    IsDocument = TypeOf MaybeDocument Is Document
End Property

'True, ���� Value - �������� ��� ����������� ������ (�� �������, �� ������...)
Public Property Get IsJust(ByRef Value As Variant) As Boolean
    IsJust = Not (VBA.IsError(Value) Or IsVoid(Value))
End Property

'�������� �� ����/������/�������� ���������
Public Property Get IsLandscape(ByVal ShapeOrRangeOrPage As Object) As Boolean
    If Not TypeOf ShapeOrRangeOrPage Is Shape _
   And Not TypeOf ShapeOrRangeOrPage Is ShapeRange _
   And Not TypeOf ShapeOrRangeOrPage Is Page Then
        Err.Raise 13, Source:="IsLandscape", _
                  Description:="Type mismatch: ShapeOrRangeOrPage ������ ���� Shape, ShapeRange ��� Page"
        Exit Property
    End If
    If ShapeOrRangeOrPage.SizeWidth > ShapeOrRangeOrPage.SizeHeight Then
        IsLandscape = True
    Else
        IsLandscape = False
    End If
End Property

Public Property Get IsLayer(ByRef MaybeLayer As Variant) As Boolean
    If Not AssignedObject(MaybeLayer) Then Exit Property
    IsLayer = TypeOf MaybeLayer Is Layer
End Property

'todo: ��������� ��� �������
Public Property Get IsOverlap( _
                    ByVal FirstShape As Shape, _
                    ByVal SecondShape As Shape _
                ) As Boolean
    
    Dim tIS As Shape
    Dim tShape1 As Shape, tShape2 As Shape
    Dim tBound1 As Shape, tBound2 As Shape
    Dim tProps As typeLayerProps
    
    If FirstShape.Type = cdrConnectorShape _
    Or SecondShape.Type = cdrConnectorShape Then _
        Exit Property
    
    '���������� ����� ���� ��� ��������
    Dim tLayer As Layer: Set tLayer = ActiveLayer
    '���������� ��������� ������� ����
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
    
    '���������� �� �� �����
    LayerPropsRestore FirstShape.Layer, tProps
    tLayer.Activate

End Property

'IsOverlap ��������� �������� - ������ �� ���������,
'�� ���� ��������� �������� � � ����������� ������� ��� ����������
Public Property Get IsOverlapBox( _
                    ByVal FirstShape As Shape, _
                    ByVal SecondShape As Shape _
                ) As Boolean
    Dim tShape As Shape
    Dim tProps As typeLayerProps
    '���������� ����� ���� ��� ��������
    Dim tLayer As Layer: Set tLayer = ActiveLayer
    '���������� ��������� ������� ����
    FirstShape.Layer.Activate
    LayerPropsPreserveAndReset FirstShape.Layer, tProps
    Dim tRect As Rect
    Set tRect = FirstShape.BoundingBox.Intersect(SecondShape.BoundingBox)
    If tRect.Width = 0 And tRect.Height = 0 Then
        IsOverlapBox = False
    Else
        IsOverlapBox = True
    End If
    '���������� �� �� �����
    LayerPropsRestore FirstShape.Layer, tProps
    tLayer.Activate
End Property

Public Property Get IsPage(ByRef MaybePage As Variant) As Boolean
    If Not AssignedObject(MaybePage) Then Exit Property
    IsPage = TypeOf MaybePage Is Page
End Property

Public Property Get IsRect(ByRef MaybeRect As Variant) As Boolean
    If Not AssignedObject(MaybeRect) Then Exit Property
    IsRect = TypeOf MaybeRect Is Rect
End Property

Public Property Get IsSameColor( _
                    ByVal MaybeColor1 As Variant, _
                    ByVal MaybeColor2 As Variant _
                ) As Boolean
    If VBA.IsEmpty(MaybeColor1) Or VBA.IsEmpty(MaybeColor2) Then Exit Property
    IsSameColor = MaybeColor1.IsSame(MaybeColor2)
End Property

'�������� �� ������ �����������, ������������ ���� ��� ������ � ����� �����
'(underlying dubs)
Public Property Get IsSameCurves( _
                    ByVal Curve1 As Curve, _
                    ByVal Curve2 As Curve _
                ) As Boolean
    Dim tNode As Node
    Dim Tolerance As Double
    '������ = 0.001 ��
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
    If Not AssignedObject(MaybeNode) Then Exit Property
    IsNode = TypeOf MaybeNode Is Node
End Property

Public Property Get IsSegment(ByRef MaybeSegment As Variant) As Boolean
    If Not AssignedObject(MaybeSegment) Then Exit Property
    IsSegment = TypeOf MaybeSegment Is Segment
End Property

Public Property Get IsShape(ByRef MaybeShape As Variant) As Boolean
    If Not AssignedObject(MaybeShape) Then Exit Property
    IsShape = TypeOf MaybeShape Is Shape
End Property

Public Property Get IsShapeRange(ByRef MaybeShapeRange As Variant) As Boolean
    If Not AssignedObject(MaybeShapeRange) Then Exit Property
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
    If Not AssignedObject(MaybeSubPath) Then Exit Property
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

Public Property Get GetCurve(ByVal MaybeShape As Variant) As Curve
    If Not IsShape(MaybeShape) Then GoTo Fail
    Dim Temp As Double
    On Error GoTo Fail
    Set GetCurve = MaybeShape.Curve
    On Error GoTo 0
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

'���������� ������� ������� �����/�������/��������
Public Property Get LesserDim(ByVal ShapeOrRangeOrPage As Object) As Double
    If Not TypeOf ShapeOrRangeOrPage Is Shape _
   And Not TypeOf ShapeOrRangeOrPage Is ShapeRange _
   And Not TypeOf ShapeOrRangeOrPage Is Page Then
        Err.Raise 13, Source:="LesserDim", _
                  Description:="Type mismatch: ShapeOrRangeOrPage ������ ���� Shape, ShapeRange ��� Page"
        Exit Property
    End If
    If ShapeOrRangeOrPage.SizeWidth < ShapeOrRangeOrPage.SizeHeight Then
        LesserDim = ShapeOrRangeOrPage.SizeWidth
    Else
        LesserDim = ShapeOrRangeOrPage.SizeHeight
    End If
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

Public Property Get ShapeIsInGroup(ByVal Shape As Shape) As Boolean
    On Error GoTo Fail
    ShapeIsInGroup = Not (Shape.ParentGroup Is Nothing)
Fail:
End Property

'���������� ��������� ����, �� ������� ����� ����� �� ������
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

'���������� Rect, ������ ��������� ������� ���� Space �� ���� ������
Public Property Get SpaceBox( _
                    ByVal MaybeShapeOrRange As Variant, _
                    ByVal Space As Double _
                ) As Rect
    If Not (IsShape(MaybeShapeOrRange) Or IsShapeRange(MaybeShapeOrRange)) _
        Then Exit Property
    Set SpaceBox = MaybeShapeOrRange.BoundingBox.GetCopy
    SpaceBox.Inflate Space, Space, Space, Space
End Property

'���������� Outline ���� ��� Empty
Public Property Get TryGetOutline(ByVal Shape As Shape) As Variant
    If ShapeHasOutline(Shape) Then Set TryGetOutline = Shape.Outline
End Property

'===============================================================================
' # ������� ����������� � ��������� ������

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
               ByVal RalativePage As Page, _
               ByVal ReferencePoint As cdrReferencePoint _
           )
    Select Case ReferencePoint
        Case cdrTopRight
            ShapesBeingAligned.TopY = RalativePage.TopY
            ShapesBeingAligned.RightX = RalativePage.RightX
        Case cdrTopMiddle
            ShapesBeingAligned.TopY = RalativePage.TopY
            ShapesBeingAligned.CenterX = RalativePage.CenterX
        Case cdrTopLeft
            ShapesBeingAligned.TopY = RalativePage.TopY
            ShapesBeingAligned.LeftX = RalativePage.LeftX
        Case cdrMiddleLeft
            ShapesBeingAligned.CenterY = RalativePage.CenterY
            ShapesBeingAligned.LeftX = RalativePage.LeftX
        Case cdrBottomLeft
            ShapesBeingAligned.BottomY = RalativePage.BottomY
            ShapesBeingAligned.LeftX = RalativePage.LeftX
        Case cdrBottomMiddle
            ShapesBeingAligned.BottomY = RalativePage.BottomY
            ShapesBeingAligned.CenterX = RalativePage.CenterX
        Case cdrBottomRight
            ShapesBeingAligned.BottomY = RalativePage.BottomY
            ShapesBeingAligned.RightX = RalativePage.RightX
        Case cdrMiddleRight
            ShapesBeingAligned.CenterY = RalativePage.CenterY
            ShapesBeingAligned.RightX = RalativePage.RightX
        Case cdrCenter
            ShapesBeingAligned.CenterY = RalativePage.CenterY
            ShapesBeingAligned.CenterX = RalativePage.CenterX
    End Select
End Sub

Public Property Get HasSize(ByRef MaybeSome As Variant) As Boolean
    If Not AssignedObject(MaybeSome) Then Exit Property
    If TypeOf MaybeSome Is Shape Then GoTo Success
    If TypeOf MaybeSome Is ShapeRange Then GoTo Success
    If TypeOf MaybeSome Is Page Then GoTo Success
    If TypeOf MaybeSome Is Rect Then GoTo Success
    Exit Property
Success:
    HasSize = True
End Property

Public Property Get HasPosition(ByRef MaybeSome As Variant) As Boolean
    If Not AssignedObject(MaybeSome) Then Exit Property
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
    'RemainingShape � ExtractedShape � Extract
    '�� ����� ���� ��������, ��� � ������
    Set RemainingShape = Shape.Curve.SubPaths.First.Extract(ExtractedShape)
    BreakApart.Add ExtractedShape
    BreakApart.AddRange BreakApart(RemainingShape)
End Function

'������������� ������ � ������ ��� ����� � ����� �����,
'� ����������� �� ��������� �����
'����������
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

'��������� �������� Shape ��� ShapeRange �� ������ ����
Public Function CopyToLayer( _
                    ByVal ShapeOrRange As Object, _
                    ByVal Layer As Layer _
                ) As Object

    If Not TypeOf ShapeOrRange Is Shape And Not TypeOf ShapeOrRange Is ShapeRange Then
        Err.Raise 13, Source:="CopyToLayer", _
                  Description:="Type mismatch: ShapeOrRange ������ ���� Shape ��� ShapeRange"
        Exit Function
    End If
    
    Set CopyToLayer = ShapeOrRange.Duplicate
    MoveToLayer CopyToLayer, Layer

End Function

'���������� Boundary
Public Function CreateBoundary(ByVal ShapeOrRange As Object) As Shape
    On Error GoTo Catch
    Dim tShape As Shape, tRange As ShapeRange
    '������ ������ �� ���, ���� ���������� ���
    If TypeOf ShapeOrRange Is Shape Then
        Set tShape = ShapeOrRange
        Set CreateBoundary = tShape.CustomCommand("Boundary", "CreateBoundary")
    ElseIf TypeOf ShapeOrRange Is ShapeRange Then
        Set tRange = ShapeOrRange
        Set CreateBoundary = tRange.CustomCommand("Boundary", "CreateBoundary")
    Else
        Err.Raise 13, Source:="CreateBoundary", _
            Description:="Type mismatch: ShapeOrRange ������ ���� Shape ��� ShapeRange"
        Exit Function
    End If
    Exit Function
Catch:
    Debug.Print Err.Number
End Function

'������ ����, ���� ����� ���� ���� - ���������� ���� ����
Public Function CreateOrFindLayer( _
                    ByVal Page As Page, _
                    ByVal Name As String _
                ) As Layer
    Set CreateOrFindLayer = Page.Layers.Find(Name)
    If CreateOrFindLayer Is Nothing Then
        Set CreateOrFindLayer = Page.CreateLayer(Name)
    End If
End Function

'���������� Crop Tool
Public Function CropTool( _
                    ByVal ShapeOrRangeOrPage As Object, _
                    ByVal x1#, ByVal y1#, _
                    ByVal x2#, ByVal y2#, _
                    Optional ByVal Angle = 0 _
                ) As ShapeRange
    If TypeOf ShapeOrRangeOrPage Is Shape Or _
         TypeOf ShapeOrRangeOrPage Is ShapeRange Or _
         TypeOf ShapeOrRangeOrPage Is Page Then
        Set CropTool = ShapeOrRangeOrPage.CustomCommand("Crop", "CropRectArea", x1, y1, x2, y2, Angle)
    Else
        Err.Raise 13, Source:="CropTool", _
            Description:="Type mismatch: ShapeOrRangeOrPage ������ ���� Shape, ShapeRange ��� Page"
        Exit Function
    End If
End Function

'�������� ����� �� Shape �� ������� Knife, ���������� ���������� �����
Public Function Dissect(ByRef Shape As Shape, ByRef Knife As Shape) As Shape
    Set Dissect = Intersect(Knife, Shape, True, True)
    Set Shape = Knife.Trim(Shape, True, False)
End Function

'����������� �������� �������� �� ����� ������ � ���������
Public Function DuplicateActivePage( _
                    ByVal NumberOfPages As Long, _
                    Optional ByVal ExcludeLayerName As String = "" _
                ) As Page
    Dim tRange As ShapeRange
    Dim tShape As Shape, sDuplicate As Shape
    Dim tProps As typeLayerProps
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

'��� ������� �� ���� ���������, ������� ������-�������� - �� ���� ����
'��� �������� �����������, ��� ������� �� ����� guides �����������
Public Function FlattenPagesToLayer(ByVal LayerName As String) As Layer

    Dim DL As Layer: Set DL = ActiveDocument.MasterPage.DesktopLayer
    Dim DLstate As Boolean: DLstate = DL.Editable
    Dim P As Page
    Dim L As Layer
    
    DL.Editable = False
    
    For Each P In ActiveDocument.Pages
        For Each L In P.Layers
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
        If P.Index <> 1 Then P.Delete
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

'���������� ���������
Public Function Intersect( _
                    ByVal SourceShape As Shape, _
                    ByVal TargetShape As Shape, _
                    Optional ByVal LeaveSource As Boolean = True, _
                    Optional ByVal LeaveTarget As Boolean = True _
                ) As Shape
                                     
    Dim tPropsSource As typeLayerProps
    Dim tPropsTarget As typeLayerProps
    
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

'���������� Join Curves
Public Function JoinCurves(ByVal ShapeOrShapes As Variant, ByVal Tolerance As Double)
    ShapeOrShapes.CustomCommand "ConvertTo", "JoinCurves", Tolerance
End Function

'�� �������� � �����������
Public Sub MatrixCopy(ByVal SourceShape As Shape, ByVal TargetShape As Shape)
    Dim tMatrix As typeMatrix
    With tMatrix
        SourceShape.GetMatrix .d11, .d12, .d21, .d22, .tx, .ty
        TargetShape.SetMatrix .d11, .d12, .d21, .d22, .tx, .ty
    End With
End Sub

'��������� ���������� Shape ��� ShapeRange �� ������ ����
Public Function MoveToLayer( _
                    ByVal MaybeShapeOrRange As Variant, _
                    ByVal MaybeLayer As Layer _
                )
    If VBA.IsEmpty(MaybeShapeOrRange) _
    Or VBA.IsEmpty(MaybeLayer) Then Exit Function
    ThrowIfNotShapeOrRange MaybeShapeOrRange
    
    Dim tSrcLayer() As Layer
    Dim tProps() As typeLayerProps
    Dim tLayersCol As Collection
    Dim i&
    
    If TypeOf MaybeShapeOrRange Is Shape Then
    
        Set tLayersCol = New Collection
        tLayersCol.Add MaybeShapeOrRange.Layer
        
    ElseIf TypeOf MaybeShapeOrRange Is ShapeRange Then
        
        If MaybeShapeOrRange.Count < 1 Then Exit Function
        Set tLayersCol = ShapeRangeLayers(MaybeShapeOrRange)
        
    Else
    
        Throw "Type mismatch: MaybeShapeOrRange ������ ���� Shape ��� ShapeRange"
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

'�������� ��������
'�����: Alex Vakulenko http://www.oberonplace.com/vba/drawmacros/delsegment.htm
Public Sub SegmentDelete(ByVal Segment As Segment)
    If Not Segment.EndNode.IsEnding Then
        Segment.EndNode.BreakApart
        Set Segment = Segment.SubPath.LastSegment
    End If
    Segment.EndNode.Delete
End Sub

'��������� ���� ������ ������
Public Sub SetOutlineColor( _
               ByVal Shapes As ShapeRange, _
               ByVal Color As Color _
           )
    Dim Shape As Shape
    For Each Shape In Shapes
        Shape.Outline.Color.CopyAssign Color
    Next Shape
End Sub

Public Sub Trim( _
               ByVal TrimmerShape As Shape, _
               ByRef TargetShape As Shape _
           )
    Set TargetShape = TrimmerShape.Trim(TargetShape)
End Sub

'�������� ������ �� CropEnvelopeShape, �� ��-������,
'������� ������� �� EXPANDBY �������� ��������
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
' # ������� ������ � �������

Public Function AddProperEndingToPath(ByVal Path As String) As String
    If Not VBA.Right$(Path, 1) = "\" Then AddProperEndingToPath = Path & "\" _
    Else: AddProperEndingToPath = Path
End Function

'���������� �� ���� ��� ����� (����� ������ ������������� �� "\")
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

'���������� ��� ����� ��� ����������
Public Property Get GetFileNameNoExt(ByVal FileName As String) As String
    If VBA.Right(FileName, 1) <> "\" And VBA.Len(FileName) > 0 Then
        GetFileNameNoExt = Left(FileName, _
            Switch _
                (InStr(FileName, ".") = 0, _
                    Len(FileName), _
                InStr(FileName, ".") > 0, _
                    InStrRev(FileName, ".") - 1))
    End If
End Property

'������ ��� ���������� �����
Public Property Get GetTempFile() As String
    GetTempFile = GetTempFolder & GetTempFileName
End Property

'��� ���������� �����
Public Property Get GetTempFileName() As String
    GetTempFileName = "elvin_" & CreateGUID & ".tmp"
End Property

'������� ��������� �����
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

'������ �����, ���� �� ����
'���������� Path ������� (��� inline-�������������)
Public Function MakeDir(ByVal Path As String) As String
    If VBA.Dir(Path, vbDirectory) = "" Then MkDir Path
    MakeDir = Path
End Function

'��������� ���� � ������
Public Function ReadFile(ByVal File As String) As String
    Dim tFileNum As Long
    tFileNum = FreeFile
    Open File For Input As #tFileNum
    ReadFile = Input(LOF(tFileNum), tFileNum)
    Close #tFileNum
End Function

'��������� ���� � ������ ����� ADODB, ����� ������ ���������
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

'�������� ���������� ����� �� ��������
Public Function SetFileExt( _
                    ByVal SourceFile As String, _
                    ByVal NewExt As String _
                ) As String
    If Right(SourceFile, 1) <> "\" And Len(SourceFile) > 0 Then
        SetFileExt = GetFileNameNoExt(SourceFile$) & "." & NewExt
    End If
End Function

'��������� ������ Content � ����, �������������, ����� � �������� temp ����,
'� �������� �����, ���� ����������
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

'��������� ������ Content � ���� ����� ADODB, ����� ������ ���������
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
' # ������ �������

Public Sub AppendCollection( _
               ByVal Destination As Collection, _
               ByVal SourceToAdd As Collection _
           )
    Dim Item As Variant
    For Each Item In SourceToAdd
        Destination.Add Item
    Next Item
End Sub

Public Sub Assign(ByRef Destination As Variant, ByRef Value As Variant)
    If VBA.IsObject(Value) Then
        Set Destination = Value
    Else
        Destination = Value
    End If
End Sub

'���������� True, ���� Value - ��� ������ � ��� ���� �� Nothing
Public Property Get AssignedObject(ByRef Variable As Variant) As Boolean
    If Not VBA.IsObject(Variable) Then Exit Property
    AssignedObject = Not Variable Is Nothing
End Property

'-------------------------------------------------------------------------------
' �������           : BoostStart, BoostFinish
' ������            : 2022.05.31
' ������            : dizzy, elvin-nsk
' ����������        : ������������ ������������ �� CtC
' �����������       : ���������������
'
' ���������:
' ~~~~~~~~~~
'
'
' �������������:
' ~~~~~~~~~~~~~~
'
'-------------------------------------------------------------------------------
Public Sub BoostStart( _
               Optional ByVal UndoGroupName As String = "", _
               Optional ByVal Optimize As Boolean = True _
           )
    If Not UndoGroupName = "" And Not ActiveDocument Is Nothing Then _
        ActiveDocument.BeginCommandGroup UndoGroupName
    If Optimize And Not Optimization Then Optimization = True
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

Public Property Get Count(ByRef Arr As Variant) As Long
    Count = UBound(Arr) - LBound(Arr) + 1
End Property

Public Property Get CountCopy(ByVal Arr As Variant) As Long
    CountCopy = UBound(Arr) - LBound(Arr) + 1
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

Public Property Get FindMaxItemNum(ByVal Collection As Collection) As Long
    FindMaxItemNum = 1
    Dim i As Long
    For i = 1 To Collection.Count
        If VBA.IsNumeric(Collection(i)) Then
            If Collection(i) > Collection(FindMaxItemNum) Then _
                FindMaxItemNum = i
        End If
    Next i
End Property

Public Property Get FindMinItemNum(ByVal Collection As Collection) As Long
    FindMinItemNum = 1
    Dim i As Long
    For i = 1 To Collection.Count
        If VBA.IsNumeric(Collection(i)) Then
            If Collection(i) < Collection(FindMinItemNum) Then _
                FindMinItemNum = i
        End If
    Next i
End Property

'�������� �� ����� ������ :) ��� ����� Even � Odd ���������� ����...
Public Property Get IsChet(ByVal X As Variant) As Boolean
    If X Mod 2 = 0 Then IsChet = True Else IsChet = False
End Property

'������� �� Number �� Divider ������
Public Property Get IsDivider( _
                    ByVal Number As Long, _
                    ByVal Divider As Long _
                ) As Boolean
    If Number Mod Divider = 0 Then IsDivider = True Else IsDivider = False
End Property

Public Property Get IsLowerCase(ByVal Str As String) As Boolean
    If VBA.LCase(Str) = Str Then IsLowerCase = True
End Property

Public Property Get IsSame( _
                        ByRef Value1 As Variant, _
                        ByRef Value2 As Variant _
                    ) As Boolean
    If VBA.IsObject(Value1) And VBA.IsObject(Value2) Then
        IsSame = Value1 Is Value2
    ElseIf Not VBA.IsObject(Value1) And Not VBA.IsObject(Value2) Then
        IsSame = (Value1 = Value2)
    End If
End Property

Public Property Get IsUpperCase(ByVal Str As String) As Boolean
    If VBA.UCase(Str) = Str Then IsUpperCase = True
End Property

Public Property Get IsVoid(ByRef Some As Variant) As Boolean
    If VBA.IsNull(Some) _
    Or VBA.IsEmpty(Some) _
    Or VBA.IsMissing(Some) Then
        IsVoid = True
        Exit Property
    End If
    If VBA.IsObject(Some) Then
        If Some Is Nothing Then
            IsVoid = True
            Exit Property
        End If
    End If
End Property

Public Sub RemoveElementFromCollection( _
               ByVal Collection As Collection, _
               ByVal Element As Variant _
           )
    If Collection.Count = 0 Then Exit Sub
    Dim i As Long
    For i = 1 To Collection.Count
        If IsSame(Element, Collection(i)) Then
            Collection.Remove i
            Exit Sub
        End If
    Next i
End Sub

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
    Debug.Print Message & VBA.CStr(Round(Timer - StartTime, 3)) & " ������"
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

Public Property Get Contains( _
                        ByRef Sequence As Variant, _
                        ByRef Items As Variant _
                    ) As Boolean
    Dim Element As Variant
    Dim Item As Variant
    Dim ItemExists As Boolean
    For Each Item In Items
        For Each Element In Sequence
            If IsSame(Item, Element) Then
                ItemExists = True
                Exit For
            End If
        Next Element
        If Not ItemExists Then Exit Property
        ItemExists = False
    Next Item
    Contains = True
End Property

Public Property Get Pack(ParamArray Items() As Variant) As Variant()
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

'������ ShapeRange �� Shape/Shapes/ShapeRange
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
            Throw "�� �������� ������"
        End If
    Next Item
End Function

Private Sub Resize(ByRef Arr As Variant, ByVal Length As Long)
    ReDim Preserve Arr(LBound(Arr) To LBound(Arr) + Length - 1)
End Sub

'��������� ����� �� LowerBound �� UpperBound
Public Property Get RndInt( _
                        ByVal LowerBound As Long, _
                        ByVal UpperBound As Long _
                    ) As Long
    RndInt = VBA.Int((UpperBound - LowerBound + 1) * VBA.Rnd + LowerBound)
End Property

'������� ���������� � ���������� / � �������� � ���� immediate
Public Sub Show(ByRef Variable As Variant)
    If VBA.IsObject(Variable) Then
        If Variable Is Nothing Then
            ShowString "[Nothing]"
        Else
            'TODO ������� ����� ��������
            ShowString "[Object]: " & VBA.TypeName(Variable)
        End If
    Else
        If VBA.IsMissing(Variable) Then
            ShowString "[Missing]"
        Else
            Select Case VBA.VarType(Variable)
                Case vbEmpty
                    ShowString "[Empty]"
                Case vbNull
                    ShowString "[Null]"
                Case vbError
                    ShowString "[Error]"
                Case vbArray
                Case vbString
                    ShowString Variable
                Case Else
                    ShowString "[Value]: " & Variable
            End Select
            
        End If
    End If
End Sub

Public Sub Throw(Optional ByVal Message As String = "����������� ������")
    VBA.Err.Raise CustomError, , Message
End Sub

'===============================================================================
' # ��������� ������� ������

'��� IsOverlap
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

Private Sub LayerPropsPreserve(ByVal L As Layer, ByRef Props As typeLayerProps)
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
Private Sub LayerPropsRestore(ByVal L As Layer, ByRef Props As typeLayerProps)
    With Props
        If L.Visible <> .Visible Then L.Visible = .Visible
        If L.Printable <> .Printable Then L.Printable = .Printable
        If L.Editable <> .Editable Then L.Editable = .Editable
    End With
End Sub
Private Sub LayerPropsPreserveAndReset( _
                ByVal L As Layer, _
                ByRef Props As typeLayerProps _
            )
    LayerPropsPreserve L, Props
    LayerPropsReset L
End Sub

Private Sub ShowString(ByVal Str As String)
    Debug.Print Str
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
    Throw "��� ������ ���� Shape ��� ShapeRange"
End Sub

Private Sub ThrowIfNotCollectionOrArray(ByRef CollectionOrArray As Variant)
    If VBA.IsObject(CollectionOrArray) Then _
        If TypeOf CollectionOrArray Is Collection Then Exit Sub
    If VBA.IsArray(CollectionOrArray) Then Exit Sub
    VBA.Err.Raise 13, Source:="lib_elvin", _
                  Description:="Type mismatch: CollectionOrArray ������ ���� Collection ��� Array"
End Sub

'===============================================================================
' # ����-����� � ������ ����� ������

Private Sub TestShow()
    Show Empty
    Show Null
    Show 3
    Show "3"
    Show New Collection
End Sub

Private Sub UnitContains()
    Debug.Assert Contains(Array(1, 2, 3), Array(3, 1, 2)) = True
    Debug.Assert Contains(Array(1, 2, 3), Array(3, 1, 4)) = False
    Debug.Print "Contains is OK"
End Sub

Private Sub UnitHasPosition()
    Debug.Assert HasPosition(CreateRect) = True
    Debug.Assert HasPosition(New Collection) = False
    Debug.Assert HasPosition(123) = False
    Debug.Print "HasPosition is OK"
End Sub

Private Sub UnitHasSize()
    Debug.Assert HasSize(CreateRect) = True
    Debug.Assert HasPosition(New Collection) = False
    Debug.Assert HasPosition(123) = False
    Debug.Print "HasSize is OK"
End Sub

Private Sub UnitIsJust()
    Debug.Assert IsJust(0) = True
    Debug.Assert IsJust(1) = True
    Debug.Assert IsJust(New Collection) = True
    Debug.Assert IsJust(Empty) = False
    Debug.Assert IsJust(Null) = False
    Debug.Assert IsJust(Nothing) = False
    Debug.Assert IsJust(VBA.CVErr(ErrorCodes.ErrorInvalidArgument)) = False
    Debug.Print "IsJust is OK"
End Sub
