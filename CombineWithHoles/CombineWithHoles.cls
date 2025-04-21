VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CombineWithHoles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'===============================================================================
'   Модуль          : CombineWithHoles
'   Версия          : 2025.04.21
'   Автор           : elvin-nsk (me@elvin.nsk.ru)
'   Сайт            : https://github.com/elvin-nsk/LowCoupledFromCore
'   Назначение      : Аналог иллюстраторовского Pathfinder - Exclude
'   Зависимости     : LibCore
'===============================================================================

Option Explicit

'===============================================================================
' # Declarations

Private Type This
    Shapes As ShapeRange
    ConvertToCurves As Boolean
    BreakApart As Boolean
End Type
Private This As This

'===============================================================================
' # Constructors

Friend Sub Inject( _
               ByVal Shapes As ShapeRange, _
               Optional ConvertToCurves As Boolean = True, _
               Optional BreakApart As Boolean = True _
           )
    With This
        Set .Shapes = Shapes
        .ConvertToCurves = ConvertToCurves
        .BreakApart = BreakApart
    End With
End Sub

'===============================================================================
' # Logic

Public Function Run(Optional ByVal SmallestFirst As Boolean = False) As ShapeRange
    If This.BreakApart Or This.ConvertToCurves Then PreProcess
    
    Dim Pool As New ShapeRange: Pool.AddRange This.Shapes
    If SmallestFirst Then
        Pool.Sort _
            "@Shape1.Width + @Shape1.Height < @Shape2.Width + @Shape2.Height"
    Else
        Pool.Sort _
            "@Shape1.Width + @Shape1.Height > @Shape2.Width + @Shape2.Height"
    End If
    
    #If DEV Then
    'Pool.ApplyUniformFill Red
    'Pool(1).Fill.ApplyUniformFill Red
    #End If
    
    Set Run = New ShapeRange
    Dim CurrentMatch As ShapeRange
    Do While Pool.Count > 0
        If SmallestFirst Then
            Set CurrentMatch = SmallestOverlaps(Pool)
        Else
            Set CurrentMatch = BiggestOverlaps(Pool)
        End If
        Pool.RemoveRange CurrentMatch
        Run.Add CurrentMatch.Combine
    Loop
End Function

'===============================================================================
' # Public

Public Function RunSmallestFirst() As ShapeRange
    Set RunSmallestFirst = Run(SmallestFirst:=True)
End Function

Public Function RunBiggestFirst() As ShapeRange
    Set RunBiggestFirst = Run(SmallestFirst:=False)
End Function

'===============================================================================
' # Helpers

Private Property Get SmallestOverlaps( _
                         ByVal SortedPool As ShapeRange _
                     ) As ShapeRange
    Set SmallestOverlaps = CreateShapeRange
    If SortedPool.Count = 0 Then Exit Property
    
    SmallestOverlaps.Add SortedPool(1)
    If SortedPool.Count = 1 Then Exit Property
    
    Dim FirstOverlapIndex As Long: FirstOverlapIndex = _
        FindFirstSmallerOverlapIndex(SortedPool(1), SortedPool)
    
    #If DEV Then
    'SortedPool(FirstOverlapIndex).Fill.ApplyUniformFill Red
    #End If
    
    If FirstOverlapIndex < 2 Then Exit Property
    SmallestOverlaps.Add SortedPool(FirstOverlapIndex)
    If FirstOverlapIndex = 2 Then Exit Property
    
    Dim i As Long
    For i = 2 To FirstOverlapIndex - 1
        If IsShapeNodesInsideOtherShape( _
            SortedPool(i), SortedPool(FirstOverlapIndex) _
        ) Then
            SmallestOverlaps.Add SortedPool(i)
        End If
    Next i
End Property

Private Property Get BiggestOverlaps( _
                         ByVal SortedPool As ShapeRange _
                     ) As ShapeRange
    Set BiggestOverlaps = CreateShapeRange
    If SortedPool.Count = 0 Then Exit Property
        
    BiggestOverlaps.Add SortedPool(1)
    If SortedPool.Count = 1 Then Exit Property
    
    Dim FirstOverlapIndex As Long: FirstOverlapIndex = _
        FindFirstBiggerOverlapIndex(SortedPool(1), SortedPool)
        
    If FirstOverlapIndex < 2 Then Exit Property
    BiggestOverlaps.Add SortedPool(FirstOverlapIndex)
    
    Dim i As Long
    For i = 2 To SortedPool.Count
        If IsShapeNodesInsideOtherShape( _
            SortedPool(i), SortedPool(1) _
        ) _
       And Not IsShapeNodesInsideOtherShape( _
            SortedPool(i), SortedPool(FirstOverlapIndex) _
        ) Then
            BiggestOverlaps.Add SortedPool(i)
        End If
    Next i
End Property

Private Property Get FindFirstSmallerOverlapIndex( _
                         ByVal Shape As Shape, _
                         ByVal Pool As ShapeRange _
                     ) As Long
    If Pool.Count = 0 Then Exit Property
    Dim i As Long
    For i = 1 To Pool.Count
        If IsShapeNodesInsideOtherShape(Shape, Pool(i)) Then
            FindFirstSmallerOverlapIndex = i
            Exit Property
        End If
    Next i
End Property

Private Property Get FindFirstBiggerOverlapIndex( _
                         ByVal Shape As Shape, _
                         ByVal Pool As ShapeRange _
                     ) As Long
    If Pool.Count = 0 Then Exit Property
    Dim i As Long
    For i = 1 To Pool.Count
        If IsShapeNodesInsideOtherShape(Pool(i), Shape) Then
            FindFirstBiggerOverlapIndex = i
            Exit Property
        End If
    Next i
End Property

Private Sub PreProcess()
    Dim Processed As New ShapeRange
    Dim Shape As Shape
    For Each Shape In This.Shapes
        If This.ConvertToCurves Then Shape.ConvertToCurves
        If This.BreakApart Then
            Processed.AddRange Shape.BreakApartEx
        Else
            Processed.Add Shape
        End If
    Next Shape
    Set This.Shapes = Processed
End Sub
