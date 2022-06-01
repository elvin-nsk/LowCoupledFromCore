VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InputData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Получение исходного пользовательского выбора."
'@ModuleDescription("Получение исходного пользовательского выбора.")

'@PredeclaredId
Option Explicit

'===============================================================================

Private Const DefaultErrMsgNoDocument As String = "Нет активного документа"
Private Const DefaultErrMsgLayerDisabled As String = _
    "Откройте текущий слой для редактирования"
Private Const DefaultErrErrNoSelection As String = "Выделите объекты"

Private Type typeThis
    Document As Document
    IsSuccess As Boolean
    Layer As Layer
    Page As Page
    Shape As Shape
    Shapes As ShapeRange
End Type
Private This As typeThis

'===============================================================================

Public Function GetDocumentOrPage( _
                    Optional ByVal ErrMsgNoDocument As String = _
                        DefaultErrMsgNoDocument _
                ) As InputData
    Set GetDocumentOrPage = New InputData
    With GetDocumentOrPage
        .InjectSuccess Not ActiveDocument Is Nothing
        If .IsSuccess Then
            .InjectDocument ActiveDocument
            .InjectPage .Document.ActivePage
            .InjectLayer GetFirstNonSpecialOrActiveLayer(.Page)
            .InjectShapes .Page.Shapes.All
            .InjectShape .Page.Shapes.First
        Else
            VBA.MsgBox ErrMsgNoDocument, vbExclamation
        End If
    End With
End Function

Public Function GetLayer( _
                    Optional ByVal ErrMsgNoDocument As String = _
                        DefaultErrMsgNoDocument, _
                    Optional ByVal LayerMustBeEnabled As Boolean = False, _
                    Optional ByVal ErrMsgLayerDisabled As String = _
                        DefaultErrMsgLayerDisabled _
                ) As InputData
    Set GetLayer = New InputData
    If ActiveDocument Is Nothing Then
        VBA.MsgBox ErrMsgNoDocument, vbExclamation
        Exit Function
    End If
    If Not LayerIsOk(ActiveLayer, LayerMustBeEnabled) Then
        VBA.MsgBox ErrMsgLayerDisabled, vbExclamation
        Exit Function
    End If
    With GetLayer
        .InjectSuccess True
        .InjectLayer ActiveLayer
        .InjectPage .Layer.Page
        .InjectDocument .Page.Parent.Parent
        .InjectShapes .Layer.Shapes.All
        .InjectShape .Shapes.FirstShape
    End With
End Function

Public Function GetShapes( _
                    Optional ByVal ErrMsgNoDocument As String = _
                        DefaultErrMsgNoDocument, _
                    Optional ByVal LayerMustBeEnabled As Boolean = False, _
                    Optional ByVal ErrMsgLayerDisabled As String = _
                        DefaultErrMsgLayerDisabled, _
                    Optional ByVal ErrNoSelection As String = _
                        DefaultErrErrNoSelection _
                ) As InputData
    Set GetShapes = New InputData
    If ActiveDocument Is Nothing Then
        VBA.MsgBox ErrMsgNoDocument, vbExclamation
        Exit Function
    End If
    If ActiveSelectionRange.Count = 0 Then
        VBA.MsgBox ErrNoSelection, vbInformation
        Exit Function
    End If
    With GetShapes
        .InjectDocument ActiveDocument
        .InjectShapes ActiveSelectionRange
        .InjectShape .Shapes.FirstShape
        .InjectPage .Shape.Page
        .InjectLayer .Shape.Layer
        .InjectSuccess LayerIsOk(ActiveLayer, LayerMustBeEnabled)
        If Not .IsSuccess Then
            VBA.MsgBox ErrMsgLayerDisabled, vbExclamation
        End If
    End With
End Function

Friend Sub InjectDocument(ByVal Document As Document)
    Set This.Document = Document
End Sub

Friend Sub InjectSuccess(ByVal Success As Boolean)
    This.IsSuccess = Success
End Sub

Friend Sub InjectLayer(ByVal Layer As Layer)
    Set This.Layer = Layer
End Sub

Friend Sub InjectPage(ByVal Page As Page)
    Set This.Page = Page
End Sub

Friend Sub InjectShape(ByVal Shape As Shape)
    Set This.Shape = Shape
End Sub

Friend Sub InjectShapes(ByVal Shapes As ShapeRange)
    Set This.Shapes = Shapes
End Sub

'===============================================================================

Public Property Get Document() As Document
    Set Document = This.Document
End Property

Public Property Get IsError() As Boolean
    IsError = Not This.IsSuccess
End Property

Public Property Get IsSuccess() As Boolean
    IsSuccess = This.IsSuccess
End Property

Public Property Get Layer() As Layer
    Set Layer = This.Layer
End Property

Public Property Get Page() As Page
    Set Page = This.Page
End Property

Public Property Get Shape() As Shape
    Set Shape = This.Shape
End Property

Public Property Get Shapes() As ShapeRange
    Set Shapes = This.Shapes
End Property

Public Property Get Self() As InputData
    Set Self = Me
End Property

'===============================================================================

Private Function GetFirstNonSpecialOrActiveLayer(ByVal Page As Page) As Layer
    Dim Layer As Layer
    For Each Layer In Page.Layers
        If Not Layer.IsSpecialLayer Then
            Set GetFirstNonSpecialOrActiveLayer = Layer
            Exit Function
        End If
    Next Layer
    Set GetFirstNonSpecialOrActiveLayer = Page.ActiveLayer
End Function

Private Function LayerIsOk( _
                     ByVal Layer As Layer, _
                     ByVal LayerMustBeEnabled As Boolean _
                 ) As Boolean
    LayerIsOk = (LayerMustBeEnabled And LayerEnabled(Layer)) _
         Or Not LayerMustBeEnabled
End Function

Private Function LayerEnabled(ByVal Layer As Layer) As Boolean
    LayerEnabled = Layer.Editable And Layer.Visible
End Function
