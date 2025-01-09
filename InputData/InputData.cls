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
'===============================================================================
'   Модуль          : InputData
'   Версия          : 2025.01.09
'   Автор           : elvin-nsk (me@elvin.nsk.ru)
'   Сайт            : https://github.com/elvin-nsk/LowCoupledFromCore
'   Назначение      : Получение исходного пользовательского выбора
'   Зависимости     : нет
'===============================================================================

'@ModuleDescription("Получение исходного пользовательского выбора.")
'@PredeclaredId
Option Explicit

'===============================================================================

Private Const DefaultErrNoDocument As String = "Нет активного документа."
Private Const DefaultErrLayerDisabled As String = _
    "Текущий слой закрыт."
Private Const DefaultErrNoSelection As String = "Выделите объекты."
Private Const DefaultErrPageEmpty As String = "Страница пустая."

Private Type This
    Document As Document
    ExpectedObject As Object
    Layer As Layer
    Ok As Boolean
    Page As Page
    Shape As Shape
    Shapes As ShapeRange
End Type
Private This As This

'===============================================================================

Public Function ExpectDocument( _
                    Optional ByVal ErrNoDocument As String = _
                        DefaultErrNoDocument _
                ) As InputData
    Set ExpectDocument = New InputData
    If ActiveDocument Is Nothing Then
        VBA.MsgBox ErrNoDocument, vbExclamation
        Exit Function
    End If
    With ExpectDocument
        .InjectOk True
        .InjectDocument ActiveDocument
        .InjectPage .Document.ActivePage
        .InjectLayer GetFirstNonSpecialOrActiveLayer(.Page)
        .InjectShapes .Page.Shapes.All
        .InjectShape .Page.Shapes.First
        .InjectExpected .Document
    End With
End Function

Public Function ExpectLayer( _
                    Optional ByVal ErrNoDocument As String = _
                        DefaultErrNoDocument, _
                    Optional ByVal LayerMustBeEnabled As Boolean = False, _
                    Optional ByVal ErrLayerDisabled As String = _
                        DefaultErrLayerDisabled _
                ) As InputData
    Set ExpectLayer = New InputData
    If ActiveDocument Is Nothing Then
        VBA.MsgBox ErrNoDocument, vbExclamation
        Exit Function
    End If
    If Not LayerIsOk(ActiveLayer, LayerMustBeEnabled) Then
        VBA.MsgBox ErrLayerDisabled, vbExclamation
        Exit Function
    End If
    With ExpectLayer
        .InjectOk True
        .InjectLayer ActiveLayer
        .InjectPage .Layer.Page
        .InjectDocument .Page.Parent.Parent
        .InjectShapes .Layer.Shapes.All
        .InjectShape .Shapes.FirstShape
        .InjectExpected .Layer
    End With
End Function

Public Function ExpectPage( _
                    Optional ByVal ErrNoDocument As String = _
                        DefaultErrNoDocument, _
                    Optional ByVal PageCanBeEmpty As Boolean = False, _
                    Optional ByVal ErrPageEmpty As String = _
                        DefaultErrPageEmpty _
                ) As InputData
    Set ExpectPage = New InputData
    If ActiveDocument Is Nothing Then
        VBA.MsgBox ErrNoDocument, vbExclamation
        Exit Function
    End If
    If (ActiveDocument.ActivePage.Shapes.Count = 0) And Not PageCanBeEmpty Then
        VBA.MsgBox ErrPageEmpty, vbExclamation
        Exit Function
    End If
    With ExpectPage
        .InjectOk True
        .InjectDocument ActiveDocument
        .InjectPage .Document.ActivePage
        .InjectLayer GetFirstNonSpecialOrActiveLayer(.Page)
        .InjectShapes .Page.Shapes.All
        .InjectShape .Page.Shapes.First
        .InjectExpected .Page
    End With
End Function

Public Function ExpectShapes( _
                    Optional ByVal ErrNoDocument As String = _
                        DefaultErrNoDocument, _
                    Optional ByVal LayerMustBeEnabled As Boolean = False, _
                    Optional ByVal ErrLayerDisabled As String = _
                        DefaultErrLayerDisabled, _
                    Optional ByVal ErrNoSelection As String = _
                        DefaultErrNoSelection _
                ) As InputData
    Set ExpectShapes = New InputData
    If ActiveDocument Is Nothing Then
        VBA.MsgBox ErrNoDocument, vbExclamation
        Exit Function
    End If
    If ActiveSelectionRange.Count = 0 Then
        VBA.MsgBox ErrNoSelection, vbInformation
        Exit Function
    End If
    With ExpectShapes
        .InjectDocument ActiveDocument
        .InjectShapes ActiveSelectionRange
        .InjectShape .Shapes.FirstShape
        .InjectPage .Shape.Page
        .InjectLayer .Shape.Layer
        .InjectOk LayerIsOk(ActiveLayer, LayerMustBeEnabled)
        If Not .Ok Then
            VBA.MsgBox ErrLayerDisabled, vbExclamation
        End If
        .InjectExpected .Shapes
    End With
End Function

Friend Sub InjectDocument(ByVal Document As Document)
    Set This.Document = Document
End Sub

Friend Sub InjectExpected(ByVal ExpectedObject As Object)
    Set This.ExpectedObject = ExpectedObject
End Sub

Friend Sub InjectOk(ByVal Ok As Boolean)
    This.Ok = Ok
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

Public Property Get Fail() As Boolean
    Fail = Not This.Ok
End Property

Public Property Get Layer() As Layer
    Set Layer = This.Layer
End Property

Public Function Ok(Optional ByRef ExpectedObject As Object) As Boolean
    Ok = This.Ok
    Set ExpectedObject = This.ExpectedObject
End Function

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
    LayerEnabled = Layer.Visible And Layer.Printable And Layer.Editable
End Function
