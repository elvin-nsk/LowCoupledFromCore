# Composer

Простой укладчик шейпов или шейпренджей по строкам. Можно ограничить длины строки по размеру или по количеству элементов, максимальную высоту в количестве строк или по размему, задать расстояние между элементами. Чешет от заданной точки вправо и вниз. Начинает работать сразу после инициализации.

## Использование

Композеру скармливается заранее подготовленная коллекция `ComposerElement`ов. `ComposerElement`ом может быть объект типа `Shape` или `ShapeRange`.

Размещённые элементы попадают в коллекцию `ComposedElements`, не вошедшие - `RemainingElements`.

Если один из параметров при инициализации равен нулю или не задан, то соответствующее ограничение не накладывается.

## Пример

```VBA
' добавляем активные шейпы в коллекцию
Dim ComposerElements As New Collection
Dim Shape As Shape
For Each Shape In ActiveSelectionRange
  ComposerElements.Add ComposerElement.Create(Shape)
Next Shape

' инициализируем и запускаем Composer с этой коллекцией:
With Composer.CreateAndCompose( _
                  Elements:=ComposerElements, _
                  StartingPoint:=FreePoint.Create(0, 297), _
                  MaxPlacesInWidth:=3, _
                  MaxPlacesInHeight:=4, _
                  MaxWidth:=0, _
                  MaxHeight:=297, _
                  HorizontalSpace:=0, _
                  VerticalSpace:=0 _
                )
End With
```

## Зависимости

[Point](Point).