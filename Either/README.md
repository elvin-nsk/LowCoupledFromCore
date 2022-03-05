# Either

Украл из языка Kotlin, понравилась концепция, описанная [здесь](https://habr.com/ru/company/piter/blog/589579/).

Как-то прижилось.

Конструкция для красивой работы с методами, которые могут выполняться с предусмотренной ошибкой. Возвращает либо ошибку, либо успешный результат.

## Использование

Если инициализировать с пустыми аргументами - `Either` примет статус "ошибка". Также можно передать в качестве ошибки некоторый объект (например, текст ошибки) - тогда надо инициализировать с пустым первым аргументом, а во втором передать "ошибочный" объект. В первом аргументе всегда передаётся верный объект. Если первый аргумент присутствует, то `Either` примет статус "успешно".

### Конструктор:
```VBA
Either.Create(Optional SuccessValue As Variant, Optional ErrorValue As Variant)
```

`SuccessValue`  - передаваемый объект в случае успеха

`ErrorValue`  - передаваемый объект в случае ошибки

### Поля:

`IsError` - статус "ошибка"
`IsSuccess` - статус "успешно"
`ErrorValue` - объект в случае ошибки
`SuccessValue` - объект в случае успеха

## Пример

```VBA
' возвращаем текущий документ
' либо, если нет открытых документов - возвращаем Either со статусом Error
Function GetDocument() as IEither
  If ActiveDocument Is Nothing Then
    Set GetDocument = Either.Create
    Exit Function
  End If
  Set GetDocument = Either.Create(ActiveDocument)
End Function

' запрашиваем активный документ в главной процедуре
' если его нет - выдаём ошибку и завершаем
Sub Main()
  Dim MyDocument as Document
  With GetDocument
    If .IsError Then
      VBA.MsgBox "Нет открытых документов"
      Exit Sub
    End If
    Set MyDocument = .SuccessValue
  End With
End Sub
```

## Зависимости

Нет.