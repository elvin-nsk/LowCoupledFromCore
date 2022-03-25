# Low-coupled from Core

Пока я в тайне от человечества пишу мега-проект Core - что-то вроде фреймворка для корела - иногда хочется иметь кусочки егойного функционала, чтобы сунуть в какой-нибудь проектик. Внутри Core всё высокосвязное, поэтому так просто не оторвёшь. Поэтому начал вытаскивать для себя переделанные кусочки с пониженной <s>социальной ответственностью</s> связанностью. Пусть лежит и на гитхабе, тем более давно хотелось похвастаться перед тов. Katachi.

## Как это юзать

Это классы. Идут они, как правило, со своим отдельным интерфейсом (традиционно с буковой I в начале). Игногда они требуют что-нибудь из соседней папки. Большинству требуется `EnumHelper` из папки `Third-party`. Также, многим классам требуется референс к `Microsoft Scripting Runtime`, для `Dictionary`, об этом я даже в описаниях указывать не буду, потому что словарь - полезная штука, добавляйте в проект по умолчанию, тем более что Scripting Runtime идёт с кореловским VBA в комплекте, то есть у пользователя оно точно будет.

Большая часть классов инициализируются не через *new*, а через встроенный конструктор. Эту тему я взял на вооружение у тов. [Мэтью Гиндона](https://github.com/retailcoder). Эти классы являются *predeclared*, имеют конструктор `Create`, который возвращает копию этого инициализированного класса (через отдельный интерфейс или дефолтный, то есть просто сам себя).

Например, `List` вызывается так:

```VBA
Dim MyList as IList
Set MyList = List.Create
```

То есть мы вызываем метод `Create` класса `List`, который возвращает нам свой новый экземпляр с интерфейсом `IList`.

Ну, пойдём по порядку.

## Список функциональных единиц

В каждой папке - подробное описание.

**[Composer](Composer)**  - простой укладчик шейпов или шейпренджей.

**[Either](https://github.com/elvin-nsk/LowCoupledFromCore/tree/main/Either)** - конструкция для работы с методами, которые могут выполняться с предусмотренной ошибкой.

**FileBrowser** - браузер файлов, standalone, так сказать.

**FileSpec** - хранит имя файла и путь к нему.

**List** - обёртка для коллекции, добавляющая функционал.

**Logger** - форма логгера чего-нибудь, типа ошибок.

**Point** - хранит две координаты точки.

**RecordList** - коллекция словарей с поиском по всем полям.

**RecordListToTableBinder** - класс для привязки `RecordList`а к `TableFile`.

**TableFile** - универсальный интерфейс для двумерного массива, читай - табличных данных.

**Third-party** - чужой код, необходимый для работы некоторых модулей.

### FileBrowser

Довольно простенький. Может вызвать окошко "открыть" ("сохранить" пока не допилил), поддерживает мультивыбор, возвращает коллекцию строк (имён файлов). Изначальный код откуда-то из интернетов, авторство мне не известно.

### FileSpec

Хранит имя файла. Можно распарсить полное имя, задать/получить путь, имя, расширение.

### List

Обёртка для коллекции, добавляющая функционал. Можно изменять элементы (Let/Set), `Append`ить другие листы и коллекции, запрашивать наличие элемента, получать копию себя.

Требует `EnumHelper`.

### Logger

Форма логгера чего-нибудь, типа ошибок. Умеет хранить ссылки на объекты корела и по кнопочке их выделять (если шейпы) или открывать (документ, страница).

Очень простецкая, наговнокодил на коленке, давно и неправда. Без конструктора и интерфейса, вызывается через *new*. Добавляем элементы через `Add`, запрашиваем проверку через `Check` - если что-то было добавлено, то форма покажется.

### Point

Просто хранит две координаты точки. Может вращать себя вокруг другой точки.

### RecordList

Относительно простая база данных. По сути, коллекция словарей с поиском по всем полям.

Требует `List` и `EnumHelper`.

### RecordListToTableBinder

Класс для привязки `RecordList`а к `TableFile`. Cначала создаёте `TableFile`, потом скармливаете биндеру - на выходе получаете заполненный из таблицы `RecordList`. Все изменения в нём (если не открыли рид-онли) сохраняются в `TableFile` по завершении жизни класса.

Инициализируется билдером, это тот же Create, только с вынесенными в пропертис параметрами, как задали параметры - инициализируем через `Build`. Можно задавать `MandatoryKey` - обязательные ключи (в таблице ключ - это колонка), если в таблице под этим ключом пустое поле, то в `RecordList` эта запись (строка таблицы) не попадёт. Также не попадёт запись с пустым полем под первичным ключом, если он задан. Ещё можно задать `UnboundKey` - это дополнительные ключи, которые не сохраняются в таблице.

### TableFile

Задумывалось как универсальный интерфейс для двумерного массива, читай - табличных данных. Под этот интерфейс я написал что-то для ковыряния `csv` (пока не выложил), `xls` напрямую через *Excel* (получилась кака, при краше остаются висеть экселевские процессы), и `ExcelConnection` - а вот это получилось что надо, на нём и остановился. Даёт доступ к экселевскому файлу через *ADODB*. Позволяет менять данные даже в открытом параллельно в экселе файле, прямо на глазах изумлённой публики.

Можно открывать xls и xlsx. По завершении жизни класса все изменения сохраняются в таблицу (если не открыли рид-онли).

Требует `FileSpec`.

### Third-party

Тут чужой код, без которого ничего работать не будет. Ну, кое-что будет, ок. Щас там `EnumHelper` - полезный хелпер, позволяющий провайдить энумератор через интерфейс (так оно в 64-битном VBA не работает). Взял [отсюда](https://github.com/cristianbuse/VBA-KeyedCollection), почитайте там, если ничего не поняли. У тов. [Кристиана Буза](https://github.com/cristianbuse) ещё много всякого интересного.

## License

Да какая, к чёрту, русскому лицензия. Но если у вас взорвётся утюг - я тут нипричём.
