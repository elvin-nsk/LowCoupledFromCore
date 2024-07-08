# Low-coupled from Core

Пока я в тайне от человечества пишу мега-проект Core - что-то вроде фреймворка для корела - иногда хочется иметь кусочки егойного функционала, чтобы сунуть в какой-нибудь проектик. Внутри Core всё высокосвязное, поэтому так просто не оторвёшь. Поэтому начал вытаскивать для себя переделанные кусочки с пониженной <s>социальной ответственностью</s> связанностью. Пусть лежит и на гитхабе, тем более давно хотелось похвастаться перед тов. Katachi.

## Как это юзать

Это классы. Ну, кроме главного модуля `LibCore`. Некоторые из них всё ещё идут с отдельным интерфейсом (традиционно с буковкой I в начале), однако я от этого отхожу в сторону того, чтобы класс сам был себе интерфейсом. Конструкторы (см. ниже) прячем за `Friend`, чтобы они не имплементились, если нам захочется заимплементить класс как интерфейс.

Игногда классам требуется что-нибудь из соседней папки. Также, многим требуется референс к `Microsoft Scripting Runtime`, для `Dictionary`, об этом я даже в описаниях указывать не буду, потому что словарь - полезная штука, добавляйте в проект по умолчанию, тем более что Scripting Runtime идёт с кореловским VBA в комплекте, то есть у пользователя оно точно будет.

Большая часть классов инициализируются не через ключевое слово *new*, а через встроенный конструктор. Эту тему я взял на вооружение у тов. [Мэтью Гиндона](https://github.com/retailcoder). Такие классы являются *predeclared*, имеют конструктор `New_`, который возвращает копию этого инициализированного класса (через отдельный интерфейс или дефолтный, то есть просто сам себя).

Например, `List` вызывается так:

```VBA
Dim MyList as IList
Set MyList = List.New_
```

То есть мы вызываем метод `New_` класса `List`, который возвращает нам свой новый экземпляр с интерфейсом `IList`.

Ну, пойдём по порядку.

## Список функциональных единиц

В некоторых папках есть подробное описание.

- [Composer](Composer) - простой укладчик шейпов или шейпренджей.
- [CsvTools](CsvTools) - парсинг CSV-файлов, в том числе под интерфейсом `TableFile` (см. ниже).
- [Either](Either) - конструкция для работы с методами, которые могут выполняться с предусмотренной ошибкой.
- [FileBrowser](FileBrowser) - браузер файлов, standalone, так сказать.
- [FileSpec](FileSpec) - хранит имя файла и путь к нему.
- [InputData](InputData) - получение исходного пользовательского выбора.
- [LibCore](LibCore) - основной модуль с большим набором всяких функций - от полезных до легаси мусора.
- [List](List) - обёртка для коллекции, добавляющая функционал.
- [Loggers](Loggers) - форма логгера чего-нибудь, типа ошибок - два варианта.
- [MarksSetter](MarksSetter) - установщик меток реза, используется в `MotifTools`.
- [MotifTools](MotifTools) - продвинутый укладчик изделий (мотивов) на лист с совмещением оборота, метками и прочими прибамбасами, основан на `Composer`'е и является его логическим развитием.
- [Point](Point) - хранит две координаты точки.
- [Json](Json) - сохранение настроек в json-файле, с управлением пресетами и шаблоном формы.
- [ProgressBar](ProgressBar) - как ни странно.
- [RecordList](RecordList) - коллекция словарей с поиском по всем полям.
- [RecordListToTableBinder](RecordListToTableBinder) - класс для привязки `RecordList`а к `TableFile`.
- [StringLocalizer](StringLocalizer) - относительно простой механизм локализации.
- [TableFile](TableFile) - универсальный интерфейс для двумерного массива, читай - табличных данных.
- [Third-party ](Third-party) - чужой код, необходимый для работы некоторых модулей.
- [ViewHandlers](ViewHandlers) - инструменты для контроля разных полей на форме.

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

Инициализируется билдером, это тот же `New_`, только с вынесенными в пропертис параметрами, как задали параметры - инициализируем через `Build`. Можно задавать `MandatoryKey` - обязательные ключи (в таблице ключ - это колонка), если в таблице под этим ключом пустое поле, то в `RecordList` эта запись (строка таблицы) не попадёт. Также не попадёт запись с пустым полем под первичным ключом, если он задан. Ещё можно задать `UnboundKey` - это дополнительные ключи, которые не сохраняются в таблице.

### TableFile

Задумывалось как универсальный интерфейс для двумерного массива, читай - табличных данных. Под этот интерфейс я написал `CsvUtilsTableFile ` для ковыряния `csv`, `xls` напрямую через *Excel* (получилась кака, при краше остаются висеть экселевские процессы), и `ExcelConnection` - а вот это получилось что надо, на нём и остановился. Даёт доступ к экселевскому файлу через *ADODB*. Позволяет менять данные даже в открытом параллельно в экселе файле, прямо на глазах изумлённой публики.

Можно открывать xls и xlsx. По завершении жизни класса все изменения сохраняются в таблицу (если не открыли рид-онли).

Требует `FileSpec`.

### Third-party

Тут чужой код, без которого ничего работать не будет. Ну, кое-что будет, ок. Щас там `EnumHelper` - полезный хелпер, позволяющий провайдить энумератор через интерфейс (так оно в 64-битном VBA не работает). Взял [отсюда](https://github.com/cristianbuse/VBA-KeyedCollection), почитайте там, если ничего не поняли. У тов. [Кристиана Буза](https://github.com/cristianbuse) ещё много всякого интересного.

## License

Да какая, к чёрту, русскому лицензия. Но если у вас взорвётся утюг - я тут нипричём.
