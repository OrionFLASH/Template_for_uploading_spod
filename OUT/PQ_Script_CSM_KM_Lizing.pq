let
    // ==================== НАСТРОЙКИ ====================
    ФорматЧисла = "0.00000",
    ДлинаТабельногоНомера = 20,
    ЗначениеПриоритета = "1",
    ЗначениеБезДубля = "-",
    ПрефиксТаблицыДанных = "_data_base",
    ПрефиксТаблицыПараметров = "_param_base",
    ИмяТаблицыСправочникФИО = "_data_fio",
    КолонкаТабельныйНомер = "ТАБЕЛЬНЫЙ НОМЕР",
    КолонкаФИО = "ФИО",
    КолонкаПоказатель = "ПОКАЗАТЕЛЬ",
    КолонкаСправочникФИО = "Клиентский менеджер",
    КолонкаСправочникКод = "Код КМ",
    КлючПараметраTypeInd = "type_ind",
    ЗначениеTypeIndTN = "TN",
    ЗначениеTypeIndFIO = "FIO",
    ТабельныйЕслиФИОНеНайден = "00000000",
    КлючПараметраFullName = "CONTEST-DATA=>FULL_NAME",
    КолонкаFullName = "CONTEST-DATA=>FULL_NAME",
    КолонкаFIO = "FIO",
    КолонкаТабельныйНеНайден = "ТАБЕЛЬНЫЙ НЕ НАЙДЕН",
    ЗначениеТабельныйНеНайден = "ДА",
    // ===================================================

    // Шаг_01_Получение_всех_смарт_таблиц_текущей_книги
    Шаг_01_Источник = Excel.CurrentWorkbook(),

    // Шаг_02_Извлечение_списка_имён_таблиц
    Шаг_02_СписокИмён = Table.Column(Шаг_01_Источник, "Name"),

    // Шаг_02б_Справочник_ФИО_с_листа_FIO
    Шаг_02б_СправочникФИО =
        let
            ТаблицаСправочника =
                try Шаг_01_Источник{[Name = ИмяТаблицыСправочникФИО]}[Content]
                otherwise
                    #table(
                        type table [#"Клиентский менеджер" = text, #"Код КМ" = text],
                        {}
                    ),
            БезПустых =
                Table.SelectRows(
                    ТаблицаСправочника,
                    each
                        Record.Field(_, КолонкаСправочникФИО) <> null
                        and Text.Trim(Text.From(Record.Field(_, КолонкаСправочникФИО))) <> ""
                        and Text.Trim(Text.From(Record.Field(_, КолонкаСправочникФИО))) <> КолонкаСправочникФИО
                ),
            СКлючомПоиска =
                Table.AddColumn(
                    БезПустых,
                    "ФИО_ключ",
                    each Text.Upper(Text.Trim(Text.From(Record.Field(_, КолонкаСправочникФИО)))),
                    type text
                ),
            БезДублейКлючей = Table.Distinct(СКлючомПоиска, {"ФИО_ключ"})
        in
            БезДублейКлючей,

    // Вспомогательные функции: безопасная проверка префикса
    ИмяНачинаетсяС = (Имя as text, Префикс as text) as logical =>
        Text.Length(Имя) >= Text.Length(Префикс)
        and Text.Range(Имя, 0, Text.Length(Префикс)) = Префикс,

    ИзвлечьИндексИзИмени = (Имя as text, Префикс as text) as nullable number =>
        let
            Суффикс = Text.Range(Имя, Text.Length(Префикс)),
            Результат = if Text.Length(Суффикс) > 0 then try Number.From(Суффикс) otherwise null else null
        in
            Результат,

    // Шаг_03_Поиск_индексов_таблиц_данных
    Шаг_03_ИндексыДанных =
        List.Sort(
            List.Select(
                List.Transform(
                    List.Select(Шаг_02_СписокИмён, each ИмяНачинаетсяС(_, ПрефиксТаблицыДанных)),
                    each ИзвлечьИндексИзИмени(_, ПрефиксТаблицыДанных)
                ),
                each _ <> null
            )
        ),

    // Шаг_04_Поиск_индексов_таблиц_параметров
    Шаг_04_ИндексыПараметров =
        List.Sort(
            List.Select(
                List.Transform(
                    List.Select(Шаг_02_СписокИмён, each ИмяНачинаетсяС(_, ПрефиксТаблицыПараметров)),
                    each ИзвлечьИндексИзИмени(_, ПрефиксТаблицыПараметров)
                ),
                each _ <> null
            )
        ),

    // Шаг_05_Сопоставление_пар_по_общему_индексу
    Шаг_05_ИндексыПар = List.Intersect({Шаг_03_ИндексыДанных, Шаг_04_ИндексыПараметров}),

    // Шаг_06_Функция_преобразования_значения_с_запятой_в_число
    ИзЗапятойВЧисло = (Значение as any) as number =>
        if Значение = null then
            0
        else if Value.Is(Значение, type number) then
            Значение
        else
            let
                Текст = Text.Replace(Text.Replace(Text.Trim(Text.From(Значение)), " ", ""), ",", "."),
                Части = Text.Split(Текст, "."),
                ЦелаяЧасть = try Number.From(Части{0}) otherwise 0,
                Результат =
                    if List.Count(Части) = 1 then
                        ЦелаяЧасть
                    else
                        let
                            ДробьТекст = Части{1},
                            ДробьЧисло = try Number.From(ДробьТекст) otherwise 0,
                            Делитель = Number.Power(10, Text.Length(ДробьТекст))
                        in
                            if Делитель = 0 then ЦелаяЧасть else ЦелаяЧасть + ДробьЧисло / Делитель
            in
                Результат,

    // Шаг_07_Функция_форматирования_числа_в_текст_с_точкой
    ЧислоВТекстСТочкой = (Число as number) as text =>
        Text.Replace(Number.ToText(Число, ФорматЧисла), ",", "."),

    // Шаг_07б_Функция_форматирования_числа_с_запятой
    ЧислоВТекстСЗапятой = (Число as number) as text =>
        Text.Replace(ЧислоВТекстСТочкой(Число), ".", ","),

    // Шаг_08_Функция_извлечения_параметра_из_таблицы_key_value
    ПолучитьПараметрЗначение = (ТаблицаПараметров as table, Ключ as text) as any =>
        let
            ИменаКолонок = Table.ColumnNames(ТаблицаПараметров),
            КолонкаКлюч = ИменаКолонок{0},
            КолонкаЗначение = ИменаКолонок{1},
            Строки = Table.ToRows(Table.SelectColumns(ТаблицаПараметров, {КолонкаКлюч, КолонкаЗначение})),
            Найдено =
                List.Select(
                    Строки,
                    each List.Count(_) >= 2 and Text.Upper(Text.Trim(Text.From(_{0}))) = Text.Upper(Ключ)
                )
        in
            if List.Count(Найдено) > 0 then Найдено{0}{1} else null,

    ПолучитьПараметрТекст = (ТаблицаПараметров as table, Ключ as text) as text =>
        let
            Значение = ПолучитьПараметрЗначение(ТаблицаПараметров, Ключ)
        in
            if Значение = null then "" else Text.From(Значение),

    // Шаг_09_Функция_сборки_параметров_для_одной_пары
    ИзвлечьПараметры = (ТаблицаПараметров as table) as record =>
        let
            TypeIndСырой = Text.Upper(Text.Trim(ПолучитьПараметрТекст(ТаблицаПараметров, КлючПараметраTypeInd)))
        in
            [
                TYPE_IND = if TypeIndСырой = ЗначениеTypeIndFIO then ЗначениеTypeIndFIO else ЗначениеTypeIndTN,
                CONTEST_CODE = ПолучитьПараметрТекст(ТаблицаПараметров, "CONTEST_CODE"),
                TOURNAMENT_CODE = ПолучитьПараметрТекст(ТаблицаПараметров, "TOURNAMENT_CODE"),
                CONTEST_DATE = ПолучитьПараметрТекст(ТаблицаПараметров, "CONTEST_DATE"),
                PLAN_VALUE_NUM = ИзЗапятойВЧисло(ПолучитьПараметрЗначение(ТаблицаПараметров, "PLAN_VALUE")),
                FULL_NAME = ПолучитьПараметрТекст(ТаблицаПараметров, КлючПараметраFullName)
            ],

    // Шаг_10_Функция_нормализации_таблицы_данных_по_режиму_TN_или_FIO
    НормализоватьТаблицуДанных = (ТаблицаДанных as table, РежимFIO as logical) as table =>
        if РежимFIO then
            let
                БезПустых =
                    Table.SelectRows(
                        ТаблицаДанных,
                        each
                            Record.Field(_, КолонкаФИО) <> null
                            and Text.Trim(Text.From(Record.Field(_, КолонкаФИО))) <> ""
                            and Text.Trim(Text.From(Record.Field(_, КолонкаФИО))) <> КолонкаФИО
                    ),
                СИсходнымFIO =
                    Table.AddColumn(
                        БезПустых,
                        КолонкаFIO,
                        each Text.Trim(Text.From(Record.Field(_, КолонкаФИО))),
                        type text
                    ),
                СКлючомПоиска =
                    Table.AddColumn(
                        СИсходнымFIO,
                        "ФИО_ключ",
                        each Text.Upper([FIO]),
                        type text
                    ),
                СПрисоединением =
                    Table.NestedJoin(
                        СКлючомПоиска,
                        {"ФИО_ключ"},
                        Шаг_02б_СправочникФИО,
                        {"ФИО_ключ"},
                        "СправочникФИО",
                        JoinKind.LeftOuter
                    ),
                СРаскрытием =
                    Table.ExpandTableColumn(
                        СПрисоединением,
                        "СправочникФИО",
                        {КолонкаСправочникКод},
                        {КолонкаСправочникКод}
                    ),
                СТабельнымКодом =
                    Table.AddColumn(
                        СРаскрытием,
                        "ТабельныйКод",
                        each
                            if Record.Field(_, КолонкаСправочникКод) = null then
                                ТабельныйЕслиФИОНеНайден
                            else
                                Text.Trim(Text.From(Record.Field(_, КолонкаСправочникКод))),
                        type text
                    ),
                СФлагомНеНайден =
                    Table.AddColumn(
                        СТабельнымКодом,
                        КолонкаТабельныйНеНайден,
                        each
                            if Record.Field(_, КолонкаСправочникКод) = null then
                                ЗначениеТабельныйНеНайден
                            else
                                ЗначениеБезДубля,
                        type text
                    ),
                СТабельным =
                    Table.AddColumn(
                        СФлагомНеНайден,
                        "MANAGER_PERSON_NUMBER",
                        each Text.PadStart([ТабельныйКод], ДлинаТабельногоНомера, "0"),
                        type text
                    )
            in
                Table.RemoveColumns(СТабельным, {"ФИО_ключ", КолонкаСправочникКод, "ТабельныйКод"})
        else
            let
                БезПустых =
                    Table.SelectRows(
                        ТаблицаДанных,
                        each
                            Record.Field(_, КолонкаТабельныйНомер) <> null
                            and Text.Trim(Text.From(Record.Field(_, КолонкаТабельныйНомер))) <> ""
                            and Text.Trim(Text.From(Record.Field(_, КолонкаТабельныйНомер))) <> КолонкаТабельныйНомер
                    ),
                СТабельным =
                    Table.AddColumn(
                        БезПустых,
                        "MANAGER_PERSON_NUMBER",
                        each Text.PadStart(Text.Trim(Text.From(Record.Field(_, КолонкаТабельныйНомер))), ДлинаТабельногоНомера, "0"),
                        type text
                    ),
                СFIO =
                    Table.AddColumn(СТабельным, КолонкаFIO, each ЗначениеБезДубля, type text),
                СФлагом =
                    Table.AddColumn(СFIO, КолонкаТабельныйНеНайден, each ЗначениеБезДубля, type text)
            in
                СФлагом,

    // Шаг_11_Функция_обработки_одной_пары_таблиц
    ОбработатьПару = (Индекс as number) as table =>
        let
            ИмяДанных = ПрефиксТаблицыДанных & Text.From(Индекс),
            ИмяПараметров = ПрефиксТаблицыПараметров & Text.From(Индекс),
            ТаблицаДанных = Шаг_01_Источник{[Name = ИмяДанных]}[Content],
            ТаблицаПараметров = Шаг_01_Источник{[Name = ИмяПараметров]}[Content],
            Параметры = ИзвлечьПараметры(ТаблицаПараметров),
            РежимFIO = Параметры[TYPE_IND] = ЗначениеTypeIndFIO,
            ЗначениеCONTEST_CODE = Параметры[CONTEST_CODE],
            ЗначениеTOURNAMENT_CODE = Параметры[TOURNAMENT_CODE],
            ЗначениеCONTEST_DATE = Параметры[CONTEST_DATE],
            ЗначениеPLAN_VALUE_NUM = Параметры[PLAN_VALUE_NUM],
            ЗначениеFullName = Параметры[FULL_NAME],
            НормализованныеДанные = НормализоватьТаблицуДанных(ТаблицаДанных, РежимFIO),
            СКодомКонтеста = Table.AddColumn(НормализованныеДанные, "CONTEST_CODE", each ЗначениеCONTEST_CODE, type text),
            СКодомТурнира = Table.AddColumn(СКодомКонтеста, "TOURNAMENT_CODE", each ЗначениеTOURNAMENT_CODE, type text),
            СДатойКонтеста = Table.AddColumn(СКодомТурнира, "CONTEST_DATE", each ЗначениеCONTEST_DATE, type text),
            СПлановымЧислом = Table.AddColumn(СДатойКонтеста, "PLAN_VALUE_число", each ЗначениеPLAN_VALUE_NUM, type number),
            СФактическимЧислом =
                Table.AddColumn(
                    СПлановымЧислом,
                    "FACT_VALUE_число",
                    each ИзЗапятойВЧисло(Record.Field(_, КолонкаПоказатель)),
                    type number
                ),
            СПлановымТекст = Table.AddColumn(СФактическимЧислом, "PLAN_VALUE", each ЧислоВТекстСТочкой([PLAN_VALUE_число]), type text),
            СФактическимТекст = Table.AddColumn(СПлановымТекст, "FACT_VALUE", each ЧислоВТекстСТочкой([FACT_VALUE_число]), type text),
            СПриоритетом = Table.AddColumn(СФактическимТекст, "priority_type", each ЗначениеПриоритета, type text),
            СFullName = Table.AddColumn(СПриоритетом, КолонкаFullName, each ЗначениеFullName, type text),
            Результат =
                Table.SelectColumns(
                    СFullName,
                    {
                        "MANAGER_PERSON_NUMBER",
                        "CONTEST_CODE",
                        "TOURNAMENT_CODE",
                        "CONTEST_DATE",
                        "PLAN_VALUE",
                        "FACT_VALUE",
                        "priority_type",
                        КолонкаFullName,
                        КолонкаFIO,
                        КолонкаТабельныйНеНайден,
                        "PLAN_VALUE_число",
                        "FACT_VALUE_число"
                    }
                )
        in
            Результат,

    // Шаг_12_Обработка_всех_найденных_пар
    Шаг_12_СписокОбработанныхТаблиц = List.Transform(Шаг_05_ИндексыПар, each ОбработатьПару(_)),

    // Шаг_13_Объединение_всех_строк
    Шаг_13_ВсеДанные =
        if List.Count(Шаг_12_СписокОбработанныхТаблиц) = 0 then
            #table(
                type table [
                    MANAGER_PERSON_NUMBER = text,
                    CONTEST_CODE = text,
                    TOURNAMENT_CODE = text,
                    CONTEST_DATE = text,
                    PLAN_VALUE = text,
                    FACT_VALUE = text,
                    priority_type = text,
                    #"CONTEST-DATA=>FULL_NAME" = text,
                    FIO = text,
                    #"ТАБЕЛЬНЫЙ НЕ НАЙДЕН" = text,
                    PLAN_VALUE_число = number,
                    FACT_VALUE_число = number
                ],
                {}
            )
        else
            Table.Combine(Шаг_12_СписокОбработанныхТаблиц),

    // Шаг_14_Создание_ключа_для_поиска_дублей
    Шаг_14_СКлючом =
        Table.AddColumn(
            Шаг_13_ВсеДанные,
            "ВнутреннийКлюч",
            each [CONTEST_CODE] & "|" & [TOURNAMENT_CODE] & "|" & [MANAGER_PERSON_NUMBER],
            type text
        ),

    // Шаг_15_Группировка_по_ключу
    Шаг_15_Группировка =
        Table.Group(
            Шаг_14_СКлючом,
            {"ВнутреннийКлюч"},
            {
                {"Количество", each Table.RowCount(_), Int64.Type},
                {"СуммаЗначений", each List.Sum([FACT_VALUE_число]), type number}
            }
        ),

    // Шаг_16_Присоединение_данных_о_дублях
    Шаг_16_СДублями =
        Table.NestedJoin(
            Шаг_14_СКлючом,
            {"ВнутреннийКлюч"},
            Шаг_15_Группировка,
            {"ВнутреннийКлюч"},
            "ДанныеГруппировки",
            JoinKind.LeftOuter
        ),

    // Шаг_17_Разворот_вложенной_таблицы_группировки
    Шаг_17_СРазвернутымиДанными =
        Table.ExpandTableColumn(
            Шаг_16_СДублями,
            "ДанныеГруппировки",
            {"Количество", "СуммаЗначений"},
            {"Количество", "СуммаЗначений"}
        ),

    // Шаг_18_Замена_null_в_служебных_полях
    Шаг_18_СЗаменойNull =
        Table.ReplaceValue(
            Table.ReplaceValue(Шаг_17_СРазвернутымиДанными, null, 1, Replacer.ReplaceValue, {"Количество"}),
            null,
            0,
            Replacer.ReplaceValue,
            {"СуммаЗначений"}
        ),

    // Шаг_19_Колонка_ДУБЛЬ
    Шаг_19_СКолонкойДубль =
        Table.AddColumn(
            Шаг_18_СЗаменойNull,
            "ДУБЛЬ",
            each if [Количество] > 1 then Text.From([Количество]) else ЗначениеБезДубля,
            type text
        ),

    // Шаг_20_Колонка_СУММА_ПО_ТАБЕЛЬНОМУ
    Шаг_20_СКолонкойСумма =
        Table.AddColumn(
            Шаг_19_СКолонкойДубль,
            "СУММА ПО ТАБЕЛЬНОМУ",
            each if [Количество] > 1 then ЧислоВТекстСЗапятой([СуммаЗначений]) else ЗначениеБезДубля,
            type text
        ),

    // Шаг_21_Колонка_В_ФОРМАТЕ
    Шаг_21_СКолонкойФормат =
        Table.AddColumn(
            Шаг_20_СКолонкойСумма,
            "В ФОРМАТЕ",
            each if [Количество] > 1 then ЧислоВТекстСТочкой([СуммаЗначений]) else ЗначениеБезДубля,
            type text
        ),

    // Шаг_22_Удаление_служебных_колонок
    Шаг_22_БезСлужебных =
        Table.RemoveColumns(Шаг_21_СКолонкойФормат, {"ВнутреннийКлюч", "Количество", "СуммаЗначений"}),

    // Шаг_23_Финальный_порядок_колонок
    Шаг_23_ИтоговаяТаблица =
        Table.ReorderColumns(
            Шаг_22_БезСлужебных,
            {
                "MANAGER_PERSON_NUMBER",
                "CONTEST_CODE",
                "TOURNAMENT_CODE",
                "CONTEST_DATE",
                "PLAN_VALUE",
                "FACT_VALUE",
                "priority_type",
                "CONTEST-DATA=>FULL_NAME",
                "ДУБЛЬ",
                "СУММА ПО ТАБЕЛЬНОМУ",
                "В ФОРМАТЕ",
                "PLAN_VALUE_число",
                "FACT_VALUE_число",
                "FIO",
                "ТАБЕЛЬНЫЙ НЕ НАЙДЕН"
            }
        )
in
    Шаг_23_ИтоговаяТаблица
