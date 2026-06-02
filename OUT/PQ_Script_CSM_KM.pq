let
    // ==================== НАСТРОЙКИ ====================
    ФорматЧисла = "0.00000",
    ДлинаТабельногоНомера = 20,
    ЗначениеПриоритета = "1",
    ЗначениеБезДубля = "-",
    ПрефиксТаблицыДанных = "_data_base",
    ПрефиксТаблицыПараметров = "_param_base",
    КолонкаТабельныйНомер = "ТАБЕЛЬНЫЙ НОМЕР",
    КолонкаПоказатель = "ПОКАЗАТЕЛЬ",
    КлючиПараметров = {"CONTEST_CODE", "TOURNAMENT_CODE", "PLAN_VALUE", "CONTEST_DATE", "CONTEST-DATA=>FULL_NAME"},
    КлючПараметраFullName = "CONTEST-DATA=>FULL_NAME",
    КолонкаFullName = "CONTEST-DATA=>FULL_NAME",
    КолонкаПараметрКлюч = "Поле",
    КолонкаПараметрЗначение = "Значение",
    // ===================================================

    // Шаг_01_Получение_всех_смарт_таблиц_текущей_книги
    Шаг_01_Источник = Excel.CurrentWorkbook(),

    // Шаг_02_Извлечение_списка_имён_таблиц
    Шаг_02_СписокИмён = Table.Column(Шаг_01_Источник, "Name"),

    // Вспомогательные функции: безопасная проверка префикса (имя может быть короче префикса)
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
    // В русской локали Number.From("1000000.00000") даёт ошибку — точка не распознаётся как дробная
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

    // Шаг_07_Функция_форматирования_числа_в_текст_с_точкой (формат 0.00000 для FACT_VALUE, PLAN_VALUE, «В ФОРМАТЕ»)
    ЧислоВТекстСТочкой = (Число as number) as text =>
        Text.Replace(Number.ToText(Число, ФорматЧисла), ",", "."),

    // Шаг_07б_Функция_форматирования_числа_с_запятой (числовой формат для «СУММА ПО ТАБЕЛЬНОМУ»)
    ЧислоВТекстСЗапятой = (Число as number) as text =>
        Text.Replace(ЧислоВТекстСТочкой(Число), ".", ","),

    // Шаг_08_Функция_извлечения_параметра_из_таблицы_key_value
    // Всегда берём 1-ю колонку как ключ, 2-ю как значение (независимо от имён «Поле»/«Значение»)
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

    // Шаг_09_Функция_сборки_параметров_для_одной_пары (все значения — простые переменные)
    ИзвлечьПараметры = (ТаблицаПараметров as table) as record =>
        [
            CONTEST_CODE = ПолучитьПараметрТекст(ТаблицаПараметров, "CONTEST_CODE"),
            TOURNAMENT_CODE = ПолучитьПараметрТекст(ТаблицаПараметров, "TOURNAMENT_CODE"),
            CONTEST_DATE = ПолучитьПараметрТекст(ТаблицаПараметров, "CONTEST_DATE"),
            PLAN_VALUE_NUM = ИзЗапятойВЧисло(ПолучитьПараметрЗначение(ТаблицаПараметров, "PLAN_VALUE")),
            FULL_NAME = ПолучитьПараметрТекст(ТаблицаПараметров, КлючПараметраFullName)
        ],

    // Шаг_10_Функция_обработки_одной_пары_таблиц
    ОбработатьПару = (Индекс as number) as table =>
        let
            ИмяДанных = ПрефиксТаблицыДанных & Text.From(Индекс),
            ИмяПараметров = ПрефиксТаблицыПараметров & Text.From(Индекс),
            ТаблицаДанных = Шаг_01_Источник{[Name = ИмяДанных]}[Content],
            ТаблицаПараметров = Шаг_01_Источник{[Name = ИмяПараметров]}[Content],
            Параметры = ИзвлечьПараметры(ТаблицаПараметров),
            // Явные переменные — не обращаться к полям записи внутри each (избегаем ошибок парсера M)
            ЗначениеCONTEST_CODE = Параметры[CONTEST_CODE],
            ЗначениеTOURNAMENT_CODE = Параметры[TOURNAMENT_CODE],
            ЗначениеCONTEST_DATE = Параметры[CONTEST_DATE],
            ЗначениеPLAN_VALUE_NUM = Параметры[PLAN_VALUE_NUM],
            ЗначениеFullName = Параметры[FULL_NAME],
            БезПустыхСтрок =
                Table.SelectRows(
                    ТаблицаДанных,
                    each
                        Record.Field(_, КолонкаТабельныйНомер) <> null
                        and Text.Trim(Text.From(Record.Field(_, КолонкаТабельныйНомер))) <> ""
                        and Text.Trim(Text.From(Record.Field(_, КолонкаТабельныйНомер))) <> КолонкаТабельныйНомер
                ),
            СТабельным =
                Table.AddColumn(
                    БезПустыхСтрок,
                    "MANAGER_PERSON_NUMBER",
                    each Text.PadStart(Text.Trim(Text.From(Record.Field(_, КолонкаТабельныйНомер))), ДлинаТабельногоНомера, "0"),
                    type text
                ),
            СКодомКонтеста = Table.AddColumn(СТабельным, "CONTEST_CODE", each ЗначениеCONTEST_CODE, type text),
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
                        "PLAN_VALUE_число",
                        "FACT_VALUE_число"
                    }
                )
        in
            Результат,

    // Шаг_11_Обработка_всех_найденных_пар
    Шаг_11_СписокОбработанныхТаблиц = List.Transform(Шаг_05_ИндексыПар, each ОбработатьПару(_)),

    // Шаг_12_Объединение_всех_строк
    Шаг_12_ВсеДанные =
        if List.Count(Шаг_11_СписокОбработанныхТаблиц) = 0 then
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
                    PLAN_VALUE_число = number,
                    FACT_VALUE_число = number
                ],
                {}
            )
        else
            Table.Combine(Шаг_11_СписокОбработанныхТаблиц),

    // Шаг_13_Создание_ключа_для_поиска_дублей
    Шаг_13_СКлючом =
        Table.AddColumn(
            Шаг_12_ВсеДанные,
            "ВнутреннийКлюч",
            each [CONTEST_CODE] & "|" & [TOURNAMENT_CODE] & "|" & [MANAGER_PERSON_NUMBER],
            type text
        ),

    // Шаг_14_Группировка_по_ключу
    Шаг_14_Группировка =
        Table.Group(
            Шаг_13_СКлючом,
            {"ВнутреннийКлюч"},
            {
                {"Количество", each Table.RowCount(_), Int64.Type},
                {"СуммаЗначений", each List.Sum([FACT_VALUE_число]), type number}
            }
        ),

    // Шаг_15_Присоединение_данных_о_дублях
    Шаг_15_СДублями =
        Table.NestedJoin(
            Шаг_13_СКлючом,
            {"ВнутреннийКлюч"},
            Шаг_14_Группировка,
            {"ВнутреннийКлюч"},
            "ДанныеГруппировки",
            JoinKind.LeftOuter
        ),

    // Шаг_16_Разворот_вложенной_таблицы_группировки
    Шаг_16_СРазвернутымиДанными =
        Table.ExpandTableColumn(
            Шаг_15_СДублями,
            "ДанныеГруппировки",
            {"Количество", "СуммаЗначений"},
            {"Количество", "СуммаЗначений"}
        ),

    // Шаг_17_Замена_null_в_служебных_полях
    Шаг_17_СЗаменойNull =
        Table.ReplaceValue(
            Table.ReplaceValue(Шаг_16_СРазвернутымиДанными, null, 1, Replacer.ReplaceValue, {"Количество"}),
            null,
            0,
            Replacer.ReplaceValue,
            {"СуммаЗначений"}
        ),

    // Шаг_18_Колонка_ДУБЛЬ
    Шаг_18_СКолонкойДубль =
        Table.AddColumn(
            Шаг_17_СЗаменойNull,
            "ДУБЛЬ",
            each if [Количество] > 1 then Text.From([Количество]) else ЗначениеБезДубля,
            type text
        ),

    // Шаг_19_Колонка_СУММА_ПО_ТАБЕЛЬНОМУ (числовой формат с запятой: 381,00000)
    Шаг_19_СКолонкойСумма =
        Table.AddColumn(
            Шаг_18_СКолонкойДубль,
            "СУММА ПО ТАБЕЛЬНОМУ",
            each if [Количество] > 1 then ЧислоВТекстСЗапятой([СуммаЗначений]) else ЗначениеБезДубля,
            type text
        ),

    // Шаг_20_Колонка_В_ФОРМАТЕ (формат 0.00000 с точкой: 381.00000)
    Шаг_20_СКолонкойФормат =
        Table.AddColumn(
            Шаг_19_СКолонкойСумма,
            "В ФОРМАТЕ",
            each if [Количество] > 1 then ЧислоВТекстСТочкой([СуммаЗначений]) else ЗначениеБезДубля,
            type text
        ),

    // Шаг_21_Удаление_служебных_колонок
    Шаг_21_БезСлужебных =
        Table.RemoveColumns(Шаг_20_СКолонкойФормат, {"ВнутреннийКлюч", "Количество", "СуммаЗначений"}),

    // Шаг_22_Финальный_порядок_колонок
    Шаг_22_ИтоговаяТаблица =
        Table.ReorderColumns(
            Шаг_21_БезСлужебных,
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
                "FACT_VALUE_число"
            }
        )
in
    Шаг_22_ИтоговаяТаблица
