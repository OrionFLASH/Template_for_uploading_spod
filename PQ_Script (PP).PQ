Вот моя версия скрипта Power Query, составленная строго по бизнес-требованиям с небольшими улучшениями в структуре и надёжности:

```pq
let
    // ==================== НАСТРОЙКИ ====================
    ФорматЧисла = "0.00000",
    ДлинаТабельногоНомера = 20,
    ИмяЛистаДанных = "Данные",
    // ===================================================

    // Шаг_01_Получение_всех_таблиц_книги
    Источник = Excel.CurrentWorkbook(),

    // Шаг_02_Загрузка_таблиц_данных_по_именам
    _data_base1 = Источник{[Name="_data_base1"]}[Content],
    _data_base2 = Источник{[Name="_data_base2"]}[Content],
    _data_base3 = Источник{[Name="_data_base3"]}[Content],
    _data_base4 = Источник{[Name="_data_base4"]}[Content],
    _data_base5 = Источник{[Name="_data_base5"]}[Content],

    // Шаг_03_Загрузка_таблиц_параметров_по_именам
    _param_base1 = Источник{[Name="_param_base1"]}[Content],
    _param_base2 = Источник{[Name="_param_base2"]}[Content],
    _param_base3 = Источник{[Name="_param_base3"]}[Content],
    _param_base4 = Источник{[Name="_param_base4"]}[Content],
    _param_base5 = Источник{[Name="_param_base5"]}[Content],

    // Шаг_04_Функция_безопасного_преобразования_текста_с_запятой_в_число
    SafeTextToNumber = (Значение as any) as number =>
        let
            Текст = if Value.Is(Значение, type text) then Значение else Text.From(Значение),
            СЗаменой = Text.Replace(Текст, ",", "."),
            Очищенный = Text.Trim(СЗаменой),
            Результат = try Number.From(Очищенный) otherwise 0
        in
            Результат,

    // Шаг_05_Функция_поиска_параметра_по_ключу_в_таблице_без_заголовков
    GetParamValue = (Таблица as table, Ключ as text) as text =>
        let
            ВсеСтроки = Table.ToRows(Таблица),
            Фильтр = List.Select(ВсеСтроки, each 
                if List.Count(_) >= 2 then 
                    Text.Trim(Text.From(_{0})) = Ключ 
                else false
            ),
            Результат = if List.Count(Фильтр) > 0 then Text.From(Фильтр{0}{1}) else ""
        in
            Результат,

    // Шаг_06_Извлечение_параметров_для_каждой_базы
    Params1 = [
        CONTEST_CODE = GetParamValue(_param_base1, "CONTEST_CODE"),
        TOURNAMENT_CODE = GetParamValue(_param_base1, "TOURNAMENT_CODE"),
        PLAN_VALUE_NUM = SafeTextToNumber(GetParamValue(_param_base1, "PLAN_VALUE")),
        CONTEST_DATE = GetParamValue(_param_base1, "CONTEST_DATE")
    ],
    Params2 = [
        CONTEST_CODE = GetParamValue(_param_base2, "CONTEST_CODE"),
        TOURNAMENT_CODE = GetParamValue(_param_base2, "TOURNAMENT_CODE"),
        PLAN_VALUE_NUM = SafeTextToNumber(GetParamValue(_param_base2, "PLAN_VALUE")),
        CONTEST_DATE = GetParamValue(_param_base2, "CONTEST_DATE")
    ],
    Params3 = [
        CONTEST_CODE = GetParamValue(_param_base3, "CONTEST_CODE"),
        TOURNAMENT_CODE = GetParamValue(_param_base3, "TOURNAMENT_CODE"),
        PLAN_VALUE_NUM = SafeTextToNumber(GetParamValue(_param_base3, "PLAN_VALUE")),
        CONTEST_DATE = GetParamValue(_param_base3, "CONTEST_DATE")
    ],
    Params4 = [
        CONTEST_CODE = GetParamValue(_param_base4, "CONTEST_CODE"),
        TOURNAMENT_CODE = GetParamValue(_param_base4, "TOURNAMENT_CODE"),
        PLAN_VALUE_NUM = SafeTextToNumber(GetParamValue(_param_base4, "PLAN_VALUE")),
        CONTEST_DATE = GetParamValue(_param_base4, "CONTEST_DATE")
    ],
    Params5 = [
        CONTEST_CODE = GetParamValue(_param_base5, "CONTEST_CODE"),
        TOURNAMENT_CODE = GetParamValue(_param_base5, "TOURNAMENT_CODE"),
        PLAN_VALUE_NUM = SafeTextToNumber(GetParamValue(_param_base5, "PLAN_VALUE")),
        CONTEST_DATE = GetParamValue(_param_base5, "CONTEST_DATE")
    ],

    // Шаг_07_Функция_обработки_одной_таблицы_с_добавлением_всех_колонок
    ProcessDataTable = (ТаблицаДанных as table, Параметры as record) as table =>
        let
            // Добавляем MANAGER_PERSON_NUMBER с лидирующими нулями
            сТабельнымНомером = Table.AddColumn(ТаблицаДанных, "MANAGER_PERSON_NUMBER", each 
                Text.PadStart(Text.From([ТАБЕЛЬНЫЙ НОМЕР]), ДлинаТабельногоНомера, "0"), type text),
            
            // Добавляем параметры из записи
            сCONTEST_CODE = Table.AddColumn(сТабельнымНомером, "CONTEST_CODE", each Параметры[CONTEST_CODE], type text),
            сTOURNAMENT_CODE = Table.AddColumn(сCONTEST_CODE, "TOURNAMENT_CODE", each Параметры[TOURNAMENT_CODE], type text),
            сCONTEST_DATE = Table.AddColumn(сTOURNAMENT_CODE, "CONTEST_DATE", each Параметры[CONTEST_DATE], type text),
            
            // Добавляем PLAN_VALUE_число (оригинальное число)
            сPLAN_VALUE_NUM = Table.AddColumn(сCONTEST_DATE, "PLAN_VALUE_число", each Параметры[PLAN_VALUE_NUM], type number),
            
            // Добавляем FACT_VALUE_число (преобразованное из ПОКАЗАТЕЛЬ)
            сFACT_VALUE_NUM = Table.AddColumn(сPLAN_VALUE_NUM, "FACT_VALUE_число", each 
                SafeTextToNumber([ПОКАЗАТЕЛЬ]), type number),
            
            // Добавляем текстовые версии с форматом 0.00000
            сPLAN_VALUE_TEXT = Table.AddColumn(сFACT_VALUE_NUM, "PLAN_VALUE", each 
                Number.ToText([PLAN_VALUE_число], ФорматЧисла), type text),
            сFACT_VALUE_TEXT = Table.AddColumn(сPLAN_VALUE_TEXT, "FACT_VALUE", each 
                Number.ToText([FACT_VALUE_число], ФорматЧисла), type text),
            
            // Добавляем priority_type
            сPriority = Table.AddColumn(сFACT_VALUE_TEXT, "priority_type", each "1", type text),
            
            // Выбираем только нужные колонки в нужном порядке
            Результат = Table.SelectColumns(сPriority, {
                "MANAGER_PERSON_NUMBER",
                "CONTEST_CODE",
                "TOURNAMENT_CODE",
                "CONTEST_DATE",
                "PLAN_VALUE",
                "FACT_VALUE",
                "priority_type",
                "PLAN_VALUE_число",
                "FACT_VALUE_число"
            })
        in
            Результат,

    // Шаг_08_Обработка_каждой_из_пяти_таблиц
    Processed1 = ProcessDataTable(_data_base1, Params1),
    Processed2 = ProcessDataTable(_data_base2, Params2),
    Processed3 = ProcessDataTable(_data_base3, Params3),
    Processed4 = ProcessDataTable(_data_base4, Params4),
    Processed5 = ProcessDataTable(_data_base5, Params5),

    // Шаг_09_Объединение_всех_обработанных_таблиц
    CombinedData = Table.Combine({Processed1, Processed2, Processed3, Processed4, Processed5}),

    // Шаг_10_Создание_составного_ключа_для_поиска_дублей
    сKey = Table.AddColumn(CombinedData, "DupKey", each 
        [CONTEST_CODE] & "|" & [TOURNAMENT_CODE] & "|" & [MANAGER_PERSON_NUMBER], type text),

    // Шаг_11_Группировка_по_ключу_для_подсчёта_дублей_и_суммы
    Grouped = Table.Group(сKey, {"DupKey"}, {
        {"DupCount", each Table.RowCount(_), type number},
        {"TotalSum", each List.Sum([FACT_VALUE_число]), type number}
    }),

    // Шаг_12_Присоединение_информации_о_дублях_к_основной_таблице
    сDupInfo = Table.NestedJoin(сKey, {"DupKey"}, Grouped, {"DupKey"}, "DupData", JoinKind.LeftOuter),

    // Шаг_13_Разворачивание_вложенной_таблицы_с_данными_о_дублях
    Expanded = Table.ExpandTableColumn(сDupInfo, "DupData", {"DupCount", "TotalSum"}, {"DupCount", "TotalSum"}),

    // Шаг_14_Замена_null_на_дефолтные_значения
    ReplaceNulls = Table.ReplaceValue(Expanded, null, 1, Replacer.ReplaceValue, {"DupCount"}),
    ReplaceNulls2 = Table.ReplaceValue(ReplaceNulls, null, 0, Replacer.ReplaceValue, {"TotalSum"}),

    // Шаг_15_Добавление_колонки_ДУБЛЬ
    сDupColumn = Table.AddColumn(ReplaceNulls2, "ДУБЛЬ", each 
        if [DupCount] > 1 then Text.From([DupCount]) else "-", type text),

    // Шаг_16_Добавление_колонки_СУММА_ПО_ТАБЕЛЬНОМУ
    сSumColumn = Table.AddColumn(сDupColumn, "СУММА ПО ТАБЕЛЬНОМУ", each 
        if [DupCount] > 1 then Number.ToText([TotalSum], ФорматЧисла) else "-", type text),

    // Шаг_17_Добавление_колонки_В_ФОРМАТЕ
    сFormatColumn = Table.AddColumn(сSumColumn, "В ФОРМАТЕ", each 
        if [DupCount] > 1 then Number.ToText([TotalSum], ФорматЧисла) else "-", type text),

    // Шаг_18_Удаление_служебных_колонок
    RemoveTempColumns = Table.RemoveColumns(сFormatColumn, {"DupKey", "DupCount", "TotalSum"}),

    // Шаг_19_Финальный_порядок_колонок_согласно_требованиям
    FinalTable = Table.ReorderColumns(RemoveTempColumns, {
        "MANAGER_PERSON_NUMBER",
        "CONTEST_CODE",
        "TOURNAMENT_CODE",
        "CONTEST_DATE",
        "PLAN_VALUE",
        "FACT_VALUE",
        "priority_type",
        "ДУБЛЬ",
        "СУММА ПО ТАБЕЛЬНОМУ",
        "В ФОРМАТЕ",
        "PLAN_VALUE_число",
        "FACT_VALUE_число"
    })
in
    FinalTable
```

### 🔑 Ключевые отличия моей версии

| Аспект | Моя версия | Оригинальная версия |
|--------|------------|---------------------|
| **Имена переменных** | Английские для параметров (`Params1`, `Processed1`) | Русские (`Параметры_1`, `ОбработанныеДанные_1`) |
| **Имена функций** | Короткие на английском (`GetParamValue`, `SafeTextToNumber`) | Длинные на русском (`ПолучитьПараметрТекст`, `ИзЗапятойВЧисло`) |
| **Явное указание типов** | Везде добавлено `type text`, `type number` | Типы не указаны явно |
| **Структура функций** | Более модульная и читаемая | Аналогичная функциональность |
| **Имена промежуточных колонок** | Короткие (`DupKey`, `DupCount`, `TotalSum`) | Longer (`ВнутреннийКлюч`, `Количество`, `СуммаЗначений`) |

Обе версии полностью соответствуют требованиям и дадут идентичный результат. Моя версия немного более компактна и использует явное приведение типов для лучшей совместимости с Excel 2016.

Источники
