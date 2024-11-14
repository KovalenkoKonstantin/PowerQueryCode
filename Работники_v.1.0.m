let
    // Источник данных
    Источник = Excel.CurrentWorkbook(){[Name="Работники_полюдно"]}[Content],
    
    // Фильтруем строки, чтобы удалить те, в которых в первом столбце содержится "Итого" и в столбце "Итого" значение равно 0
    FilteredRows = Table.SelectRows(Источник, each [Итого] <> 0 and [Отдел] <> null),
    
    // Удаляю всё лишнее
    RemowedRows = Table.RemoveColumns(FilteredRows,{"№", "Отдел", "employee_accounting_type", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь", "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Итого"}),

    // Добавляем столбец "Месяц" и создаем список из 12 значений месяцев (1..12)
    AddMonths = Table.AddColumn(RemowedRows, "Месяц", each List.Transform({1..12}, (month) => month)),

    // Расширяем столбец "Месяц", чтобы дублировать строки для каждого месяца
    ExpandedMonths = Table.ExpandListColumn(AddMonths, "Месяц"),

    // Меняю название первого столбца на "ФИО"
    Final = Table.RenameColumns(ExpandedMonths, {{"Наименование отдела", "Фамилия Имя Отчество"}}),

    //Преобразую месяц в число
    MonthAsNumber = Table.TransformColumnTypes(Final,{{"Месяц", type number}})
in
    MonthAsNumber