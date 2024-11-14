let
    // Загрузка первой таблицы
    Source = Excel.CurrentWorkbook(){[Name="НС"]}[Content],

    //Удаление лишних столбцов
    #"Удаленные столбцы" = Table.RemoveColumns(Source,{"Подразделение", "Месяц", "Год"}),

    //Форматировние даты и суммы
    #"Измененный тип" = Table.TransformColumnTypes(#"Удаленные столбцы",{{"Период (конец месяца)", type date}, {"Сумма", type number}})
in
    #"Измененный тип"