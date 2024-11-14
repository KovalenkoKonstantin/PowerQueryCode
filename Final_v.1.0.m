let
    // Загрузка первой таблицы
    Source1 = Excel.CurrentWorkbook(){[Name="ПМ"]}[Content],

    // Загрузка второй таблицы
    Source3 = Excel.CurrentWorkbook(){[Name="Работники"]}[Content],

    // Объединение таблиц вертикально
    Combined = Table.Combine({Source1, Source3}),

    //Удаление лишних столбцов
    #"Удаленные столбцы" = Table.RemoveColumns(Combined,{"Подразделение", "Месяц", "Год", "Фамилия Имя Отчество"}),

    //Форматировние даты и суммы
    #"Измененный тип" = Table.TransformColumnTypes(#"Удаленные столбцы",{{"Период (конец месяца)", type date}, {"Сумма", type number}})
in
    #"Измененный тип"