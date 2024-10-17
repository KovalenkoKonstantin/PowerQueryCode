let
    //v.1.1
    // Подключение к SQL-серверу, указывая имя сервера и название базы данных
    server = Sql.Database("msk-sql-02", "RKM"),
    
    // Извлекаем значение year_id из диапазона "Параметры" в текущем рабочем файле Excel
    // Применяем функцию Number.ToText для преобразования в текстовую строку
    year_id = Number.ToText(Excel.CurrentWorkbook(){[Name="Параметры"]}[Content][year_id]{0}),
    
    // Извлекаем значение company_id из диапазона "Параметры" в текущем рабочем файле Excel
    // Применяем функцию Number.ToText для преобразования в текстовую строку
    company_id = Number.ToText(Excel.CurrentWorkbook(){[Name="Параметры"]}[Content][id]{0}),
    
    // Форматируем дату в виде строки 'YYYY-MM-DD', представляющей 31 июля текущего года
    // Получаем year_id и добавляем строку "-07-31" для обозначения конкретной даты
    dat = "'" & Number.ToText(Excel.CurrentWorkbook(){[Name="Параметры"]}[Content][year_id]{0}) & "-07-31'",
    
    // Выполняем SQL-запрос к базе данных с помощью вызова функции Value.NativeQuery
    Источник = Value.NativeQuery(
        server,
        "exec GetEmployeeChangesRefreshAlt " & company_id & ", " & year_id & ", " & dat & ""
    )
in
    // Возвращаем результат выполнения запроса (переменную "Источник")
    Источник