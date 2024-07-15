let
    // Подключаемся к базе данных SQL Server
    server = Sql.Database("msk-sql-02", "RKM"),

    // Получаем значение параметра year_id из таблицы "Параметры" в текущей книге Excel
    year_id = Number.ToText(Excel.CurrentWorkbook(){[Name="Параметры"]}[Content][year_id]{0}),

    // Получаем значение параметра company_id из таблицы "Параметры" в текущей книге Excel
    company_id = Number.ToText(Excel.CurrentWorkbook(){[Name="Параметры"]}[Content][id]{0}),

    // Определяем имя хранимой процедуры, которую будем вызывать
    query = "GetEmployeeListShchepetova",

    // Выполняем хранимую процедуру с параметром company_id
    Источник = Value.NativeQuery(
        server,
        "exec " & query & " " & company_id
    )
in
    Источник