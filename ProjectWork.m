// let
//     Источник = Sql.Database("msk-sql-02", "RKM", [Query="execute GetProJectWork 9, 2024;"])
// in
//     Источник

let
    // Подключаемся к базе данных SQL Server
    server = Sql.Database("msk-sql-02", "RKM"),

    // Получаем значение параметра start_year_number из таблицы "Параметры" в текущей книге Excel
    start_year_number = Number.ToText(Excel.CurrentWorkbook(){[Name="Параметры"]}[Content][start_year_number]{0}),
	
	// Получаем значение параметра end_year_number из таблицы "Параметры" в текущей книге Excel
    end_year_number = Number.ToText(Excel.CurrentWorkbook(){[Name="Параметры"]}[Content][end_year_number]{0}),

    // Получаем значение параметра company_id из таблицы "Параметры" в текущей книге Excel
    company_id = Number.ToText(Excel.CurrentWorkbook(){[Name="Параметры"]}[Content][company_id]{0}),

    // Определяем имя хранимой процедуры, которую будем вызывать
    query = "GetProJectWork",

    // Выполняем хранимую процедуру с параметрами company_id и year_number
    Источник = Value.NativeQuery(
        server,
        "exec " & query & " " & company_id & ", " & start_year_number & "," & end_year_number & ""
    )
in
    Источник