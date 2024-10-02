let
    // Подключаемся к базе данных SQL Server
    server = Sql.Database("msk-sql-02", "RKM"),

    // Получаем значение параметра start_year_number из таблицы "Параметры" в текущей книге Excel
    start_year_number = Number.ToText(Excel.CurrentWorkbook(){[Name="Параметры"]}[Content][start_year_number]{0}),
	
	// Получаем значение параметра end_year_number из таблицы "Параметры" в текущей книге Excel
    end_year_number = Number.ToText(Excel.CurrentWorkbook(){[Name="Параметры"]}[Content][end_year_number]{0}),

    // Получаем значение параметра company_id из таблицы "Параметры" в текущей книге Excel
    company_id = Number.ToText(Excel.CurrentWorkbook(){[Name="Параметры"]}[Content][company_id]{0}),

    // Получаем данные из умной таблицы "Сотрудники"
    employeesTable = Excel.CurrentWorkbook(){[Name="Сотрудники"]}[Content],    
    // Убираем строку с "Итого"
    filteredEmployeesTable = Table.SelectRows(employeesTable, each [Сотрудник] <> "Итого"),
    // Получаем только список сотрудников
    employeeNamesList = filteredEmployeesTable[Сотрудник],    
    // Конкатенируем имена сотрудников в одну строку, разделяя запятыми
    employeeNames = Text.Combine(employeeNamesList, ","),

    // Определяем имя хранимой процедуры, которую будем вызывать
    query = "GetSalaryBudgetRefresh",

    // Формируем вызов хранимой процедуры с параметрами company_id, start_year_number, end_year_number, и employee_names
    Источник = Value.NativeQuery(
        server,
        "exec " & query & " " & company_id & ", " & start_year_number & ", " & end_year_number & ", '" & employeeNames & "'"
    )
in
    Источник