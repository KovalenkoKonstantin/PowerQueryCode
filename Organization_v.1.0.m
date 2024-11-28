let
    //v.1.0
    // Подключаемся к базе данных SQL Server
    server = Sql.Database("msk-sql-02", "RKM"),

    // Определяем имя хранимой процедуры, которую будем вызывать
    query = "Organization_v_1_0",

    // Выполняем хранимую процедуру с параметрами company_id и year_number
    Источник = Value.NativeQuery(
        server,
        "exec " & query & ""
    )
in
    Источник