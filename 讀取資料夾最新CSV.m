let
    //=================================
    // 讀取資料夾最新CSV
    //=================================
    Source = Folder.Files("D:\warrant"),

    CsvFiles =
        Table.SelectRows(
            Source,
            each Text.EndsWith([Extension], ".csv")
        ),

    SortedFiles =
        Table.Sort(
            CsvFiles,
            {{"Date modified", Order.Descending}}
        ),

    LatestFile =
        SortedFiles{0}[Content],

    CsvData =
        Csv.Document(
            LatestFile,
            [
                Delimiter=",",
                Encoding=65001,
                QuoteStyle=QuoteStyle.Csv
            ]
        ),

    PromoteHeaders =
        Table.PromoteHeaders(
            CsvData,
            [PromoteAllScalars=true]
        ),

    //=================================
    // 欄位型態
    //=================================
    ChangeType =
        Table.TransformColumnTypes(
            PromoteHeaders,
            {
                {"權證代號", type text},
                {"權證簡稱", type text},
                {"標的代號", type text},
                {"標的名稱", type text},
                {"收盤價", type number},
                {"收盤價/指數", type number},
                {"履約價格(元)/點數", type number},
                {"行使比例", type number},
                {"權證類型", type text},
                {"履約方式", type text},
                {"履約截止日", type text},
                {"上市日期", type text}
            }
        ),

    //=================================
    // 民國日期轉西元
    //=================================
    AddExpireDate =
        Table.AddColumn(
            ChangeType,
            "履約截止日_Date",
            each
                try
                    let
                        txt = [履約截止日],
                        y = Number.FromText(Text.BeforeDelimiter(txt,"年")) + 1911,
                        m = Number.FromText(Text.BetweenDelimiters(txt,"年","月")),
                        d = Number.FromText(Text.BetweenDelimiters(txt,"月","日"))
                    in
                        #date(y,m,d)
                otherwise null,
            type date
        ),

    AddListDate =
        Table.AddColumn(
            AddExpireDate,
            "上市日期_Date",
            each
                try
                    let
                        txt = [上市日期],
                        y = Number.FromText(Text.BeforeDelimiter(txt,"年")) + 1911,
                        m = Number.FromText(Text.BetweenDelimiters(txt,"年","月")),
                        d = Number.FromText(Text.BetweenDelimiters(txt,"月","日"))
                    in
                        #date(y,m,d)
                otherwise null,
            type date
        ),

    //=================================
    // 價內價外(%)
    //=================================
    AddMoneyness = Table.AddColumn(Step6_Issuer, "價內/價外", each 
            let S = [#"收盤價/指數"], K = [#"履約價格(元)/點數"], T = [權證類型] in 
            if K = 0 or K = null then null else if T = "認購" then (S/K)-1 else if T = "認售" then 1-(S/K) else null, type number),

    //=================================
    // 剩餘天數
    //=================================
    AddRemainDays =
        Table.AddColumn(
            AddMoneyness,
            "剩餘天數",
            each
                Duration.Days(
                    [履約截止日_Date]
                    -
                    Date.From(DateTime.LocalNow())
                ),
            Int64.Type
        ),

    //=================================
    // 槓桿倍數
    // (無Delta時先作為有效槓桿)
    //=================================
    AddLeverage =
        Table.AddColumn(
            AddRemainDays,
            "有效槓桿",
            each
                try
                    ([收盤價/指數] * [行使比例])
                    / [收盤價]
                otherwise null,
            type number
        ),

    //=================================
    // 發行商
    // 依權證尾碼判斷
    //=================================
    AddIssuer =
        Table.AddColumn(
            AddLeverage,
            "發行商",
            each
                Text.End(
                    Text.Select(
                        [權證代號],
                        {"A".."Z"}
                    ),
                    1
                ),
            type text
        ),

    //=================================
    // 標的類別
    // 可自行擴充
    //=================================
    AddCategory =
        Table.AddColumn(
            AddIssuer,
            "標的類別",
            each
                if Text.Length([標的代號]) = 4
                then "個股"
                else "指數",
            type text
        ),

    //=================================
    // 最終欄位
    //=================================
    SelectColumns =
        Table.SelectColumns(
            AddCategory,
            {
                "權證代號",
                "權證簡稱",
                "收盤價",
                "標的代號",
                "標的名稱",
                "收盤價/指數",
                "履約價格(元)/點數",
                "價內/價外",
                "剩餘天數",
                "有效槓桿",
                "權證類型",
                "履約方式",
                "履約截止日_Date",
                "上市日期_Date",
                "行使比例",
                "發行商",
                "標的類別"
            }
        ),

    RenameColumns =
        Table.RenameColumns(
            SelectColumns,
            {
                {"履約截止日_Date","履約截止日"},
                {"上市日期_Date","上市日期"}
            }
        )

in
    RenameColumns