let
    fn_處理權證資料 = (optional SourceTable as nullable table) as table =>
    let
        // 若未傳入表格，直接從 D:\warant 載入最後修改時間最新的 CSV。
        SourceTableResolved =
            if SourceTable <> null then
                SourceTable
            else
                let
                    CsvFiles = Table.SelectRows(
                        Folder.Files("D:\warant"),
                        each Text.Lower([Extension]) = ".csv" and not Text.StartsWith([Name], "~$")
                    ),
                    SortedCsvFiles = Table.Sort(CsvFiles, {{"Date modified", Order.Descending}, {"Name", Order.Descending}}),
                    LatestCsv = if Table.RowCount(SortedCsvFiles) = 0 then error "在 D:\warant 找不到 CSV 檔。" else SortedCsvFiles{0},
                    ImportedCsv = Csv.Document(LatestCsv[Content], [Delimiter = ",", Columns = 18, Encoding = 65001, QuoteStyle = QuoteStyle.Csv])
                in
                    ImportedCsv,

        CleanText = (value) => if value = null then "" else Text.Clean(Text.Trim(Text.From(value))),

        // 0. 定義民國年轉日期函數
        fn_民國轉日期 = (datestr) => 
            if datestr = null or datestr = "" then null else
            let
                CleanStr = Text.Trim(datestr),
                Split = Text.SplitAny(CleanStr, "年月日/"),
                Year = if Number.From(Split{0}) < 1900 then Number.From(Split{0}) + 1911 else Number.From(Split{0}),
                Month = Split{1},
                Day = Split{2},
                DateVal = try Date.From(Text.From(Year) & "/" & Month & "/" & Day) otherwise null
            in
                DateVal,

        // 1. 自動定位標頭 (強健版：自動尋找含有「權證代號」的列)
        CheckHeader = if List.Contains(Table.ColumnNames(SourceTableResolved), "權證代號") then SourceTableResolved 
                      else 
                        let
                            // 搜尋前 10 列，找出哪一列包含 "權證代號"
                            HeaderSearch = Table.FirstN(SourceTableResolved, 10),
                            Rows = Table.ToRows(HeaderSearch),
                            HeaderPos = List.PositionOf(
                                List.Transform(Rows, (r) => List.AnyTrue(List.Transform(r, each Text.Contains(CleanText(_), "權證代號")))),
                                true
                            ),
                            // 如果找不到，預設不跳過；如果找到了，跳過該列之前的列並提升標頭
                            ActualSkip = if HeaderPos = -1 then error "找不到含有「權證代號」的表頭列。" else HeaderPos,
                            Promoted = Table.PromoteHeaders(Table.Skip(SourceTableResolved, ActualSkip), [PromoteAllScalars = true])
                        in
                            Promoted,
        CleanHeaderNames = Table.TransformColumnNames(CheckHeader, each CleanText(_)),

        // 2. 清理數值與移除逗號 (增加更多容錯)
        CleanNumbers = Table.TransformColumns(CleanHeaderNames, {
            {"收盤價", each if _ is text then Number.From(Text.Replace(Text.Replace(_, ",", ""), "－", "0")) else _, type number},
            {"收盤價/指數", each if _ is text then Number.From(Text.Replace(Text.Replace(_, ",", ""), "－", "0")) else _, type number},
            {"履約價格(元)/點數", each if _ is text then Number.From(Text.Replace(Text.Replace(_, ",", ""), "－", "0")) else _, type number},
            {"行使比例", each if _ is text then Number.From(Text.Replace(Text.Replace(_, ",", ""), "－", "0")) else _, type number}
        }),

        // 3. 執行日期轉換
        ConvertDates = Table.TransformColumns(CleanNumbers, {
            {"上市日期", fn_民國轉日期, type date},
            {"履約開始日", fn_民國轉日期, type date},
            {"最後交易日", fn_民國轉日期, type date},
            {"履約截止日", fn_民國轉日期, type date}
        }),

        // 4. 強制設定基本型態
        Step4_Types = Table.TransformColumnTypes(ConvertDates, {
            {"權證代號", type text},
            {"標的代號", type text},
            {"權證類型", type text},
            {"履約方式", type text}
        }),

        // 5. 名稱清理
        Step5_Cleaning = Table.TransformColumns(Step4_Types, {
            {"標的名稱", (x) => Text.Replace(x, "*", ""), type text}
        }),

        // 6. 提取發行商
        Step6_Issuer = Table.AddColumn(Step5_Cleaning, "發行商", each 
            let 
                Issuers = {"元大", "凱基", "群益", "統一", "富邦", "元富", "國票", "永豐", "中信", "台新", "國泰", "兆豐", "華南", "康和", "第一", "亞東", "日盛", "玉山"},
                Name = [權證簡稱],
                Found = List.Select(Issuers, (i) => Text.Contains(Name, i))
            in 
                if List.Count(Found) > 0 then Found{0} else "其他", type text),

        // 7. 計算衍生欄位
        AddMoneyness = Table.AddColumn(Step6_Issuer, "價內/價外", each 
            let S = [#"收盤價/指數"], K = [#"履約價格(元)/點數"], T = [權證類型] in 
            if K = 0 or K = null then null else if T = "認購" then (S/K)-1 else if T = "認售" then 1-(S/K) else null, type number),

        AddRemainingDays = Table.AddColumn(AddMoneyness, "剩餘天數", each 
            if [履約截止日] = null then null else Duration.Days([履約截止日] - Date.From(DateTime.LocalNow())), Int64.Type),

        AddLeverage = Table.AddColumn(AddRemainingDays, "有效槓桿", each 
            if [收盤價] > 0 then ([#"收盤價/指數"] * [行使比例]) / [收盤價] else null, type number),

        AddSubjectType = Table.AddColumn(AddLeverage, "標的類別", each 
            if Text.StartsWith([標的代號], "IX") or [標的名稱] = "臺股指數" then "指數" else "個股", type text),

        // 8. 重新排列
        Final = Table.SelectColumns(AddSubjectType, {
            "權證代號", "權證簡稱", "收盤價", "標的代號", "標的名稱", "收盤價/指數", "履約價格(元)/點數", 
            "價內/價外", "剩餘天數", "有效槓桿", "權證類型", "履約方式", "履約截止日", "上市日期", 
            "行使比例", "發行商", "標的類別"
        })
    in
        Final
in
    fn_處理權證資料
