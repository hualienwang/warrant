let
    // 每次重新整理時，從 D:\warrant 找最後修改時間最新的 twse-stocks-*.csv。
    CsvFiles = Table.SelectRows(
        Folder.Files("D:\warrant"),
        each Text.Lower([Extension]) = ".csv"
            and Text.StartsWith([Name], "twse-stocks-")
            and not Text.StartsWith([Name], "~$")
    ),
    
    // 移除前 2 列（標題列 & 分組標題列），第 3 列自動提升為標頭
    SortedCsvFiles = Table.Sort(CsvFiles, {{"Date modified", Order.Descending}, {"Name", Order.Descending}}),
    LatestCsv = if Table.RowCount(SortedCsvFiles) = 0 then error "在 D:\warant 找不到 twse-stocks-*.csv 檔。" else SortedCsvFiles{0},
    SourceTable = Csv.Document(LatestCsv[Content], [Delimiter = ",", Columns = 16, Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),

    // 清理欄位名稱空白
    ActualSkip=2,
    PromoteHeaders = Table.PromoteHeaders(Table.Skip(SourceTable, ActualSkip), [PromoteAllScalars = true]),
    CleanHeaderNames = Table.TransformColumnNames(PromoteHeaders, each Text.Trim(_)),

    // 文字型態欄位
    SetTextTypes = Table.TransformColumnTypes(CleanHeaderNames, {
        {"證券代號", type text},
        {"證券名稱", type text},
        {"漲跌(+/-)", type text}
    }),

    // 數值型態（先移除千分位逗號）
    CleanCommas = Table.TransformColumnNames(SetTextTypes, each Text.Trim(_)),
    SetNumberTypes = Table.TransformColumns(CleanCommas, {
        {"成交股數", each if _ is text then Number.From(Text.Replace(_, ",", "")) else _, type number},
        {"成交筆數", each if _ is text then Number.From(Text.Replace(_, ",", "")) else _, type number},
        {"成交金額", each if _ is text then Number.From(Text.Replace(_, ",", "")) else _, type number},
        {"開盤價", each if _ is text then Number.From(Text.Replace(_, ",", "")) else _, type number},
        {"最高價", each if _ is text then Number.From(Text.Replace(_, ",", "")) else _, type number},
        {"最低價", each if _ is text then Number.From(Text.Replace(_, ",", "")) else _, type number},
        {"收盤價", each if _ is text then Number.From(Text.Replace(_, ",", "")) else _, type number},
        {"漲跌價差", each if _ is text then Number.From(Text.Replace(_, ",", "")) else _, type number},
        {"最後揭示買價", each if _ is text then Number.From(Text.Replace(_, ",", "")) else _, type number},
        {"最後揭示買量", each if _ is text then Number.From(Text.Replace(_, ",", "")) else _, type number},
        {"最後揭示賣價", each if _ is text then Number.From(Text.Replace(_, ",", "")) else _, type number},
        {"最後揭示賣量", each if _ is text then Number.From(Text.Replace(_, ",", "")) else _, type number},
        {"本益比", each if _ is text then Number.From(Text.Replace(_, ",", "")) else _, type number}
    }),
    Final = Table.SelectColumns(SetNumberTypes, {"證券代號","證券名稱","開盤價","最高價","最低價",
        "收盤價","漲跌價差","成交股數","成交筆數","成交金額","最後揭示買價","最後揭示買量","最後揭示賣價",
        "最後揭示賣量","本益比","漲跌(+/-)"
    })
in
    Final
