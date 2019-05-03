let

    BBSource = Csv.Document(
        File.Contents("O:\IFRS9_New\Source Extracts\extractPersBCU.csv"),
    [Delimiter=",", Columns=100, Encoding=1252, QuoteStyle=QuoteStyle.None]),

    EffDate = Number.From(BBSource{0}[Column6]),
    File = Text.From(List.Max(
        List.Select(
            Table.Column(
                Table.TransformColumns(
                    Folder.Files("O:\IFRS9_New\Loan Stages"),
                {{"Name", each Number.From(Text.Range(_, 0, 8)), type number}}),
            "Name"),
        each _ < EffDate))),

    Source = Csv.Document(
        File.Contents("O:\IFRS9_New\Loan Stages\" & File & ".csv"),
    [Delimiter=",", Columns=4, Encoding=1252, QuoteStyle=QuoteStyle.None]),

    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",
        {{"Grouping", type text},
        {"Branch Number", type text},
        {"Stage", Int64.Type},
        {"Final Loan ECL", type text}}),
    #"Removed Top Rows" = Table.Skip(#"Changed Type",1),

    #"Cleaned Table" = Table.TransformColumns(#"Removed Top Rows",
        {{"Grouping", each try Text.Trim(_) otherwise null, type number}})

in
    #"Cleaned Table"
