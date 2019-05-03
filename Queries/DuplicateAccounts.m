let
    Source = Excel.CurrentWorkbook(){[Name="Table7"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Loan ID", type text}, {"Source", type text}, {"Translated Loan ID", type text}})
in
    #"Changed Type"