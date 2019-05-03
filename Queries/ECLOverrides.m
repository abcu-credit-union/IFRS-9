let
    Source = Excel.CurrentWorkbook(){[Name="Table10"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Account Number", type text}, {"ECL Override", Int64.Type}, {"Source", type text}})
in
    #"Changed Type"