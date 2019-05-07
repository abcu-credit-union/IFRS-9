let

    EffDate = Date.ToText(Date.FromText(ABCUSource{0}[Column6]), "YYYYMMDD"),
    MedianRate = List.Median(
        List.Select(
            Table.Column(ABCUSource, "Column27"),
        each _ <> 0 and _ <> 1 and _ <> null)),

//Imports person extracts (RCCU and BCU) and applies a source column

    BCUSource = Csv.Document(
        File.Contents("O:\IFRS9_New\Source Extracts\extractCommercialBCU.csv"),
    [Delimiter=",", Columns=60, Encoding=1252, QuoteStyle=QuoteStyle.None]),
    #"Added BCUSource" = Table.AddColumn(BCUSource, "Source", each "BCU"),
    RCCUSource = Table.RemoveRows(
        Csv.Document(
            File.Contents("O:\IFRS9_New\Source Extracts\extractCommercialRCCU.csv"),
        [Delimiter=",", Columns=60, Encoding=1252, QuoteStyle=QuoteStyle.None]),
     0),
    #"Added RCCUSource" = Table.AddColumn(RCCUSource, "Source", each "RCCU"),
    ABCUSource = Table.Combine({#"Added BCUSource", #"Added RCCUSource"}),
    #"Renamed Columns" = Table.RenameColumns(ABCUSource,
        {{"Column1", "RowType"},
        {"Column2", "Credit Union Division"},
        {"Column3", "Branch Name"},
        {"Column4", "Branch Number"},
        {"Column5", "Branch Municipality"},
        {"Column6", "Branch Province"},
        {"Column7", "Branch Postal Code"},
        {"Column8", "Application ID"},
        {"Column9", "Application Authorized Limit"},
        {"Column10", "Security Type"},
        {"Column11", "Security Value"},
        {"Column12", "Security Valuation Date"},
        {"Column13", "Security Property Postal Code"},
        {"Column14", "Loan ID"},
        {"Column15", "Loan Initiation Date"},
        {"Column16", "Loan Maturity Date"},
        {"Column17", "NAICS Code"},
        {"Column18", "Loan Type"},
        {"Column19", "Loan SubType"},
        {"Column20", "Loan Purpose"},
        {"Column21", "Channel"},
        {"Column22", "Collateral"},
        {"Column23", "Amortization"},
        {"Column24", "Initial Term"},
        {"Column25", "Current Term"},
        {"Column26", "Rate Type"},
        {"Column27", "Rate"},
        {"Column28", "Compounding"},
        {"Column29", "Payment Frequency"},
        {"Column30", "Remaining Number of Payments"},
        {"Column31", "Authorized Loan Amount"},
        {"Column32", "Outstanding Loan Amount"},
        {"Column33", "Minimum Periodic Payment"},
        {"Column34", "Prepayment Options"},
        {"Column35", "Prepayment Percentage"},
        {"Column36", "Allowable Prepayment Amount"},
        {"Column37", "Actual Prepayment Amount"},
        {"Column38", "Actual Prepayment Time Series"},
        {"Column39", "Original Loan Risk Rating"},
        {"Column40", "Current Loan Risk Rating"},
        {"Column41", "Current Loan Risk Rating Date"},
        {"Column42", "Loan Stage Override"},
        {"Column43", "Loan Stage Entry Date"},
        {"Column44", "Previous Loan Stage"},
        {"Column45", "Default Reasons"},
        {"Column46", "Cost Recovery"},
        {"Column47", "Delinquency Status"},
        {"Column48", "Delinquency Days Overdue"},
        {"Column49", "Delinquency Date"},
        {"Column50", "Modification Type"},
        {"Column51", "Carrying Value of Original Loan"},
        {"Column52", "Modification Date"},
        {"Column53", "Stage Of Loan Prior to Modification"},
        {"Column54", "Modification Gain or Loss"},
        {"Column55", "Company Name"},
        {"Column56", "Risk Rating"},
        {"Column57", "Risk Rating Last Updated Date"},
        {"Column58", "EBITDA"},
        {"Column59", "Last Annual Review"},
        {"Column60", "Comment"}}),

//The below section goes through security items and cleans the data
     #"Clean security information" = Table.FromRecords(
        Table.TransformRows(#"Renamed Columns", (r) => Record.TransformFields(r,
            {{"Security Type", each if Text.Contains(_,";") then
                Text.BeforeDelimiter(_, ";")
            else _},
            {"Security Value", each if Text.Contains(_,";") then
                Text.BeforeDelimiter(_, ";")
            else _},
            {"Security Valuation Date", each if Text.Contains(_,";") then
                Text.BeforeDelimiter(_, ";")
            else if _ = "" then EffDate
            else _},
            {"Security Property Postal Code", each if Text.Contains(r[Security Type], "PROP") and Text.Contains(_, ";") then
                Text.BeforeDelimiter(_, ";")
            else if Text.Contains(Text.BeforeDelimiter(r[Security Type], ";"), "PROP") and _ = "" then
                r[Branch Postal Code]
            else if Text.Contains(Text.BeforeDelimiter(r[Security Type], ";"), "PROP") = false then
                ""
            else _}}))),
/*
Below section imports the original loan values, parses the "Original Loan Amount" column for blanks, and adds the NOTEOPENAMT from WH_ACCTCOMMON
if there is a blank.  Once that is done it deletes the NOTEOPENAMT column. It also adds the ownername which is used to correct person numbers in the
"Company Name" column
*/
    #"Added original loan amounts" = Table.ExpandTableColumn(
        Table.NestedJoin(#"Clean security information", {"Loan ID", "Source"}, MtgOrgLoanAmt, {"ACCTNBR", "SOURCE"}, "Org Amt", JoinKind.LeftOuter),
    "Org Amt", {"NOTEOPENAMT", "OWNERNAME"}),
/*
Adds missing original loan amounts and corrects Company Name when the primary tax holder is a person
*/
    #"Added missing orginal loan amounts & ownername" = Table.RemoveColumns(Table.FromRecords(
        Table.TransformRows(#"Added original loan amounts", (r) =>
            Record.TransformFields(r,
                {{"Authorized Loan Amount", each if _ = "" then r[NOTEOPENAMT] else _},
                {"Company Name", each try
                    if (Number.From(_) * 1) = Number.From(_) then
                        "Error: " & _ & " is not a valid name"
                    else _
                otherwise if _ = null then "Error: Name is missing"
                else _}}))),
    {"NOTEOPENAMT"}),

    #"Added Last Annual Review" = Table.TransformColumns(#"Added missing orginal loan amounts & ownername",
        {{"Last Annual Review", each if _ = "" then "Error: Last annual review is missing" else _}}),
/*
The below section adds ECL overrides and places them into the comment field.  This information is stored in .xlsx files in the IFRS9_New folder
and does use VBA for the initial import
*/
    #"Added ECL Overrides" = Table.ExpandTableColumn(
        Table.NestedJoin(#"Added Last Annual Review", {"Loan ID", "Source"}, ECLOverrides, {"Account Number", "Source"}, "ECL", JoinKind.LeftOuter),
    "ECL", {"ECL Override"}),
    #"Added override to comment field" = Table.FromRecords(
        Table.TransformRows(#"Added ECL Overrides", (r) => Record.TransformFields(r,
            {{"Allowable Prepayment Amount", each if r[ECL Override] <> null then
                Number.ToText(r[ECL Override], "F")
            else _}}))),

/*
Previous loan stages from the last IFRS9 update need to be inputted into this report.  The PreviousStages table selects to most loan stages
from O:\IFRS9_New\Loan Stages by looking at the file names for dates.  The stages are merged into this table and matches are looked for. Where matches
exist the value is inputted in the Previous Loan Stage column. Where there is no match, TOBEDEFINED is entered.
*/
    #"Added Previous Loan Stages" = Table.ExpandTableColumn(
        Table.NestedJoin(#"Added override to comment field", {"Loan ID", "Branch Number"},
            PreviousStages, {"Grouping", "Branch Number"},
        "Stages", JoinKind.LeftOuter),
        "Stages", {"Stage"}),

    #"Adjusted Previous Loan Stage Column" = Table.FromRecords(
        Table.TransformRows(#"Added Previous Loan Stages", (r) => Record.TransformFields(r,
            {{"Previous Loan Stage", each if r[Stage] <> null then "STAGE" & Text.From(r[Stage]) else "TOBEDEFINED"}}))),

    #"Misc Changes" = Table.FromRecords(
        Table.TransformRows(#"Adjusted Previous Loan Stage Column", (r) => Record.TransformFields(r,
            {{"Company Name", each if Text.Length(_) >= 50 then
                Text.Range(_, 0, 50)
            else _},
            {"Collateral", each if _ <> r[Security Type] and r[Security Type] <> "UNSUP" then
                "Error: " & _ & " does not match " & r[Security Type]
            else _},
            {"Rate", each if _ = "" then "Error: Rate is missing" else _}}))),

    #"Removed Columns" = Table.RemoveColumns(#"Misc Changes",
        {"Source",
        "OWNERNAME",
        "ECL Override",
        "Stage"}),
    #"Sorted Rows" = Table.Sort(#"Removed Columns",{{"RowType", Order.Descending}})


in
    #"Sorted Rows"
