let

    EffDate = Date.ToText(Date.FromText(ABCUSource{0}[Column6]), "YYYYMMDD"),


//Imports person extracts (RCCU and BCU) and applies a source column

    BCUSource = Csv.Document(
        File.Contents("O:\IFRS9_New\Source Extracts\extractResidentialBCU.csv"),
    [Delimiter=",", Columns=110, Encoding=1252, QuoteStyle=QuoteStyle.None]),
    #"Added BCUSource" = Table.AddColumn(BCUSource, "Source", each "BCU"),
    RCCUSource = Table.RemoveRows(
        Csv.Document(
            File.Contents("O:\IFRS9_New\Source Extracts\extractResidentialRCCU.csv"),
        [Delimiter=",", Columns=110, Encoding=1252, QuoteStyle=QuoteStyle.None]), 0),
    #"Added RCCUSource" = Table.AddColumn(RCCUSource, "Source", each "RCCU"),
    ABCUSource = Table.Combine({#"Added BCUSource", #"Added RCCUSource"}),

    MedianCreditScore = List.Median(
        List.Select(
            Table.Column(ABCUSource, "Column72"),
        each _ <> 0 and _ <> 1 and _ <> null)),

    #"Named Columns" = Table.RenameColumns(#"ABCUSource",
        {{"Column1", "RowType"},
        {"Column2", "Credit Union Division"},
        {"Column3", "Branch Name"},
        {"Column4", "Branch Number"},
        {"Column5", "Branch Municipality"},
        {"Column6", "Branch Province"},
        {"Column7", "Branch Postal Code"},
        {"Column8", "Loan ID"},
        {"Column9", "Property Type"},
        {"Column10", "Property Usage Type"},
        {"Column11", "Property Municipality"},
        {"Column12", "Property Province"},
        {"Column13", "Property Postal Code"},
        {"Column14", "Joint Ownership"},
        {"Column15", "Valuation Source"},
        {"Column16", "Property Valuation"},
        {"Column17", "Last Valuation Date"},
        {"Column18", "Also Collateral For"},
        {"Column19", "Origination Date"},
        {"Column20", "Original Loan Amount"},
        {"Column21", "Original Term"},
        {"Column22", "Current Term"},
        {"Column23", "Repricing Date"},
        {"Column24", "Insured Mortgage"},
        {"Column25", "Insurance Company"},
        {"Column26", "Rate Type"},
        {"Column27", "Rate"},
        {"Column28", "Compounding"},
        {"Column29", "Prepayment Options"},
        {"Column30", "Prepayment Percentage"},
        {"Column31", "Allowable Prepayment Amount"},
        {"Column32", "Actual Prepayment Amount"},
        {"Column33", "Actual Prepayment Time Series"},
        {"Column34", "Original Down Payment"},
        {"Column35", "Renewal Loan Amount"},
        {"Column36", "Improvement Amount Borrowed"},
        {"Column37", "Payment Frequency"},
        {"Column38", "Remaining Number of Payments"},
        {"Column39", "Skip Payment Option"},
        {"Column40", "Principal and Interest Amount"},
        {"Column41", "Outstanding Loan Amount"},
        {"Column42", "Amortization"},
        {"Column43", "Renewal"},
        {"Column44", "Last Renewal Date"},
        {"Column45", "Refinancing"},
        {"Column46", "Refinance Amount"},
        {"Column47", "Refinance Reasons"},
        {"Column48", "Refinancing Date"},
        {"Column49", "Loan Stage Override"},
        {"Column50", "Loan Stage Entry Date"},
        {"Column51", "Previous Loan Stage"},
        {"Column52", "Default Reasons"},
        {"Column53", "Cost Recovery"},
        {"Column54", "Delinquency Status"},
        {"Column55", "Delinquency Days overdue"},
        {"Column56", "Delinquency Date"},
        {"Column57", "Joint Borrowers"},
        {"Column58", "Networth of Borrowers"},
        {"Column59", "Significant Risk Increase"},
        {"Column60", "Modification Type"},
        {"Column61", "Carrying Value of Original Loan"},
        {"Column62", "Modification Date"},
        {"Column63", "Modification Gain or Loss"},
        {"Column64", "Stage of Loan Prior to Modification"},
        {"Column65", "Borrower ID1"},
        {"Column66", "Loan Party Type1"},
        {"Column67", "Original Credit Bureau1"},
        {"Column68", "Original Credit Score Type1"},
        {"Column69", "Original Credit Score1"},
        {"Column70", "Current Credit Bureau1"},
        {"Column71", "Current Credit Score Type1"},
        {"Column72", "Current Credit Score1"},
        {"Column73", "Current Credit Score Update Date1"},
        {"Column74", "Borrower ID2"},
        {"Column75", "Loan Party Type2"},
        {"Column76", "Original Credit Bureau2"},
        {"Column77", "Original Credit Score Type2"},
        {"Column78", "Original Credit Score2"},
        {"Column79", "Current Credit Bureau2"},
        {"Column80", "Current Credit Score Type2"},
        {"Column81", "Current Credit Score2"},
        {"Column82", "Current Credit Score Update Date2"},
        {"Column83", "Borrower ID3"},
        {"Column84", "Loan Party Type3"},
        {"Column85", "Original Credit Bureau3"},
        {"Column86", "Original Credit Score Type3"},
        {"Column87", "Original Credit Score3"},
        {"Column88", "Current Credit Bureau3"},
        {"Column89", "Current Credit Score Type3"},
        {"Column90", "Current Credit Score3"},
        {"Column91", "Current Credit Score Update Date3"},
        {"Column92", "Borrower ID4"},
        {"Column93", "Loan Party Type4"},
        {"Column94", "Original Credit Bureau4"},
        {"Column95", "Original Credit Score Type4"},
        {"Column96", "Original Credit Score4"},
        {"Column97", "Current Credit Bureau4"},
        {"Column98", "Current Credit Score Type4"},
        {"Column99", "Current Credit Score4"},
        {"Column100", "Current Credit Score Update Date4"},
        {"Column101", "Borrower ID5"},
        {"Column102", "Loan Party Type5"},
        {"Column103", "Original Credit Bureau5"},
        {"Column104", "Original Credit Score Type5"},
        {"Column105", "Original Credit Score5"},
        {"Column106", "Current Credit Bureau5"},
        {"Column107", "Current Credit Score Type5"},
        {"Column108", "Current Credit Score5"},
        {"Column109", "Current Credit Score Update Date5"},
        {"Column110", "Comment"}}),

    #"Added missing valuation information" = Table.FromRecords(
        Table.TransformRows(#"Named Columns", (r) =>
            Record.TransformFields(r,
                {{"Valuation Source", each if List.Contains({"CMHC", "GENW"}, r[Insurance Company]) then
                    "CMHC"
                else if _ = "" then
                    "OTHR"
                else _},
                {"Property Valuation", each if _ = "" then 0 else _},
                {"Last Valuation Date", each if _ = "" then EffDate else _}}))),
    #"Changed Type" = Table.TransformColumnTypes(#"Added missing valuation information",{{"Loan ID", type text}}),

/*
Below section imports the original loan values, parses the "Original Loan Amount" column for blanks, and adds the NOTEOPENAMT from
WH_ACCTCOMMON if there is a blank.  Once that is done it deletes the NOTEOPENAMT column.  It also adds the TAXRPTFORPERSNBR which is used later in the
code to add new credit score information
*/
    #"Added original loan amounts" = Table.ExpandTableColumn(
        Table.NestedJoin(#"Changed Type", {"Loan ID", "Source"}, MtgOrgLoanAmt, {"ACCTNBR", "SOURCE"}, "Org Amt", JoinKind.LeftOuter),
    "Org Amt", {"NOTEOPENAMT", "TAXRPTFORPERSNBR"}),
    #"Added missing orginal loan amounts" = Table.RemoveColumns(Table.FromRecords(
        Table.TransformRows(#"Added original loan amounts", (r) =>
            Record.TransformFields(r,
                {"Original Loan Amount", each if _ = "" then r[NOTEOPENAMT] else _}))),
    {"NOTEOPENAMT"}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Added missing orginal loan amounts",
        {{"Current Credit Score1", Int64.Type},
        {"Original Credit Score1", Int64.Type},
        {"Current Credit Score Update Date1", Int64.Type}}),
/*
The section below is made up of two parts.  The first imports the most recent credit scores and their respective run dates.  At the same time any
null values in the "Run Date" and "Current Credit Score Update Date1" columns are replaced with a 0 value. Once that is done a <= is compares to
two credit score date columns and if the import is more recent the files information is updated.
*/
    #"Imported new credit scores" = Table.TransformColumns(
        Table.ExpandTableColumn(
            Table.NestedJoin(#"Changed Type1", {"TAXRPTFORPERSNBR", "Source"},
                CreditScores, {"Unique Record ID", "Source"},
            "Credit Info", JoinKind.LeftOuter),
        "Credit Info", {"Fico Risk Score 8 - Score", "Run Date"}, {"Credit Score", "Run Date"}),
    {{"Run Date", each if _ = null then 0 else _},
    {"Current Credit Score Update Date1", each if _ = null then 0 else _}}),

    #"Update credit scores" = Table.FromRecords(
        Table.TransformRows(#"Imported new credit scores", (r) => Record.TransformFields(r,
            {{"Current Credit Bureau1", each if r[Current Credit Score Update Date1] <= r[Run Date] then
                "TransUnion"
            else _},
            {"Current Credit Score Type1", each if r[Current Credit Score Update Date1] <= r[Run Date] then
                "TUFICO"
            else _},
            {"Current Credit Score1", each if r[Current Credit Score Update Date1] <= r[Run Date] then
                r[Credit Score]
            else _},
            {"Current Credit Score Update Date1", each if r[Current Credit Score Update Date1] <= r[Run Date] then
                r[Run Date]
            else _}}))),


//Corrects missing information for borrower1
     #"Fill missing credit score information" = Table.FromRecords(
        Table.TransformRows(#"Update credit scores", (r) => Record.TransformFields(r,
            {{"Original Credit Bureau1", each if _ = "" then "TransUnion" else _},
            {"Original Credit Score Type1", each if _ = "" then "TUCV" else _},
            {"Original Credit Score1", each if List.Contains({null, 0, "000"}, _)
                and List.Contains({null, 0, "000"}, r[Current Credit Score1]) = false then
                    r[Current Credit Score1]
            else if List.Contains({0, null, "000"}, _) = false then _
            else MedianCreditScore},
            {"Current Credit Bureau1", each if _ = "" then "TransUnion" else _},
            {"Current Credit Score Type1", each if _ = "" then "TUCV" else _},
            {"Current Credit Score1", each if List.Contains({null, 0, "000"}, _)
                and List.Contains({null, 0, "000"}, r[Original Credit Score1]) = false then
                    r[Original Credit Score1]
            else if List.Contains({0, null, "000"}, _) = false then _
            else MedianCreditScore},
            {"Current Credit Score Update Date1", each if _ = 0 then EffDate else _}}))),


//The below code removes borrower information for non-primary borrowers if the information is not complete

    #"Remove non-primary borrowers if info missing" = Table.FromRecords(
        Table.TransformRows(#"Fill missing credit score information", (r) => Record.TransformFields(r,
            {{"Borrower ID2", each if r[Original Credit Bureau2] = "" then "" else _},
            {"Loan Party Type2", each if r[Original Credit Bureau2] = "" then "" else _},
            {"Borrower ID3", each if r[Original Credit Bureau3] = "" then "" else _},
            {"Loan Party Type3", each if r[Original Credit Bureau3] = "" then "" else _},
            {"Borrower ID4", each if r[Original Credit Bureau4] = "" then "" else _},
            {"Loan Party Type4", each if r[Original Credit Bureau4] = "" then "" else _},
            {"Borrower ID5", each if r[Original Credit Bureau5] = "" then "" else _},
            {"Loan Party Type5", each if r[Original Credit Bureau5] = "" then "" else _}}))),
    #"Sorted Rows" = Table.Sort(#"Remove non-primary borrowers if info missing",{{"RowType", Order.Descending}}),
/*
The below section adds ECL overrides and places them into the comment field.  This information is stored in .xlsx files in the IFRS9_New folder and
does use VBA for the initial import
*/
    #"Added ECL Overrides" = Table.ExpandTableColumn(
        Table.NestedJoin(#"Sorted Rows", {"Loan ID", "Source"}, ECLOverrides, {"Account Number", "Source"}, "ECL", JoinKind.LeftOuter),
    "ECL", {"ECL Override"}),
    #"Added override to comment field" = Table.FromRecords(
        Table.TransformRows(#"Added ECL Overrides", (r) => Record.TransformFields(r,
            {{"Comment", each if r[ECL Override] <> null then
                "<%ECL=" & Number.ToText(r[ECL Override], "F") & "%>"
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
            {{"Previous Loan Stage", each if r[Stage] <> null then r[Stage] else "TOBEDEFINED"}}))),

    #"Removed Columns" = Table.RemoveColumns(#"Adjusted Previous Loan Stage Column",
        {"Source",
        "ECL Override",
        "TAXRPTFORPERSNBR",
        "Credit Score",
        "Run Date",
        "Stage"}),

    #"Moved header row to top" = Table.Sort(#"Removed Columns",{{"RowType", Order.Descending}})

in
    #"Moved header row to top"
