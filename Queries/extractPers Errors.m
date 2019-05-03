let
//Calculates the median credit score and pulls the current date from the extract
    EffDate = Date.FromText(ABCUSource{0}[Column6]),
    MedianCreditScore = List.Median(Table.Column(ABCUSource, "Column62")),

//Imports person extracts (RCCU and BCU) and applies a source column
    BCUSource = Csv.Document(
        File.Contents("O:\IFRS9_New\Source Extracts\extractPersBCU.csv"),[Delimiter=",", Columns=100, Encoding=1252, QuoteStyle=QuoteStyle.None]),
    #"Added BCUSource" = Table.AddColumn(BCUSource, "Source", each "BCU"),
    RCCUSource = Table.RemoveRows(
        Csv.Document(
            File.Contents("O:\IFRS9_New\Source Extracts\extractPersRCCU.csv"),
        [Delimiter=",", Columns=100, Encoding=1252, QuoteStyle=QuoteStyle.None]), 0),
    #"Added RCCUSource" = Table.AddColumn(RCCUSource, "Source", each "RCCU"),

//Combines the two extracts and names the 100+ columns
    ABCUSource = Table.Combine({#"Added BCUSource", #"Added RCCUSource"}),
    #"Renamed Columns" = Table.RenameColumns(ABCUSource,
        {{"Column1", "RowType"},
        {"Column2", "Credit Union Division"},
        {"Column3", "Branch Name"},
        {"Column4", "Branch Number"},
        {"Column5", "Branch Municipality"},
        {"Column6", "Branch Province"},
        {"Column7", "Branch Postal Code"},
        {"Column8", "Loan ID"},
        {"Column9", "Origination Date"},
        {"Column10", "Maturity Date"},
        {"Column11", "Loan Type"},
        {"Column12", "Loan SubType"},
        {"Column13", "Loan Purpose"},
        {"Column14", "Channel"},
        {"Column15", "Collateral"},
        {"Column16", "Term"},
        {"Column17", "Rate Type"},
        {"Column18", "Rate"},
        {"Column19", "Property Type"},
        {"Column20", "Property Usage Type"},
        {"Column21", "Property Municipality"},
        {"Column22", "Property Province"},
        {"Column23", "Property Postal Code"},
        {"Column24", "Valuation Source"},
        {"Column25", "Property Valuation"},
        {"Column26", "Last Valuation Date"},
        {"Column27", "Also Collateral For"},
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
        {"Column39", "Loan Stage Override"},
        {"Column40", "Loan Stage Entry Date"},
        {"Column41", "Previous Loan Stage"},
        {"Column42", "Default Reasons"},
        {"Column43", "Delinquency Status"},
        {"Column44", "Delinquency Days Overdue"},
        {"Column45", "Delinquency Date"},
        {"Column46", "Has Guarantor"},
        {"Column47", "Joint Borrowers"},
        {"Column48", "Networth of Borrowers"},
        {"Column49", "Significant Risk Increase"},
        {"Column50", "Modification Type"},
        {"Column51", "Carrying Value of Orginal Loan"},
        {"Column52", "Modification Date"},
        {"Column53", "Stage of Loan Prior to Modification"},
        {"Column54", "Modification Gain or Loss"},
        {"Column55", "Borrower ID1"},
        {"Column56", "Loan Party Type1"},
        {"Column57", "Original Credit Bureau1"},
        {"Column58", "Original Credit Score Type1"},
        {"Column59", "Original Credit Score1"},
        {"Column60", "Current Credit Bureau1"},
        {"Column61", "Current Credit Score Type1"},
        {"Column62", "Current Credit Score1"},
        {"Column63", "Current Credit Score Update Date1"},
        {"Column64", "Borrower ID2"},
        {"Column65", "Loan Party Type2"},
        {"Column66", "Original Credit Bureau2"},
        {"Column67", "Original Credit Score Type2"},
        {"Column68", "Original Credit Score2"},
        {"Column69", "Current Credit Bureau2"},
        {"Column70", "Current Credit Score Type2"},
        {"Column71", "Current Credit Score2"},
        {"Column72", "Current Credit Score Update Date2"},
        {"Column73", "Borrower ID3"}, {"Column74", "Loan Party Type3"},
        {"Column75", "Original Credit Bureau3"},
        {"Column76", "Original Credit Score Type3"},
        {"Column77", "Original Credit Score3"},
        {"Column78", "Current Credit Bureau3"},
        {"Column79", "Current Credit Score Type3"},
        {"Column80", "Current Credit Score3"},
        {"Column81", "Current Credit Score Update Date3"},
        {"Column82", "Borrower ID4"},
        {"Column83", "Loan Party Type4"},
        {"Column84", "Original Credit Bureau4"},
        {"Column85", "Original Credit Score Type4"},
        {"Column86", "Original Credit Score4"},
        {"Column87", "Current Credit Bureau4"},
        {"Column88", "Current Credit Score Type4"},
        {"Column89", "Current Credit Score4"},
        {"Column90", "Current Credit Score Update Date4"},
        {"Column91", "Borrower ID5"},
        {"Column92", "Loan Party Type5"},
        {"Column93", "Original Credit Bureau5"},
        {"Column94", "Original Credit Score Type5"},
        {"Column95", "Original Credit Score5"},
        {"Column96", "Current Credit Bureau5"},
        {"Column97", "Current Credit Score Type5"},
        {"Column98", "Current Credit Score5"},
        {"Column99", "Current Credit Score Update Date5"},
        {"Column100", "Comment"}}),

/*LOCs with the major code CML are not populating maturity dates and remaining number of payments. Account types are added to the query that identify loans as AOD, Loan, or HELOC which is then used to identify the loans that need maturity's manually added.  After that is
completed the next two steps add maturity dates and remaining number of payments to the exisiting columns where a blank value exists*/

    #"Added Account Types" = Table.ExpandTableColumn(
        Table.NestedJoin(#"Renamed Columns", {"Loan ID", "Source"}, AccountTypes, {"ACCTNBR", "SOURCE"}, "Type", JoinKind.LeftOuter),
    "Type", {"Type", "TAXRPTFORPERSNBR"}),
    #"Corrected missing maturity dates, terms and loan purpose" = Table.FromRecords(
        Table.TransformRows(#"Added Account Types", (r) => Record.TransformFields(r,
            {{"Maturity Date", each if r[Type] = "HELOC" and _ = "" then
                "Error: Missing maturity date"
            else _},
            {"Term", each if r[Type] = "HELOC" and _ = "" then
                "Error: Missing term"
            else _},
            {"Loan Purpose", each if r[Type] = "HELOC" then
                "Error: Missing loan purpose"
            else _}}))),
    #"Corrected missing remaining payments" = Table.FromRecords(
        Table.TransformRows(#"Corrected missing maturity dates, terms and loan purpose", (r) => Record.TransformFields(r,
            {"Remaining Number of Payments", each if r[Type] = "HELOC" and _ = "" then
                "Error: Missing number of payments"
            else _}))),

//Some loans have origination date and maturity date but are failing to populate the term.  This calculates out the term and fills the field

    #"Added missing terms" = Table.FromRecords(
        Table.TransformRows(#"Corrected missing remaining payments", (r) => Record.TransformFields(r,
            {"Term", each if _ = "" and r[Maturity Date] <> "" and r[Origination Date] <> "" then
                "Error: Missing Term"
            else _}))),

//The following steps identify and correct errors in the extract.  These errors include missing security information, missing credit score information, etc. The column names can be used to identify the type of error being corrected or identified in the following in the
//following code

    #"Fill missing security information" = Table.FromRecords(
        Table.TransformRows(#"Added missing terms", (r) => Record.TransformFields(r,
            {{"Property Type", each if r[Collateral] = "PROP" and _ = "" then
                "Error: Property Type Required"
            else _},
            {"Property Usage Type", each if r[Collateral] = "PROP" and _ = "" then
                "Error: Missing usage type"
            else _},
            {"Property Municipality", each if r[Collateral] = "PROP" and _ = "" then
                "Error: Missing municipality"
            else _},
            {"Property Province", each if r[Collateral] = "PROP" and _ = "" then
                "Error: Missing province"
            else _},
            {"Property Postal Code", each if r[Collateral] = "PROP" and Text.Length( _) < 7 then
                "Error: Missing postal code"
            else _},
            {"Valuation Source", each if r[Collateral] = "PROP" and _ = "" then
                "Error: missing valuation source"
            else _},
            {"Property Valuation", each if r[Collateral] = "PROP" and _ = "" then
                "Error: Missing valuation"
            else Number.From(_)}}))),
    #"Change loan type based on security" = Table.FromRecords(
        Table.TransformRows(#"Fill missing security information", (r) => Record.TransformFields(r,
            {"Loan Purpose", each if r[Collateral] = "PROP" and r[Loan SubType] = "REVLV" then "Error: " & _ & " is wrong loan type"
        else _}))),

    #"Changed Type" = Table.TransformColumnTypes(#"Change loan type based on security",
        {{"Current Credit Score1", Int64.Type},
        {"Original Credit Score1", Int64.Type},
        {"Borrower ID1", Int64.Type},
        {"Source", type text},
        {"Current Credit Score Update Date1", Int64.Type},
        {"Outstanding Loan Amount", type number}}),
    #"Corrected misc errors" = Table.FromRecords(
        Table.TransformRows(#"Changed Type", (r) => Record.TransformFields(r,
            {{"Delinquency Days Overdue", each if Text.Length(_) >3 then "Error: Delinquency can't be" & _ else _},
            {"Outstanding Loan Amount", each if _ < 0 then "Error: Outstanding loan amount cant be negative" else _}}))),

//The section below is made up of two parts.  The first imports the most recent credit scores and their respective run dates.  At the same time any null values in the "Run Date" and "Current Credit Score Update Date1" columns are replaced with a 0 value.
//Once that is done a <= is compares to two credit score date columns and if the import is more recent the files information is updated.

    #"Imported new credit scores" = Table.TransformColumns(
        Table.ExpandTableColumn(
            Table.NestedJoin(#"Corrected misc errors", {"TAXRPTFORPERSNBR", "Source"}, CreditScores, {"Unique Record ID", "Source"}, "Credit Info", JoinKind.LeftOuter),
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


    #"Fill missing credit score information" = Table.FromRecords(
        Table.TransformRows(#"Imported new credit scores", (r) => Record.TransformFields(r,
            {{"Original Credit Bureau1", each if _ = "" then "Error: missing bureau" else _},
            {"Original Credit Score Type1", each if _ = "" then "Error: Missing score type" else _},
            {"Original Credit Score1", each if List.Contains({null, 0}, _) and List.Contains({null, 0}, r[Current Credit Score1]) = false then
                "Error: Missing score"
            else if _ <> null then _
            else "Error: Missing score"},
            {"Current Credit Bureau1", each if _ = "" then "Error: Missing bureau" else _},
            {"Current Credit Score Type1", each if _ = "" then "Error: Missing score type" else _},
            {"Current Credit Score1", each if List.Contains({null, 0}, _) and List.Contains({null, 0}, r[Original Credit Score1]) = false then
                "Error: Missing score"
            else if _ <> null then _
            else "Error: Missing score"},
            {"Current Credit Score Update Date1", each if _ = 0 then "Error: Missing score update date" else _}}))),

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

    #"Removed user created columns" = Table.RemoveColumns(#"Remove non-primary borrowers if info missing",{"Source", "Type", "TAXRPTFORPERSNBR", "Credit Score", "Run Date"}),
    #"Moved header row to top" = Table.Sort(#"Removed user created columns",{{"RowType", Order.Descending}}),
    #"Replaced Value" = Table.ReplaceValue(#"Moved header row to top","","TOBEDEFINED",Replacer.ReplaceValue,{"Previous Loan Stage"})
in
    #"Replaced Value"
