Sub importECLOverides()
    Dim extractDate As String
    Dim BCUSource As String
    Dim RCCUSource As String
    Dim lRow As String 'Generates the last row used in Rng1
    Dim lRow2 As String 'Generates the last row used in Rng2
    Dim Rng1 As Range 'Specifies the last row with data in the BCU & RCCU Workbooks
    Dim Rng2 As Range 'This range is used to delete the old data

'Sets initial variables and pauses screen updating.  The date provided in the extracts is used to determine what
'folder to open.  If the folder name is different then the extract date and error will be thrown

    Application.ScreenUpdating = False

'lRow2 and Rng2 are used to set the range of old data so it can be removed later.
    lRow2 = ThisWorkbook.Sheets("ECL Overrides").Cells(Rows.Count, "A").End(xlUp).Row
    Set Rng2 = ThisWorkbook.Sheets("ECL Overrides").Range("A2:C" & lRow2)

    Let extractDate = ThisWorkbook.Sheets("uploadPersLoans").Range("F2").Value
    BCUSource = "O:\IFRS9_New\ECL Overrides\" & extractDate & "\BCU.xlsx"
    RCCUSource = "O:\IFRS9_New\ECL Overrides\" & extractDate & "\RCCU.xlsx"

'This section opens the RCCU Specific Allowances workbook and dynamically defines the range
'that contains data using lRow
    Set RCCUData = Workbooks.Open(RCCUSource)
    lRow = RCCUData.Sheets(1).Cells(500, 1).End(xlUp).Row
    Set Rng1 = RCCUData.Sheets(1).Range("A1:A" & lRow)

'The for loop goes through the range of data in the RCCU workbook and adds the account numbers, specific provision
'and source to the Upload Prep Tool
    For Each cell In Rng1
        If cell <> "" And IsNumeric(cell.Value) Then
            lRow2 = ThisWorkbook.Sheets("ECL Overrides").Cells(Rows.Count, "A").End(xlUp).Offset(1, 0).Row
            With ThisWorkbook.Sheets("ECL Overrides")
                .Cells(lRow2, 1).Value = cell.Value
                .Cells(lRow2, 2).Value = cell.Offset(0, 11)
                .Cells(lRow2, 3).Value = "RCCU"
            End With
        End If
    Next

    Workbooks("RCCU.xlsx").Close savechanges:=False

'Changes lRow and Rng1 from the RCCU workbook to the BCU workbook
    Set BCUData = Workbooks.Open(BCUSource)
    lRow = BCUData.Sheets(1).Cells(500, 1).End(xlUp).Row
    Set Rng1 = BCUData.Sheets(1).Range("A1:A" & lRow)

'The for loop goes through the range of data in the BCU workbook and adds the account numbers, specific provision
'and source to the Upload Prep Tool
    For Each cell In Rng1
        If cell <> "" And IsNumeric(cell.Value) Then
            lRow2 = ThisWorkbook.Sheets("ECL Overrides").Cells(Rows.Count, "A").End(xlUp).Offset(1, 0).Row
            With ThisWorkbook.Sheets("ECL Overrides")
                .Cells(lRow2, 1).Value = cell.Value
                .Cells(lRow2, 2).Value = cell.Offset(0, 11)
                .Cells(lRow2, 3).Value = "BCU"
            End With
        End If
    Next

    Workbooks("BCU.xlsx").Close savechanges:=False

'Deletes the range of old data
    Rng2.EntireRow.Delete shift:=xlUp

    Application.ScreenUpdating = True


End Sub
