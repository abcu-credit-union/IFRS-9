Attribute VB_Name = "extractPrep"
Sub uploadPrep()
    Dim lRow As String
    Dim actWorkbook As String
    Dim Rng1 As Range
    
    'ActiveWorkbook.Save
    
    'lRow finds the last row of data in a column, this is used to dynamically
    'set the range
    lRow = ActiveWorkbook.ActiveSheet.Cells(1, 1).End(xlDown).Row
    
    'This section determines which extract you are using and saves the .csv in the defined path
    'If you want to change the save location you will make that change here
    If ThisWorkbook.ActiveSheet.Name = "uploadCommercial" Then
        ThisWorkbook.ActiveSheet.Copy
        ActiveWorkbook.SaveAs Filename:="O:\IFRS 9\Extract Work Area\Upload CSVs\uploadCommercial.csv", FileFormat:=xlCSVWindows
        Workbooks("uploadCommercial.csv").Activate
        Set Rng1 = ActiveWorkbook.ActiveSheet.Range("A2: BH" & lRow)
    ElseIf ThisWorkbook.ActiveSheet.Name = "uploadPersLoans" Then
        ThisWorkbook.ActiveSheet.Copy
        ActiveWorkbook.SaveAs Filename:="O:\IFRS 9\Extract Work Area\Upload CSVs\uploadPersLoans.csv", FileFormat:=xlCSVWindows
        Workbooks("uploadPersLoans.csv").Activate
        Set Rng1 = ActiveWorkbook.ActiveSheet.Range("A2: CV" & lRow)
    ElseIf ThisWorkbook.ActiveSheet.Name = "uploadResidental" Then
        ThisWorkbook.ActiveSheet.Copy
        ActiveWorkbook.SaveAs Filename:="O:\IFRS 9\Extract Work Area\Upload CSVs\uploadResidental.csv", FileFormat:=xlCSVWindows
        Workbooks("uploadResidental.csv").Activate
        Set Rng1 = ActiveWorkbook.ActiveSheet.Range("A2:DF" & lRow)
    Else
        GoTo namingError
    End If
    
    'Loops through all cells in the input and deletes the contents if there is
    'an error message
    For Each cell In Rng1
        If Left(cell, 5) = "Error" Then
            cell.Value = ""
        End If
    Next
    
    actWorkbook = ActiveWorkbook.FullName
    
    'Determines the current active sheet and formats the data accordingly
    If ActiveWorkbook.Name = "uploadCommercial.csv" Then
        ActiveWorkbook.Close
        Workbooks.Open (actWorkbook)
        With Workbooks("uploadCommercial.csv").Sheets("uploadCommercial")
            .Range("A2:M" & lRow).NumberFormat = "@"
            .Range("B2").NumberFormat = "0.00"
            .Range("F2:G2").NumberFormat = "0"
            .Range("I2:J2").NumberFormat = "0"
            .Range("I3:I" & lRow).NumberFormat = "0.00"
            .Range("N3:N" & lRow).NumberFormat = "0"
            .Range("AA3:AA" & lRow).NumberFormat = "0.0000"
            .Range("AE3:AG" & lRow).NumberFormat = "0.00"
            .Range("AI3:AK" & lRow).NumberFormat = "0.00"
            .Range("AT3:AT" & lRow).NumberFormat = "0.00"
            .Range("AY3:AY" & lRow).NumberFormat = "0.00"
            .Range("BB3:BB" & lRow).NumberFormat = "0.00"
            .Range("BF3:BF" & lRow).NumberFormat = "0.00"
            .Range("I2:BH2").ClearContents
            .Rows(1).EntireRow.Delete
        End With
    ElseIf ActiveWorkbook.Name = "uploadPersLoans.csv" Then
        ActiveWorkbook.Close
        Workbooks.Open (actWorkbook)
        With Workbooks("uploadPersLoans.csv").Sheets("uploadPersLoans")
            .Range("A2: CV" & lRow).NumberFormat = "@"
            .Range("H2: H" & lRow).NumberFormat = "0"
            .Range("P2: P" & lRow).NumberFormat = "0"
            .Range("R2: R" & lRow).NumberFormat = "0.0000"
            .Range("Y2: Y" & lRow).NumberFormat = "0.00"
            .Range("AE2: AG" & lRow).NumberFormat = "0.00"
            .Range("AI2: AK" & lRow).NumberFormat = "0.00"
            .Range("AR2: AR" & lRow).NumberFormat = "0"
            .Range("AV2: AV" & lRow).NumberFormat = "0.00"
            .Range("AY2: AY" & lRow).NumberFormat = "0.00"
            .Range("BB2: BB" & lRow).NumberFormat = "0.00"
            .Range("BG2: BG" & lRow).NumberFormat = "0"
            .Range("BJ2: BJ" & lRow).NumberFormat = "0"
            .Range("BO2: BO" & lRow).NumberFormat = "0"
            .Range("BR2: BR" & lRow).NumberFormat = "0"
            .Range("BX2: BX" & lRow).NumberFormat = "0"
            .Range("CB2: CB" & lRow).NumberFormat = "0"
            .Range("CH2: CH" & lRow).NumberFormat = "0"
            .Range("CK2: CK" & lRow).NumberFormat = "0"
            .Range("CQ2: CQ" & lRow).NumberFormat = "0"
            .Range("CT2: CT" & lRow).NumberFormat = "0"
            .Range("I2: CV2").ClearContents
            .Rows(1).EntireRow.Delete
        End With
    ElseIf ActiveWorkbook.Name = "uploadResidental.csv" Then
        ActiveWorkbook.Close
        Workbooks.Open (actWorkbook)
        With Workbooks("uploadResidental.csv").Sheets("uploadResidental")
            .Range("A2: DF" & lRow).NumberFormat = "@"
            .Range("H2: H" & lRow).NumberFormat = "0"
            .Range("P2: P" & lRow).NumberFormat = "0.00"
            .Range("T2: T" & lRow).NumberFormat = "0.00"
            .Range("U2: U" & lRow).NumberFormat = "0"
            .Range("V2: V" & lRow).NumberFormat = "0"
            .Range("AA2: AA" & lRow).NumberFormat = "0.0000"
            .Range("AD2: AD" & lRow).NumberFormat = "0.00"
            .Range("AE2: AE" & lRow).NumberFormat = "0.00"
            .Range("AF2: AF" & lRow).NumberFormat = "0.00"
            .Range("AH2: AH" & lRow).NumberFormat = "0.00"
            .Range("AI2: AI" & lRow).NumberFormat = "0.00"
            .Range("AJ2: AJ" & lRow).NumberFormat = "0.00"
            .Range("AK2: AK" & lRow).NumberFormat = "0"
            .Range("AL2: AL" & lRow).NumberFormat = "0"
            .Range("AN2: AN" & lRow).NumberFormat = "0.00"
            .Range("AO2: AO" & lRow).NumberFormat = "0.00"
            .Range("AP2: AP" & lRow).NumberFormat = "0"
            .Range("AT2: AT" & lRow).NumberFormat = "0.00"
            .Range("BA2: BA" & lRow).NumberFormat = "0.00"
            .Range("BC2: BC" & lRow).NumberFormat = "0"
            .Range("BF2: BF" & lRow).NumberFormat = "0.00"
            .Range("BI2: BI" & lRow).NumberFormat = "0.00"
            .Range("BK2: BK" & lRow).NumberFormat = "0.00"
            .Range("BT2: BT" & lRow).NumberFormat = "0"
            .Range("BZ2: BZ" & lRow).NumberFormat = "0"
            .Range("CC2: CC" & lRow).NumberFormat = "0"
            .Range("CI2: CI" & lRow).NumberFormat = "0"
            .Range("CL2: CL" & lRow).NumberFormat = "0"
            .Range("CR2: CR" & lRow).NumberFormat = "0"
            .Range("CU2: CU" & lRow).NumberFormat = "0"
            .Range("DA2: DA" & lRow).NumberFormat = "0"
            .Range("DD2: DD" & lRow).NumberFormat = "0"
            .Rows(1).EntireRow.Delete
        End With
    End If
    ActiveWorkbook.Save
    ActiveWorkbook.Close
Exit Sub
namingError: MsgBox "Please verify the worksheet names are correct before running VBA again"
    
End Sub

