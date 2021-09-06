Sub UpdateCOB()

Dim PivotName As Variant


' Update Pivot Tables on sheet "Status and Loss Change"

Worksheets("Status & Loss Change").Activate

If Not ActiveSheet.AutoFilter Is Nothing Then Cells.AutoFilter

    For Each PivotName In Array("CurrentStatus", "CurrentDirect", "CurrentAdj", "Category", "Risk")
        ActiveSheet.PivotTables(PivotName).PivotFields("[COB Date].[COB Date].[COB Date]") _
        .CurrentPageName = Sheets("Summary").Range("G2")
    Next


    For Each PivotName In Array("PreviousStatus", "PreviousDirect", "PreviousAdj")
        ActiveSheet.PivotTables(PivotName).PivotFields("[COB Date].[COB Date].[COB Date]") _
        .CurrentPageName = Sheets("Summary").Range("E2")
    Next

'Autofill Column E to J
Range("E9:J" & Cells(Rows.Count, "A").End(xlUp).Row).ClearContents
Range("E8:J8").AutoFill Destination:=Range("E8:J" & Cells(Rows.Count, "A").End(xlUp).Row - 1)


'Autofill Column AG
Range("AG13:AG" & Cells(Rows.Count, "T").End(xlUp).Row).ClearContents
Range("AG12").AutoFill Destination:=Range("AG12:AG" & Cells(Rows.Count, "X").End(xlUp).Row - 1)



' Update Pivot Tables on sheet "Loss Change Detail"

Worksheets("Loss Change Detail").Activate

If Not ActiveSheet.AutoFilter Is Nothing Then Cells.AutoFilter

    For Each PivotName In Array("CurrentLoss", "CurrentAdj")
        ActiveSheet.PivotTables(PivotName).PivotFields("[COB Date].[COB Date].[COB Date]") _
       .CurrentPageName = Sheets("Summary").Range("G2")
    Next


    For Each PivotName In Array("PreviousLoss", "PreviousAdj")
        ActiveSheet.PivotTables(PivotName).PivotFields("[COB Date].[COB Date].[COB Date]") _
        .CurrentPageName = Sheets("Summary").Range("E2")
    Next

'Autofill Column E to F
Range("E10:F" & Cells(Rows.Count, "A").End(xlUp).Row).ClearContents
Range("E9:F9").AutoFill Destination:=Range("E9:F" & Cells(Rows.Count, "A").End(xlUp).Row - 1)


'Autofill Column L to M
Range("L10:M" & Cells(Rows.Count, "H").End(xlUp).Row).ClearContents
Range("L9:M9").AutoFill Destination:=Range("L9:M" & Cells(Rows.Count, "H").End(xlUp).Row - 1)

'Autofill Column T to U
Range("T10:U" & Cells(Rows.Count, "P").End(xlUp).Row).ClearContents
Range("T9:U9").AutoFill Destination:=Range("T9:U" & Cells(Rows.Count, "P").End(xlUp).Row - 1)


'Autofill Column AA to AB
Range("AA10:AB" & Cells(Rows.Count, "W").End(xlUp).Row).ClearContents
Range("AA9:AB9").AutoFill Destination:=Range("AA9:AB" & Cells(Rows.Count, "W").End(xlUp).Row - 1)


' Update Pivot Tables on sheet "Reference Stats"
Worksheets("Reference Stats").Activate

If Not ActiveSheet.AutoFilter Is Nothing Then Cells.AutoFilter

    For Each PivotName In Array("PivotTable1", "PivotTable2", "PivotTable4", "PivotTable5")
        ActiveSheet.PivotTables(PivotName).PivotFields("[COB Date].[COB Date].[COB Date]") _
        .CurrentPageName = Sheets("Summary").Range("G2")
    Next
    
End Sub

___________________________________________________________________________________________________

Sub WeeklyNew()

Sheets("Weekly New ORI").Columns("A:Q").EntireColumn.Clear

' New Direct
Select Case Sheets("Summary").Range("B7")
    Case Is <> 0
        Sheets("Report Master").Range("A1:O16").Copy Sheets("Weekly New ORI").Range("A1")
        Sheets("Status & Loss Change").Select
        If Not ActiveSheet.AutoFilter Is Nothing Then Cells.AutoFilter
        Range("G7").AutoFilter Field:=7, Criteria1:="New"
        Range("H7").AutoFilter Field:=8, Criteria1:="Direct"
        Range("A8:A" & Cells(Rows.Count, "A").End(xlUp).Row).SpecialCells(xlCellTypeVisible).Copy
        Sheets("Weekly New ORI").Select
        Range("A" & Cells(Rows.Count, "A").End(xlUp).Row).PasteSpecial

Range("A16", Range("A16").Offset(Sheets("Summary").Range("B7") - 1, 0)).Select
With Selection
        .NumberFormat = "General"
        .Value = .Value
        .HorizontalAlignment = xlLeft
        End With

    Case Else
        Sheets("Report Master").Range("A1:A7").Copy Sheets("Weekly New ORI").Range("A1")
        
End Select



' New Adjustment
Select Case Sheets("Summary").Range("B8")
    Case Is <> 0
        Sheets("Report Master").Range("A18:O20").Copy Sheets("Weekly New ORI").Range("A" & Cells(Rows.Count, "A").End(xlUp).Row).Offset(2, 0)
        Sheets("Status & Loss Change").Select
        If Not ActiveSheet.AutoFilter Is Nothing Then Cells.AutoFilter
        Range("G7").AutoFilter Field:=7, Criteria1:="New"
        Range("H7").AutoFilter Field:=8, Criteria1:="Adjustment"
        Range("A8:A" & Cells(Rows.Count, "A").End(xlUp).Row).SpecialCells(xlCellTypeVisible).Copy
        Sheets("Weekly New ORI").Select
        Range("A" & Cells(Rows.Count, "A").End(xlUp).Row).PasteSpecial

With Selection
        .NumberFormat = "General"
        .Value = .Value
        .HorizontalAlignment = xlLeft
        End With

    Case Else
        Sheets("Report Master").Range("A1").Copy Sheets("Weekly New ORI").Range("A1")
        
End Select


' Indirect/NearMiss
Select Case Sheets("Summary").Range("B9")
    Case Is <> 0
        Sheets("Report Master").Range("A22:N24").Copy Sheets("Weekly New ORI").Range("A" & Cells(Rows.Count, "A").End(xlUp).Row).Offset(2, 0)
        Sheets("Status & Loss Change").Select
        If Not ActiveSheet.AutoFilter Is Nothing Then Cells.AutoFilter
        Range("G7").AutoFilter Field:=7, Criteria1:="New"
        Range("H7").AutoFilter Field:=8, Criteria1:="Indirect/Near Miss"
        Range("A8:A" & Cells(Rows.Count, "A").End(xlUp).Row).SpecialCells(xlCellTypeVisible).Copy
        Sheets("Weekly New ORI").Select
        Range("A" & Cells(Rows.Count, "A").End(xlUp).Row).PasteSpecial

With Selection
        .NumberFormat = "General"
        .Value = .Value
        .HorizontalAlignment = xlLeft
        End With

    Case Else
        Sheets("Report Master").Range("A1").Copy Sheets("Weekly New ORI").Range("A1")
        
End Select



End Sub

_________________________________________________________________________________________________

