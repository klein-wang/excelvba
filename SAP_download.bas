
Sub Sales_Plant()
'
' Sales_Plant Macro
'

'
    Sheets("by Sales Org").Select
    Range("I2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Specification ID TBD").Select
    Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Dim Count As Integer
    Count = Sheets("by Sales Org").Range("I20000").End(xlUp).Row
        
    Sheets("by Plant").Select
    Range("I2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Specification ID TBD").Select
    Range("A" & Count + 1).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("A:A").Select
    Application.CutCopyMode = False
    ActiveSheet.Range("$A$1:$C$10160").RemoveDuplicates Columns:=1, _
        Header:=xlYes
End Sub
Sub Macro_Haz()
'
' Macro_Haz Macro
'

'
    Sheets("Haz CAS").Select
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Specification ID TBD").Select
    Range("B2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Dim Count As Integer
    Count = Sheets("Haz CAS").Range("I2000").End(xlUp).Row
    
    Sheets("Haz SYN").Select
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Specification ID TBD").Select
    Range("B" & Count + 1).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("B:B").Select
    Application.CutCopyMode = False
    ActiveSheet.Range("$B$1:$B$10160").RemoveDuplicates Columns:=1, Header:= _
        xlYes
        
        
    Dim Count2 As Integer
    Count2 = Sheets("Specification ID TBD").Range("A20000").End(xlUp).Row
    
    Range("C2").FormulaR1C1 = "=NOT(ISERROR(VLOOKUP(RC[-2],C[-1],1,0)))"
    Range("C2").Select
    Selection.AutoFill Destination:=Range("C2:C" & Count2)
        

End Sub
Sub Macro_NA_Haz()
'
' Macro_NA_Haz Macro
'

'
    Sheets("Sheet1").Select
    Application.CutCopyMode = False
    Selection.ClearContents

    Sheets("Specification ID TBD").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$J$395").AutoFilter Field:=3, Criteria1:="FALSE"
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Sheets("Sheet1").Select
    Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    Sheets("Specification ID TBD").Select
    Application.CutCopyMode = False
    ActiveSheet.ShowAllData
   
    Sheets("Sheet1").Select
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Specification ID TBD").Select
    Range("D2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub

Sub SaveFiles()

    Path = ActiveWorkbook.Path
    Dim a, b As String
    a = Format(Date, "yyyymmdd") 'µ±Ç°ÄêÔÂÈÕ
    b = Format(Time, "hhmm") 'µ±Ç°Ê±¼ä
    

    Sheets("by Sales Org").Select
    Sheets("by Sales Org").Copy
    ActiveWorkbook.SaveAs Filename:=Path & "\files\Basic Info_Sorg_" + a + b + ".xlsx" _
        , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWindow.Close     '¹Ø±ÕÒÑ¾­Éú³ÉµÄÎÄ¼þ

    Sheets("by Plant").Select
    Sheets("by Plant").Copy
    ActiveWorkbook.SaveAs Filename:=Path & "\files\Basic Info_Plant_" + a + b + ".xlsx" _
        , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWindow.Close     '¹Ø±ÕÒÑ¾­Éú³ÉµÄÎÄ¼þ

    Sheets("Haz CAS").Select
    Sheets("Haz CAS").Copy
    ActiveWorkbook.SaveAs Filename:=Path & "\files\HAZ_CAS_" + a + b + ".xlsx" _
        , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWindow.Close     '¹Ø±ÕÒÑ¾­Éú³ÉµÄÎÄ¼þ

    Sheets("Haz SYN").Select
    Sheets("Haz SYN").Copy
    ActiveWorkbook.SaveAs Filename:=Path & "\files\HAZ_SYN_" + a + b + ".xlsx" _
        , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWindow.Close     '¹Ø±ÕÒÑ¾­Éú³ÉµÄÎÄ¼þ

    Sheets("Non-Haz").Select
    Sheets("Non-Haz").Copy
    Range("A1").Select
        For r = 1 To 20
            If Cells(1, r) <> "Spec." And Cells(1, r) <> "Data record" And Cells(1, r) <> "Remarks" And Cells(1, r) <> "" Then
                Columns(r).Delete
                r = r - 1
            ElseIf Cells(1, r) = "" And Cells(2, r) <> "" Then
                Columns(r).Delete
                r = r - 1
            End If
            Next
    ActiveWorkbook.SaveAs Filename:=Path & "\files\HAZ_NA_" + a + b + ".xlsx" _
        , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWindow.Close     '¹Ø±ÕÒÑ¾­Éú³ÉµÄÎÄ¼þ



    MsgBox "All files have been saved successfully at " & Path & "\files"
    
    Sheets("Specification ID TBD").Select

End Sub

Sub Macro_clear_data()
'
' Macro_clear_data Macro
'

'
    Dim response As String
    response = MsgBox("È·¶¨É¾³ýËùÓÐ±í¸ñÖÐµÄÊý¾ÝÂð?", vbYesNo, "Notice")

    If response = 6 Then  '6 stands for Yes, 7 stands for No


    Range("A2:D2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    
    Sheets("Sheet1").Select
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    
    Sheets("Non-Haz").Select
    Range("A2:T2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    
    Sheets("by Sales Org").Select
    Range("A2:T2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    
    Sheets("by Plant").Select
    Range("A2:T2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    
    Sheets("Haz CAS").Select
    Range("A2:T2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    
    Sheets("Haz SYN").Select
    Range("A2:T2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    
    End If
    
    Sheets("Specification ID TBD").Select
    
    
End Sub
