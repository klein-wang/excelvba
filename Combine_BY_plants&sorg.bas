Sub SAP()
'
' SAP Macro
'

'
    Path = ActiveWorkbook.Path & "\files"
    Dim count As Integer
    Dim a As String
        a = Format(Date, "yyyymmdd") 'date
    Dim name As Variant
        name = Array("Basic Info_Plant_", "Basic Info_Sorg_") 'Name of two files
    Dim filename1, filename2 As Variant
        filename1 = Dir(Path & "\" & name(0) & a & "*.xlsx")
        filename2 = Dir(Path & "\" & name(1) & a & "*.xlsx")
    
    ChDir Path
    Workbooks.Open Filename:=filename1  'Go to By Plant file
    count = Sheets("by Plant").Range("A1").End(xlDown).Row
    Sheets("by Plant").Select
    Sheets("by Plant").Copy After:=Workbooks("Basic Info.xlsm").Sheets(1)
    Workbooks.Open Filename:=filename1
    ActiveWindow.Close
    
    Sheets("by Plant").Select
    Range("J2").Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(RC[-9],RC[-2])"
    Range("J2").Select
    Selection.AutoFill Destination:=Range("J2:J" & count)
    
    Columns("A:I").Select
    Selection.Copy
    Sheets("Basic Info").Select
    Range("A1").Select
    ActiveSheet.Paste
    
    Workbooks.Open Filename:=filename2 'Go to By Sales Org file
    
    Sheets("by Sales Org").Select
    Sheets("by Sales Org").Copy After:=Workbooks("Basic Info.xlsm").Sheets(2)
    Workbooks.Open Filename:=filename2
    ActiveWindow.Close
    
    Sheets("by Sales Org").Select
    Range("A2:I2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Basic Info").Select
    Range("A" & count + 1).Select
    ActiveSheet.Paste
    
    Columns("G:H").Select
    Selection.ClearContents
    
    Columns("A:A").Select 'remove duplicates
    ActiveSheet.Range("$A$1:$I$10160").RemoveDuplicates Columns:=Array(1, 2, 3, 4, 5, _
        6, 7, 8, 9), Header:=xlYes

    'rename columns
    [J1].Value = "plant"
    [K1].Value = "sales org"
    [L1].Value = "valid"
    
    'Import formula
    Range("J2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(AND(ISERROR(VLOOKUP(RC[-9],'by Plant'!C1,1,0))=FALSE,ISERROR(VLOOKUP(RC[-9],'by Plant'!C10,1,0)=TRUE)),""X"","""")"
    Range("K2").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-10],'by Sales Org'!C1:C8,8,0),"""")"
    Range("L2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC[-6]=""X"",AND(RC[-2]=""X"",RC[-1]=""X"")),""X"","""")"
    Range("J2:L2").Select
    Selection.AutoFill Destination:=Range("J2:L10160")
    
    Range("J1").Select
    Selection.Copy
    Range("J1:L1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    
    ActiveWorkbook.Save
    
End Sub
_________________________________________________________________________________
Sub savefile()
'
' savefile Macro
'

'
    Path = ActiveWorkbook.Path
    Dim a As String
        a = Format(Date, "yyyymmdd") 'date
    Dim Org As String
    Org = Sheets("by Sales Org").Range("G2")
    
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet3").Select
    Sheets("Sheet3").name = Org 'choose the name of legal entity
    
    Sheets("Basic Info").Select
    Columns("A:A").Select 'Material No
    Selection.Copy
    Sheets(Org).Select
    Columns("A:A").Select
    ActiveSheet.Paste

    Dim count As Integer
    count = Sheets("Basic Info").Range("A1").End(xlDown).Row
    
    Sheets(Org).Select
    Range("B2").Select
    ActiveCell.FormulaR1C1 = _
        "=CONCATENATE(TEXT(RC[-1]*1,""#""),'by Sales Org'!RC[5])" '&Sorg
    Range("B2").Select
    Selection.AutoFill Destination:=Range("B2:B" & count)
 
    Sheets("Basic Info").Select
    Columns("I:I").Select 'SPEC ID
    Selection.Copy
    Sheets(Org).Select
    Columns("C:C").Select
    ActiveSheet.Paste
    
    Sheets("Basic Info").Select 'Material Description, MTyp, Dv, UVP
    Columns("B:E").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets(Org).Select
    Range("D1").Select
    ActiveSheet.Paste
    
    Sheets("Basic Info").Select 'del flag
    Columns("L:L").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets(Org).Select
    Columns("H:H").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets(Org).Select
    
    'rename the columns
    [B1].Value = "&Org"
    [C1].Value = "SPEC ID"
    [H1].Value = "Del.flag"
    
    Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("B1:H1").Select 'paste the formatting
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveWorkbook.Save

    
    Sheets(Org).Select 'save as new file
    Sheets(Org).Copy
    ActiveWorkbook.SaveAs Filename:= _
        Path & "\" & Org & "\" & Org & "_" & a & ".xlsx", FileFormat:= _
        xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWindow.Close
    
    
End Sub
______________________________________________________________________________
Sub clean_data()
'
' clean_data Macro
'

'
    Dim Org As String
    Org = Sheets("by Sales Org").Range("G2")
        
    'delete three sheets
    Sheets("by Plant").Select
    Application.CutCopyMode = False
    ActiveWindow.SelectedSheets.Delete
    Sheets("by Sales Org").Select
    ActiveWindow.SelectedSheets.Delete
    Sheets(Org).Select
    ActiveWindow.SelectedSheets.Delete
    
    'clear the content of 'Basic Info'
    Sheets("Basic Info").Select
    Columns("A:L").Select
    Selection.ClearContents
    Range("A3").Select
    Selection.Copy
    Range("A1:I1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
   
    ActiveWorkbook.Save
    
   
End Sub


