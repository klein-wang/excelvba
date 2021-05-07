Sub Macro_try()
'
' Macro_try Macro
'

'
    Dim user As String
        user = Environ("username") 'change username
    Dim a, b As String
        a = Format(Date, "yyyymmdd") 'date
        b = Format(Date, "yymmdd") 'shorter date
    Dim name As Variant
        name = Array("物料库", "社会化预归类库", "企业内部归类库", "待归类物料", "排除清单")
    
    Path = "C:\Users\" & user & "\Downloads\"
    mainfile = "Material List_Master Data Check_" & b & ".xlsm" '最终导入的文件名
    
    Dim filename1, filename2, filename3, filename4, filename5 As Variant
        filename1 = Dir(Path & name(0) & a & "*.xls")
        filename2 = Dir(Path & name(1) & a & "*.xls")
        filename3 = Dir(Path & name(2) & a & "*.xls")
        filename4 = Dir(Path & name(3) & a & "*.xls")
        filename5 = Dir(Path & name(4) & a & "*.xls")
    
    
    
    ChDir Path 'open default filepath
    
    '物料库
    Workbooks.Open filename:=filename1
    Columns("A:V").Select
    Selection.Copy
    Windows(mainfile).Activate
    Sheets("物理库").Select
    Columns("B:W").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False 'Èç¹ûÊý¾ÝÏÔÊ¾Òì³££¬¿ÉÄÜÊÇColumnBÃ»ÓÐÏÂÀ­¡£
    Workbooks(filename1).Activate
    ActiveWindow.Close
    
     '社会化预归类
    Workbooks.Open filename:=filename2
    Columns("A:S").Select
    Selection.Copy
    Windows(mainfile).Activate
    Sheets("社会化预归类").Select
    Columns("C:U").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False 'Èç¹ûÊý¾ÝÏÔÊ¾Òì³££¬¿ÉÄÜÊÇColumnA,BÃ»ÓÐÏÂÀ­¡£
    Windows(filename2).Activate
    ActiveWindow.Close
    
     '企业内部归类
    Workbooks.Open filename:=filename3
    Columns("A:N").Select
    Selection.Copy
    Windows(mainfile).Activate
    Sheets("企业内部预归类").Select
    Columns("B:O").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False 'Èç¹ûÊý¾ÝÏÔÊ¾Òì³££¬¿ÉÄÜÊÇColumnBÃ»ÓÐÏÂÀ­¡£
    Windows(filename3).Activate
    ActiveWindow.Close
    
     '待归类
    Workbooks.Open filename:=filename4
    Columns("A:Q").Select
    Selection.Copy
    Windows(mainfile).Activate
    Sheets("待归类物料").Select
    Columns("B:R").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False 'Èç¹ûÊý¾ÝÏÔÊ¾Òì³££¬¿ÉÄÜÊÇColumnBÃ»ÓÐÏÂÀ­¡£
    Windows(filename4).Activate
    ActiveWindow.Close
    
     '排除清单
    Workbooks.Open filename:=filename5
    Columns("A:S").Select
    Selection.Copy
    Windows(mainfile).Activate
    Sheets("排除清单").Select
    Columns("B:T").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False 'Èç¹ûÊý¾ÝÏÔÊ¾Òì³££¬¿ÉÄÜÊÇColumnBÃ»ÓÐÏÂÀ­¡£
    Windows(filename5).Activate
    ActiveWindow.Close
    
    Sheets("Dashboard").Select
    
End Sub
