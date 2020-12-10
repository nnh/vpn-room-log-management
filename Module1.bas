Attribute VB_Name = "Module1"
Sub list()
'
Application.ScreenUpdating = False
Application.DisplayAlerts = False
'
    Dim dstSheet    As Worksheet
    Set dstSheet = ThisWorkbook.Worksheets(1)

    Dim srcBook     As Workbook
    Dim srcSheet    As Worksheet

    Dim LogFolder   As String
    Dim buf         As String
    Dim a           As String
    Dim lastmonth   As String
    Dim month       As String
    Dim nextmonth   As String
    Dim LastLog     As String
    Dim M           As String
    
    Dim i           As Long
    
    LogFolder = dstSheet.Range("H1")

    a = dstSheet.Range("I1") & "/" & dstSheet.Range("J1")
    lastmonth = DateAdd("m", -1, a)
    lastmonth = Format(lastmonth, "yyyymm")
    lastmonth = "\access.log-" & lastmonth & "*"
    
    buf = Dir(LogFolder & lastmonth)

    i = 0
    Do While buf <> ""
        i = i + 1
        Set srcBook = Workbooks.Open(LogFolder + "\" + buf)
        Set srcSheet = srcBook.Worksheets(1)
        srcSheet.Select
        LstRow1 = srcSheet.Cells(Rows.Count, 1).End(xlUp).Row
        srcSheet.Range("A1:A" & LstRow1).Copy
        
        LstRow2 = dstSheet.Cells(Rows.Count, 1).End(xlUp).Row
        dstSheet.Range("A" & LstRow2).Offset(1, 0).PasteSpecial xlPasteAll

        srcBook.Close False

        buf = Dir()
    Loop
    
    month = Format(a, "yyyymm")
    month = "\access.log-" & month & "*"
    
    buf = Dir(LogFolder & month)

    i = 0
    Do While buf <> ""
        i = i + 1
        Set srcBook = Workbooks.Open(LogFolder + "\" + buf)
        Set srcSheet = srcBook.Worksheets(1)
        srcSheet.Select
        LstRow1 = srcSheet.Cells(Rows.Count, 1).End(xlUp).Row
        srcSheet.Range("A1:A" & LstRow1).Copy
        
        LstRow2 = dstSheet.Cells(Rows.Count, 1).End(xlUp).Row
        dstSheet.Range("A" & LstRow2).Offset(1, 0).PasteSpecial xlPasteAll

        srcBook.Close False

        buf = Dir()
    Loop
    
    nextmonth = DateAdd("m", 1, a)
    nextmonth = Format(nextmonth, "yyyymm")
    nextmonth = "\access.log-" & nextmonth & "*"
    
    buf = Dir(LogFolder & nextmonth)

    i = 0
    Do While buf <> ""
        i = i + 1
        Set srcBook = Workbooks.Open(LogFolder + "\" + buf)
        Set srcSheet = srcBook.Worksheets(1)
        srcSheet.Select
        LstRow1 = srcSheet.Cells(Rows.Count, 1).End(xlUp).Row
        srcSheet.Range("A1:A" & LstRow1).Copy
        
        LstRow2 = dstSheet.Cells(Rows.Count, 1).End(xlUp).Row
        dstSheet.Range("A" & LstRow2).Offset(1, 0).PasteSpecial xlPasteAll

        srcBook.Close False

        buf = Dir()
    Loop
    
    LastLog = "\access.log"
    
    buf = Dir(LogFolder & LastLog)

    Set srcBook = Workbooks.Open(LogFolder + "\" + buf)
    Set srcSheet = srcBook.Worksheets(1)
    srcSheet.Select
    LstRow1 = srcSheet.Cells(Rows.Count, 1).End(xlUp).Row
    srcSheet.Range("A1:A" & LstRow1).Copy
    
    LstRow2 = dstSheet.Cells(Rows.Count, 1).End(xlUp).Row
    dstSheet.Range("A" & LstRow2).Offset(1, 0).PasteSpecial xlPasteAll

    srcBook.Close False

    buf = Dir()

    Sheets("data").Select
    Columns("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(3, 1), Array(6, 1), Array(15, 1)), _
        TrailingMinusNumbers:=True
    Cells.Select

    Columns("A:D").Select
    Columns("A:D").EntireColumn.AutoFit
    Range("A1").Select
    
    M = Range("K1").Value
    ActiveSheet.Range("A:A").AutoFilter Field:=1, Criteria1:= _
    "<>*" & M & "*", Operator:=xlAnd
    Range(Selection, Selection.End(xlDown)).Select
    Cells.Select
    Selection.Delete Shift:=xlUp
    Range("A1").Select

    Sheets("data").Select
    Sheets("data").Copy Before:=Worksheets(1)
    ActiveSheet.Name = "check"
    

    Sheets("check").Select
    ActiveSheet.Range("D:D").AutoFilter Field:=1, Criteria1:= _
    "<>*Call*", Operator:=xlAnd
    Range(Selection, Selection.End(xlDown)).Select
    Cells.Select
    Selection.Delete Shift:=xlUp
    Range("A1").Select

    Range("E1").Select
    ActiveCell.FormulaR1C1 = _
        "=MID(RC[-1],FIND(""'"",RC[-1])+1,LEN(RC[-1])-FIND(""'"",RC[-1])-1)"

    Range("E1").Copy
    Selection.AutoFill Destination:=Range("E1:E" & Range("D" & Rows.Count).End(xlUp).Row)
    Range("E1:E" & Range("D" & Rows.Count).End(xlUp).Row).Select
    
    Range("E:E").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Cells.Select

    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("check").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("check").Sort.SortFields.Add Key:=Range("E:E") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("check").Sort.SortFields.Add Key:=Range("B:B") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("check").Sort.SortFields.Add Key:=Range("C:C") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("check").Sort
        .SetRange Range("A:E")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Columns("A:E").Select
    Columns("A:E").EntireColumn.AutoFit
    Range("A1").Select
    
'    ActiveWorkbook.Save
    
Application.ScreenUpdating = True
Application.DisplayAlerts = True
    
End Sub
