Attribute VB_Name = "vpn_management"
Option Explicit
Type pivottable_info
    cst_pivottable_name As String
    input_ws As Worksheet
    output_ws As Worksheet
    range_area As String
End Type
Const cst_outputSheetname As String = "overtime"
Const cst_vpndstSheetname As String = "data"

Public Sub get_vpn_logs()
'
Application.ScreenUpdating = False
Application.DisplayAlerts = False
'
    Dim inputSheet As Worksheet
    Set inputSheet = ThisWorkbook.Worksheets("vpn_input")
    Dim dstSheet    As Worksheet
    Set dstSheet = ThisWorkbook.Worksheets(cst_vpndstSheetname)
    dstSheet.Cells.Clear

    Dim LogFolder   As String
    Dim buf         As String
    Dim a           As String
    Dim lastmonth   As String
    Dim month       As String
    Dim nextmonth   As String
    Dim LastLog     As String
    Dim M           As String
    Dim temp_ws     As Worksheet
        
    LogFolder = inputSheet.Range("H1")

    a = inputSheet.Range("I1") & "/" & inputSheet.Range("J1")
    lastmonth = DateAdd("m", -1, a)
    lastmonth = Format(lastmonth, "yyyymm")
    lastmonth = "¥access.log-" & lastmonth & "*"
    
    buf = Dir(LogFolder & lastmonth)
    Call getVPNInfo(buf, LogFolder, dstSheet)
    
    month = Format(a, "yyyymm")
    month = "¥access.log-" & month & "*"
    
    buf = Dir(LogFolder & month)
    Call getVPNInfo(buf, LogFolder, dstSheet)
    
    nextmonth = DateAdd("m", 1, a)
    nextmonth = Format(nextmonth, "yyyymm")
    nextmonth = "¥access.log-" & nextmonth & "*"
    
    buf = Dir(LogFolder & nextmonth)
    Call getVPNInfo(buf, LogFolder, dstSheet)
    
    LastLog = "¥access.log"
    
    buf = Dir(LogFolder & LastLog)
    Call getVPNInfo(buf, LogFolder, dstSheet)
    
    dstSheet.Columns("A:A").TextToColumns Destination:=Range("A1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(3, 1), Array(6, 1), Array(15, 1)), _
        TrailingMinusNumbers:=True
    dstSheet.Columns("A:D").EntireColumn.AutoFit
    dstSheet.Range("A1").Select
    
    M = inputSheet.Range("K1").Value
    dstSheet.Range("A:A").AutoFilter Field:=1, Criteria1:= _
    "<>*" & M & "*", Operator:=xlAnd
    dstSheet.Cells.Delete Shift:=xlUp
    dstSheet.Range("A1").Select

    For Each temp_ws In ThisWorkbook.Worksheets
        If temp_ws.Name = cst_checkSheetname Then
            temp_ws.Delete
            Exit For
        End If
    Next temp_ws
    dstSheet.Copy before:=ThisWorkbook.Worksheets(1)
    ActiveSheet.Name = cst_checkSheetname
    With ThisWorkbook.Worksheets(cst_checkSheetname)
        .Range("D:D").AutoFilter Field:=1, Criteria1:="<>*Call*", Operator:=xlAnd
        .Cells.Delete Shift:=xlUp
        .Range("E1").FormulaR1C1 = "=MID(RC[-1],FIND(""'"",RC[-1])+1,LEN(RC[-1])-FIND(""'"",RC[-1])-1)"
        .Range("E1").AutoFill Destination:=.Range("E1:E" & .Range("D" & .Rows.Count).End(xlUp).Row)
        .Range("E:E").Copy
        .Range("E:E").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        .Cells.Select

        Application.CutCopyMode = False
        With .Sort
            .SortFields.Clear
            .SortFields.Add Key:=Range("E:E"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SortFields.Add Key:=Range("B:B"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SortFields.Add Key:=Range("C:C"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SetRange Range("A:E")
            .Header = xlGuess
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With

        .Range("A:E").EntireColumn.AutoFit
        .Range("A1").Select
    End With
    
    ThisWorkbook.Save
    Call getHolidayData
    Call checkOvertime
    Call checkConnectedIPaddress
    ThisWorkbook.Worksheets(cst_outputSheetname).Move before:=ThisWorkbook.Worksheets(cst_checkSheetname)
    Call outputPDF(Array(cst_outputSheetname, cst_checkSheetname), "¥¥ARONAS¥Archives¥Log¥VPN¥")
    
Application.ScreenUpdating = True
Application.DisplayAlerts = True
ActiveWorkbook.Save
    
End Sub

Public Sub checkOvertime()
    Const cst_hmsCol As Integer = 3
    Dim categoryCol As Integer
    Dim holidaySheet As Worksheet
    Dim checkSheet As Worksheet
    Dim outputSheet As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim targetYear As String
    Dim targetMonth As String
    Dim tempDate As Date
    Dim tempWeekday As Integer
    Dim tempMatch As Variant
    
    Set holidaySheet = ThisWorkbook.Worksheets(cst_holidaySheetName)
    Set checkSheet = ThisWorkbook.Worksheets(cst_checkSheetname)
    Set outputSheet = addWorksheet(ThisWorkbook, cst_outputSheetname, checkSheet)
    
    With checkSheet.Cells.SpecialCells(xlCellTypeLastCell)
        lastRow = .Row
        categoryCol = .Column + 1
    End With
    
    targetMonth = outputSheet.Cells(1, 1).Value
    targetYear = getTargetYear(targetMonth)
    
    For i = lastRow To 1 Step -1
        With outputSheet
            tempDate = CDate(.Cells(i, 1).Value & " " & .Cells(i, 2).Value & ", " & targetYear)
            tempWeekday = Weekday(tempDate)
            tempMatch = Null
            On Error Resume Next
            If (tempWeekday <> vbSunday) And (tempWeekday <> vbSaturday) Then
                tempMatch = WorksheetFunction.Match(CLng(tempDate), holidaySheet.Range("A:A"), 0)
                If IsEmpty(tempMatch) Or IsNull(tempMatch) Then
                    If (CDate("22:00:00") < CDate(.Cells(i, cst_hmsCol).Value)) Or (CDate(.Cells(i, cst_hmsCol).Value) < CDate("5:00:00")) Then
                        .Cells(i, categoryCol).Value = "深夜"
                    Else
                        .Rows(i).Delete
                    End If
                Else
                    .Cells(i, categoryCol).Value = "休日"
                End If
            Else
                .Cells(i, categoryCol).Value = "休日"
            End If
            On Error GoTo 0
        End With
    Next i
    ' insert header
    outputSheet.Rows(1).Insert
    With outputSheet.Cells(1, 1)
        .Value = "月"
        .Offset(0, 1).Value = "日"
        .Offset(0, 2).Value = "時間"
        .Offset(0, 3).Value = "メッセージ"
        .Offset(0, 4).Value = "ユーザー"
        .Offset(0, 5).Value = "区分"
    End With
    outputSheet.Cells.EntireColumn.AutoFit
End Sub

Private Function getTargetYear(targetMonth As String) As Integer
    Dim targetYear As String
    
    If targetMonth = "Dec" Then
        targetYear = Year(Now) - 1
    Else
        targetYear = Year(Now)
    End If
    
    getTargetYear = targetYear
    
End Function

Public Sub checkConnectedIPaddress()
    Const output_sheetname As String = "connected_from"
    Dim dataSheet As Worksheet
    Dim outputSheet As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim tempstr As String
    Dim tempstr_2 As Variant
    Dim tempstr_3 As Variant
    Dim str_ip As String
    Dim str_user As String
    Dim output_row As Long
        
    Set dataSheet = ThisWorkbook.Worksheets("data")
    Set outputSheet = addWorksheet(ThisWorkbook, "connected_from")
    lastRow = dataSheet.Cells.SpecialCells(xlCellTypeLastCell).Row
    output_row = 1
    outputSheet.Cells(output_row, 1).Value = "月"
    outputSheet.Cells(output_row, 2).Value = "日"
    outputSheet.Cells(output_row, 3).Value = "時間"
    outputSheet.Cells(output_row, 4).Value = "接続元IPアドレス"
    outputSheet.Cells(output_row, 5).Value = "ユーザー"
    For i = 1 To lastRow
        tempstr = dataSheet.Cells(i, 4).Value
        If InStr(tempstr, "connected from") > 0 Then
            output_row = output_row + 1
            tempstr_2 = Split(tempstr, " ")
            outputSheet.Cells(output_row, 1).Value = dataSheet.Cells(i, 1).Value
            outputSheet.Cells(output_row, 2).Value = dataSheet.Cells(i, 2).Value
            outputSheet.Cells(output_row, 3).Value = dataSheet.Cells(i, 3).Value
            outputSheet.Cells(output_row, 4).Value = tempstr_2(UBound(tempstr_2))
        ElseIf InStr(tempstr, "Call detected from user") Then
            tempstr_3 = Split(tempstr, " ")
            outputSheet.Cells(output_row, 5).Value = tempstr_3(UBound(tempstr_3))
        End If
    Next i
    outputSheet.Columns(3).NumberFormatLocal = "[$-x-systime]h:mm:ss AM/PM"
    outputSheet.Cells.EntireColumn.AutoFit
    Call createPivottable(ThisWorkbook, output_sheetname, "summary_connected_from")

End Sub

Private Sub createPivottable(output_wb As Workbook, input_ws_name As String, output_ws_name As String)
    Dim pv_info As pivottable_info
    pv_info = setPivottableInfo(output_wb, input_ws_name, output_ws_name, "pv1")
    output_wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=pv_info.input_ws.Range(pv_info.range_area)).createPivottable _
                                 TableDestination:=pv_info.output_ws.Range("A1"), TableName:=pv_info.cst_pivottable_name
    With pv_info.output_ws.PivotTables(pv_info.cst_pivottable_name)
        .InGridDropZones = True
        .RowAxisLayout xlTabularRow
        .AddDataField pv_info.output_ws.PivotTables(pv_info.cst_pivottable_name).PivotFields("ユーザー")
        .PivotFields("ユーザー").Orientation = xlRowField
        .AddDataField pv_info.output_ws.PivotTables(pv_info.cst_pivottable_name).PivotFields("接続元IPアドレス")
        .PivotFields("接続元IPアドレス").Orientation = xlRowField
        .PivotFields("ユーザー").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
        On Error Resume Next
        .PivotFields("個数 / ユーザー").Orientation = xlHidden
        .PivotFields("個数 / 接続元IPアドレス").Orientation = xlHidden
        On Error GoTo 0
    End With
    output_wb.Worksheets(output_ws_name).Cells.EntireColumn.AutoFit
End Sub

Private Function setPivottableInfo(output_wb As Workbook, input_ws_name As String, output_ws_name As String, pv_name As String) As pivottable_info
    Dim pv_info As pivottable_info
    Dim outputSheet As Worksheet
    Set outputSheet = addWorksheet(output_wb, output_ws_name)
    outputSheet.Activate
    With pv_info
        .cst_pivottable_name = pv_name
        Set .input_ws = output_wb.Worksheets(input_ws_name)
        Set .output_ws = output_wb.Worksheets(output_ws_name)
        .range_area = "A:E"
    End With
    setPivottableInfo = pv_info
End Function
Private Sub getVPNInfo(buf As String, LogFolder As String, dstSheet As Worksheet)
    Dim i As Long
    Dim srcBook     As Workbook
    Dim srcSheet    As Worksheet
    Dim LstRow1     As Long
    Dim LstRow2     As Long
    i = 0
    Do While buf <> ""
        i = i + 1
        Set srcBook = Workbooks.Open(LogFolder + "¥" + buf)
        Set srcSheet = srcBook.Worksheets(1)
        With srcSheet
            .Select
            LstRow1 = .Cells(.Rows.Count, 1).End(xlUp).Row
            .Range("A1:A" & LstRow1).Copy
        End With
        With dstSheet
            LstRow2 = .Cells(.Rows.Count, 1).End(xlUp).Row
            .Range("A" & LstRow2).Offset(1, 0).PasteSpecial xlPasteAll
        End With
        srcBook.Close False

        buf = Dir()
    Loop

End Sub





