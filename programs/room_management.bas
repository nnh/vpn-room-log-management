Attribute VB_Name = "room_management"
Option Explicit

Public Sub get_room_logs()

Application.ScreenUpdating = False
Application.DisplayAlerts = False
    Const cst_inputSheetName As String = "room_input"
    Const cst_outputSheetName As String = "room_list"
    Call getHolidayData
    
    Dim LstRow  As Long
    Dim LstRow1 As Long
    Dim LstRow2 As Long
    
    Dim inputSheet As Worksheet
    Set inputSheet = ThisWorkbook.Worksheets(cst_inputSheetName)
    inputSheet.Cells.Clear
    Dim outputSheet As Worksheet

    Dim srcBook As Workbook
    Dim srcSheet As Worksheet

    Dim buf As String
    Dim inputPath As String
    inputPath = getInputDir("¥¥aronas¥Archives¥Log¥DC入退室¥rawdata¥")
    buf = Dir(inputPath & "¥*.csv")

    Dim i As Long
    Dim j As Long
    Dim overtime As overtime_info
    i = 0
    j = 0
    Do While buf <> ""
        i = i + 1
        Set srcBook = Workbooks.Open(inputPath + "¥" + buf)
        Set srcSheet = srcBook.Worksheets(1)
        
        LstRow1 = srcSheet.Cells(srcSheet.Rows.Count, 1).End(xlUp).Row
        If i = 1 Then
            srcSheet.Range("A4:H" & LstRow1).Copy
        Else
            srcSheet.Range("A5:H" & LstRow1).Copy
        End If
        j = j + LstRow1 - 4
        
        LstRow2 = inputSheet.Cells(inputSheet.Rows.Count, 1).End(xlUp).Row
        inputSheet.Range("A" & LstRow2).Offset(1, 0).PasteSpecial xlPasteAll

        srcBook.Close False

        buf = Dir()
    Loop
    inputSheet.Rows(1).Delete
    inputSheet.Cells.EntireColumn.AutoFit
    Set outputSheet = addWorksheet(ThisWorkbook, cst_outputSheetName, inputSheet)
    ThisWorkbook.Save
    Set overtime.targetWorksheet = outputSheet
    overtime.targetStartRow = 2
    Set overtime.holidayWorksheet = ThisWorkbook.Worksheets(cst_holidaySheetName)
    With outputSheet.Cells.SpecialCells(xlCellTypeLastCell)
        overtime.targetLastRow = .Row
        overtime.categoryColumnNumber = .Column + 1
        overtime.monthColumnNumber = .Column + 2
        overtime.dayColumnNumber = .Column + 3
        overtime.timeColumnNumber = .Column + 4
    End With
    For i = overtime.targetStartRow To overtime.targetLastRow
        With outputSheet
            .Cells(i, overtime.monthColumnNumber) = month(.Cells(i, 1).Value)
            .Cells(i, overtime.dayColumnNumber) = Day(.Cells(i, 1).Value)
            .Cells(i, overtime.timeColumnNumber) = TimeValue(.Cells(i, 1).Value)
        End With
    Next i
    overtime.targetYear = Left(Right(inputPath, 6), 4)
    overtime.targetMonth = Mid(Right(inputPath, 6), 5, 2)
    Call setOvertimeInfo(overtime)
    ThisWorkbook.Save
    overtime.targetLastRow = outputSheet.Cells.SpecialCells(xlCellTypeLastCell).Row
    For i = overtime.targetLastRow To overtime.targetStartRow Step -1
        If Trim(outputSheet.Cells(i, 6).Value = "") Then
            outputSheet.Rows(i).Delete
        End If
    Next i
    For i = overtime.timeColumnNumber To overtime.monthColumnNumber Step -1
        outputSheet.Columns(i).Delete
    Next i
    Call outputPDF(Array(cst_outputSheetName), "¥¥aronas¥Archives¥Log¥DC入退室¥", "VPN ", xlLandscape)

Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub

