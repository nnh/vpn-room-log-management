Attribute VB_Name = "common"
Option Explicit
Public Const cst_holidaySheetName As String = "holiday"
Public Const cst_checkSheetname As String = "check"
Public Type overtime_info
    targetWorksheet As Worksheet
    targetLastRow As Long
    targetStartRow As Long
    holidayWorksheet As Worksheet
    targetYear As String
    targetMonth As String
    monthColumnNumber As Long
    dayColumnNumber As Long
    timeColumnNumber As Long
    categoryColumnNumber As Long
End Type
Public Function getInputDir(str_path As String) As String
    '処理対象月をYYYYMMの形式で指定してください。空白ならば処理日の前月になります。
    Const yyyymm As String = ""
    Dim parentPath As String
    Dim logPath As String
    Dim temp_ymd As Date
    Dim yyyy As String
    Dim mm As String
    Dim targetFolderName As String
    
    If yyyymm = "" Then
        temp_ymd = DateAdd("m", -1, Date)
        yyyy = Year(temp_ymd)
        mm = month(temp_ymd)
        If Len(mm) = 1 Then
            mm = "0" & mm
        End If
        targetFolderName = yyyy & mm
    Else
        targetFolderName = yyyymm
    End If
    
    logPath = str_path & targetFolderName
    getInputDir = logPath

End Function
Public Sub getVpnAndRoomAccessReport()
    Call get_vpn_logs
    Call get_room_logs
    MsgBox "処理が終了しました"
End Sub

Public Sub getHolidayData()
    Const cst_holidayFolderName As String = "public_holiday"
    Const cst_holidayFileName As String = "祝日入力シート（事務局用）.xlsx"
    Const cst_copyFromSheetName As String = "祝日"
    Dim inputPath As String
    Dim i As Integer
    Dim holidayWB As Workbook
    Dim copyFromWS As Worksheet
    Dim copyToWS As Worksheet
    Set copyToWS = addWorksheet(ThisWorkbook, cst_holidaySheetName)
    inputPath = ThisWorkbook.Path
    For i = 0 To 0
        inputPath = Left(inputPath, InStrRev(inputPath, "¥") - 1)
    Next i
    inputPath = inputPath & "¥" & cst_holidayFolderName & "¥" & cst_holidayFileName
    On Error GoTo FINL_L
    Workbooks.Open Filename:=inputPath
    Set holidayWB = ActiveWorkbook
    Set copyFromWS = holidayWB.Worksheets(cst_copyFromSheetName)
    copyFromWS.Cells.Copy (copyToWS.Cells(1, 1))
FINL_L:
    holidayWB.Close savechanges:=False
End Sub

Public Sub outputPDF(outputWSNames As Variant, outputPath As String, deleteStr As String, printOrientation As Variant)
    Dim outputName As String
    Dim outputWS As Worksheet
    Dim outputWSName As Variant
    outputName = Replace(ThisWorkbook.Name, ".xlsm", ".pdf")
    outputName = Replace(outputName, deleteStr, "")
    For Each outputWSName In ThisWorkbook.Worksheets(outputWSNames)
        Set outputWS = ThisWorkbook.Worksheets(outputWSName.Name)
        With outputWS.PageSetup
            .Orientation = printOrientation
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = False
            .CenterHeader = outputWS.Name
        End With
    Next outputWSName
    ThisWorkbook.Worksheets(outputWSNames).Select
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=outputPath & outputName
    ThisWorkbook.Worksheets(cst_checkSheetname).Select
End Sub

Public Function addWorksheet(targetWorkbook As Workbook, sheetname As String, Optional copyfromSheet As Worksheet = Nothing) As Worksheet
    Dim outputSheet As Worksheet
    ' Delete outputsheet
    On Error Resume Next
    Set outputSheet = targetWorkbook.Worksheets(sheetname)
    On Error GoTo 0
    If Not outputSheet Is Nothing Then
        outputSheet.Delete
    End If
    If Not copyfromSheet Is Nothing Then
        copyfromSheet.Copy after:=copyfromSheet
    Else
        ThisWorkbook.Worksheets.Add
    End If
    ActiveSheet.Name = sheetname
    Set addWorksheet = ThisWorkbook.Worksheets(sheetname)
End Function

Public Sub setOvertimeInfo(target As overtime_info)
    Dim i As Long
    Dim tempDate As Date
    Dim tempWeekday As Integer
    Dim tempMatch As Variant
    For i = target.targetLastRow To target.targetStartRow Step -1
        With target.targetWorksheet
            tempDate = CDate(.Cells(i, target.monthColumnNumber).Value & " " & .Cells(i, target.dayColumnNumber).Value & ", " & target.targetYear)
            tempWeekday = Weekday(tempDate)
            tempMatch = Null
            On Error Resume Next
            If (tempWeekday <> vbSunday) And (tempWeekday <> vbSaturday) Then
                tempMatch = WorksheetFunction.Match(CLng(tempDate), target.holidayWorksheet.Range("A:A"), 0)
                If IsEmpty(tempMatch) Or IsNull(tempMatch) Then
                    If (CDate("22:00:00") < CDate(.Cells(i, target.timeColumnNumber).Value)) Or (CDate(.Cells(i, target.timeColumnNumber).Value) < CDate("5:00:00")) Then
                        .Cells(i, target.categoryColumnNumber).Value = "深夜"
                    Else
                        .Rows(i).Delete
                    End If
                Else
                    .Cells(i, target.categoryColumnNumber).Value = "休日"
                End If
            Else
                .Cells(i, target.categoryColumnNumber).Value = "休日"
            End If
            On Error GoTo 0
        End With
    Next i
End Sub
