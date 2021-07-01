Attribute VB_Name = "common"
Option Explicit
Public Const cst_holidaySheetName As String = "holiday"
Public Const cst_checkSheetname As String = "check"

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

Public Sub outputPDF(outputWSNames As Variant, outputPath As String)
    Dim outputName As String
    Dim outputWS As Worksheet
    Dim outputWSName As Variant
    outputName = Replace(ThisWorkbook.Name, ".xlsm", ".pdf")
    For Each outputWSName In ThisWorkbook.Worksheets(outputWSNames)
        Set outputWS = ThisWorkbook.Worksheets(outputWSName.Name)
        With outputWS.PageSetup
            .Orientation = xlPortrait
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

