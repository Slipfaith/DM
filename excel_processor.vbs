' VBScript for processing Excel files
' Usage: cscript //NoLogo excel_processor.vbs <file_path> <header_color>

If WScript.Arguments.Count < 2 Then
    WScript.Echo "Usage: cscript //NoLogo excel_processor.vbs <file> <header_color>"
    WScript.Quit 1
End If

Dim filePath, headerColor
filePath = WScript.Arguments(0)
headerColor = CLng(WScript.Arguments(1))

Dim excel, wb, sheet
Set excel = CreateObject("Excel.Application")
excel.Visible = False
excel.ScreenUpdating = False
excel.DisplayAlerts = False
excel.EnableEvents = False
On Error Resume Next
excel.Calculation = -4135 ' xlCalculationManual
On Error GoTo 0

Set wb = excel.Workbooks.Open(filePath)

For Each sheet In wb.Sheets
    ProcessSheet sheet, headerColor
Next

wb.Save
wb.Close False
excel.Quit

Sub ProcessSheet(sheet, headerColor)
    Dim headerRange
    Set headerRange = FindHeader(sheet, headerColor)
    If Not headerRange Is Nothing Then
        RestructureSheet sheet, headerRange
    End If
End Sub

Function FindHeader(sheet, headerColor)
    Dim usedRange, rowsCount, colsCount, row, col
    Set usedRange = sheet.UsedRange
    rowsCount = usedRange.Rows.Count
    colsCount = usedRange.Columns.Count
    For row = 1 To rowsCount
        If row > 20 Then Exit For
        For col = 1 To colsCount
            If sheet.Cells(row, col).Interior.Color = headerColor Then
                Set FindHeader = FindHeaderRange(sheet, row)
                Exit Function
            End If
        Next
    Next
    Set FindHeader = Nothing
End Function

Function FindHeaderRange(sheet, headerRow)
    Dim usedRange, colsCount, col, firstCol, lastCol, value
    Set usedRange = sheet.UsedRange
    colsCount = usedRange.Columns.Count
    firstCol = 0
    lastCol = 0
    For col = 1 To colsCount
        value = sheet.Cells(headerRow, col).Value
        If Not IsEmpty(value) Then
            If firstCol = 0 Then firstCol = col
            lastCol = col
        End If
    Next
    If firstCol > 0 And lastCol > 0 Then
        Set FindHeaderRange = sheet.Range(sheet.Cells(headerRow, firstCol), sheet.Cells(headerRow, lastCol))
    Else
        Set FindHeaderRange = Nothing
    End If
End Function

Function HasDataInRange(sheet, row, startCol, endCol)
    Dim col, value
    HasDataInRange = False
    For col = startCol To endCol
        value = sheet.Cells(row, col).Value
        If Not IsEmpty(value) Then
            HasDataInRange = True
            Exit Function
        End If
    Next
End Function

Sub RestructureSheet(sheet, headerRange)
    Dim headerRow, startCol, endCol, usedRange, lastRow, headerHeight
    Dim row, insertRow
    headerRow = headerRange.Row
    startCol = headerRange.Column
    endCol = startCol + headerRange.Columns.Count - 1
    Set usedRange = sheet.UsedRange
    lastRow = usedRange.Row + usedRange.Rows.Count - 1
    headerHeight = sheet.Rows(headerRow).RowHeight
    row = headerRow + 1
    Do While row <= lastRow
        If HasDataInRange(sheet, row, startCol, endCol) Then
            sheet.Rows(row).Copy
            insertRow = row + 1
            sheet.Rows(insertRow).Insert -4121
            sheet.Rows(insertRow).PasteSpecial -4104
            row = insertRow + 1

            sheet.Rows(row).Insert -4121
            sheet.Rows(row).Clear
            sheet.Rows(row).RowHeight = 15
            lastRow = lastRow + 2
            row = row + 1
        Else
            row = row + 1
        End If
    Loop
    sheet.Application.CutCopyMode = False
    sheet.Rows(headerRow).RowHeight = headerHeight
End Sub
