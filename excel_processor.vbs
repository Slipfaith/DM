Option Explicit

Const xlShiftDown = -4121
Const xlPasteAll = -4104
Const xlCalculationManual = -4135
Const xlCalculationAutomatic = -4105

Dim args
Set args = WScript.Arguments
If args.Count < 2 Then
    WScript.Echo "Usage: cscript excel_processor.vbs <file> <headerColor>"
    WScript.Quit 1
End If

Dim filepath, headerColor
filepath = args(0)
headerColor = CLng(args(1))

Dim excel, wb, sheet
Set excel = CreateObject("Excel.Application")
excel.Visible = False

On Error Resume Next
Set wb = excel.Workbooks.Open(filepath)
If Err.Number <> 0 Then
    WScript.Echo "Error opening workbook: " & Err.Description
    excel.Quit
    WScript.Quit 1
End If
On Error GoTo 0

excel.ScreenUpdating = False
excel.Calculation = xlCalculationManual

For Each sheet In wb.Sheets
    ProcessSheet sheet, headerColor
Next

excel.Calculation = xlCalculationAutomatic
excel.ScreenUpdating = True
wb.Save
wb.Close False
excel.Quit

Sub ProcessSheet(sheet, headerColor)
    Dim headerRange
    Set headerRange = FindHeader(sheet, headerColor)
    If headerRange Is Nothing Then Exit Sub

    Dim headerRow, startCol, endCol
    headerRow = headerRange.Row
    startCol = headerRange.Column
    endCol = startCol + headerRange.Columns.Count - 1

    Dim usedRange, lastRow
    Set usedRange = sheet.UsedRange
    lastRow = usedRange.Row + usedRange.Rows.Count - 1

    Dim blocks, row
    Set blocks = CreateObject("System.Collections.ArrayList")
    row = headerRow + 1
    Do While row <= lastRow
        If sheet.Rows(row).Hidden Then
            row = row + 1
        ElseIf HasDataInRange(sheet, row, startCol, endCol) Then
            Dim blkStart, blkEnd, nextRow
            blkStart = row
            blkEnd = row
            nextRow = row + 1
            If nextRow <= lastRow Then
                If (Not sheet.Rows(nextRow).Hidden) And HasFormulas(sheet, nextRow, startCol, endCol) Then
                    blkEnd = nextRow
                    row = row + 2
                Else
                    row = row + 1
                End If
            Else
                row = row + 1
            End If
            blocks.Add Array(blkStart, blkEnd)
        Else
            row = row + 1
        End If
    Loop

    Dim i
    For i = blocks.Count - 1 To 0 Step -1
        Dim block, sRow, eRow, insertRow, r, c
        block = blocks(i)
        sRow = block(0)
        eRow = block(1)

        sheet.Range(CStr(sRow) & ":" & CStr(eRow)).Copy
        insertRow = eRow + 1
        sheet.Rows(insertRow).Insert xlShiftDown

        For r = 0 To eRow - sRow
            For c = startCol To endCol
                Dim sourceCell, dupCell, formula, colLetter, aboveRow
                Set sourceCell = sheet.Cells(sRow + r, c)
                Set dupCell = sheet.Cells(insertRow + r, c)
                If sourceCell.HasFormula Then
                    formula = sourceCell.Formula
                    If InStr(1, UCase(formula), "LEN(") > 0 Or InStr(1, UCase(formula), "ДЛСТР(") > 0 Then
                        colLetter = Split(sheet.Cells(1, c).Address, "$")(1)
                        aboveRow = insertRow + r - 1
                        formula = RegReplaceLen(formula, colLetter & aboveRow)
                    End If
                    dupCell.Formula = formula
                End If
            Next
        Next

        CopyShapes sheet, sRow, eRow, insertRow

        Dim emptyRow
        emptyRow = eRow + (eRow - sRow + 1) + 1
        sheet.Rows(emptyRow).Insert xlShiftDown
        sheet.Rows(emptyRow).Clear
        sheet.Rows(emptyRow).RowHeight = 15
    Next
End Sub

Function FindHeader(sheet, headerColor)
    Dim usedRange, r, c, firstCol, lastCol, cell
    Set usedRange = sheet.UsedRange
    firstCol = 0
    lastCol = 0
    For r = 1 To 20
        If r > usedRange.Rows.Count Then Exit For
        If sheet.Rows(r).Hidden Then GoTo NextRow
        For c = 1 To usedRange.Columns.Count
            Set cell = sheet.Cells(r, c)
            If cell.Interior.Color = headerColor Then
                If firstCol = 0 Then firstCol = c
                lastCol = c
            End If
        Next
        If firstCol <> 0 Then
            Set FindHeader = sheet.Range(sheet.Cells(r, firstCol), sheet.Cells(r, lastCol))
            Exit Function
        End If
NextRow:
    Next
    Set FindHeader = Nothing
End Function

Function HasFormulas(sheet, row, startCol, endCol)
    Dim c
    For c = startCol To endCol
        If sheet.Cells(row, c).HasFormula Then
            HasFormulas = True
            Exit Function
        End If
    Next
    HasFormulas = False
End Function

Function HasDataInRange(sheet, row, startCol, endCol)
    Dim c, val
    For c = startCol To endCol
        val = sheet.Cells(row, c).Value
        If Not IsEmpty(val) Then
            HasDataInRange = True
            Exit Function
        End If
    Next
    HasDataInRange = False
End Function

Sub CopyShapes(sheet, startRow, endRow, targetStartRow)
    On Error Resume Next
    Dim targetEndRow, idx, shape, topRow, bottomRow, rowOffset, targetCell, newTop, newLeft, posKey
    targetEndRow = targetStartRow + (endRow - startRow)

    For idx = 1 To sheet.Shapes.Count
        Set shape = sheet.Shapes(idx)
        topRow = shape.TopLeftCell.Row
        bottomRow = shape.BottomRightCell.Row
        If (topRow >= targetStartRow And topRow <= targetEndRow) Or _
           (bottomRow >= targetStartRow And bottomRow <= targetEndRow) Then
            Exit Sub
        End If
    Next

    Dim dict
    Set dict = CreateObject("Scripting.Dictionary")

    For idx = 1 To sheet.Shapes.Count
        Set shape = sheet.Shapes(idx)
        topRow = shape.TopLeftCell.Row
        If topRow >= startRow And topRow <= endRow Then
            rowOffset = topRow - startRow
            Set targetCell = sheet.Cells(targetStartRow + rowOffset, shape.TopLeftCell.Column)
            newTop = targetCell.Top + (shape.Top - shape.TopLeftCell.Top)
            newLeft = targetCell.Left + (shape.Left - shape.TopLeftCell.Left)
            posKey = CStr(Round(newLeft, 2)) & ":" & CStr(Round(newTop, 2))
            If Not dict.Exists(posKey) Then
                shape.Copy
                sheet.Paste
                Dim newShape
                Set newShape = sheet.Shapes(sheet.Shapes.Count)
                newShape.Top = newTop
                newShape.Left = newLeft
                dict.Add posKey, True
            End If
        End If
    Next
End Sub

Function RegReplaceLen(formula, replacement)
    Dim re
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = "(LEN|ДЛСТР)\s*\([^)]+\)"
    RegReplaceLen = re.Replace(formula, "$1(" & replacement & ")")
End Function
