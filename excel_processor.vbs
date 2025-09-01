Option Explicit

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

' Logging helper
Sub Log(msg)
    WScript.Echo msg
End Sub

' Helper for minimum
Function Min(a, b)
    If a < b Then
        Min = a
    Else
        Min = b
    End If
End Function

Function Max(a, b)
    If a > b Then
        Max = a
    Else
        Max = b
    End If
End Function

Function RegexReplaceLen(formula, repl)
    Dim re
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "(LEN|ДЛСТР)\s*\([^)]+\)"
    re.IgnoreCase = True
    RegexReplaceLen = re.Replace(formula, "$1(" & repl & ")")
End Function

Sub DeleteShapesInRow(sheet, row)
    Dim i, shape, mainRow
    For i = sheet.Shapes.Count To 1 Step -1
        Set shape = sheet.Shapes(i)
        mainRow = RowWithMaxShapeOverlap(sheet, shape)
        If mainRow = row Then shape.Delete
    Next
End Sub

Sub CopyShapesForRow(sheet, shapesList, targetRow)
    Dim i, item, shape, col, deltaTop, deltaLeft, targetCell, newShape
    For i = 0 To shapesList.Count - 1
        item = shapesList(i)
        Set shape = item(0)
        col = item(2)
        deltaTop = shape.Top - shape.TopLeftCell.Top
        deltaLeft = shape.Left - shape.TopLeftCell.Left
        Set targetCell = sheet.Cells(targetRow, col)
        Set newShape = shape.Duplicate
        newShape.Top = targetCell.Top + deltaTop
        newShape.Left = targetCell.Left + deltaLeft
    Next
End Sub

Function RowWithMaxShapeOverlap(sheet, shape)
    Dim top, bottom, firstRow, lastRow, row, rowTop, rowBottom, overlap, maxOverlap, mainRow
    top = shape.Top
    bottom = shape.Top + shape.Height
    firstRow = shape.TopLeftCell.Row
    lastRow = shape.BottomRightCell.Row
    maxOverlap = 0
    mainRow = firstRow
    For row = firstRow To lastRow
        rowTop = sheet.Rows(row).Top
        rowBottom = rowTop + sheet.Rows(row).Height
        overlap = Min(bottom, rowBottom) - Max(top, rowTop)
        If overlap > maxOverlap Then
            maxOverlap = overlap
            mainRow = row
        End If
    Next
    RowWithMaxShapeOverlap = mainRow
End Function

Function MapShapesByRow(sheet)
    Dim rowToShapes, i, shape, mainRow, col, shapeName, list
    Set rowToShapes = CreateObject("Scripting.Dictionary")
    For i = 1 To sheet.Shapes.Count
        Set shape = sheet.Shapes(i)
        mainRow = RowWithMaxShapeOverlap(sheet, shape)
        col = shape.TopLeftCell.Column
        shapeName = shape.Name
        If Not rowToShapes.Exists(mainRow) Then
            Set list = CreateObject("System.Collections.ArrayList")
            rowToShapes.Add mainRow, list
        End If
        rowToShapes(mainRow).Add Array(shape, shapeName, col)
    Next
    Set MapShapesByRow = rowToShapes
End Function

Function HasFormulas(sheet, row, startCol, endCol)
    Dim col
    If sheet.Rows(row).Hidden Then
        HasFormulas = False
        Exit Function
    End If
    For col = startCol To endCol
        If sheet.Cells(row, col).HasFormula Then
            HasFormulas = True
            Exit Function
        End If
    Next
    HasFormulas = False
End Function

Function HasDataInRange(sheet, row, startCol, endCol)
    Dim col, val
    If sheet.Rows(row).Hidden Then
        HasDataInRange = False
        Exit Function
    End If
    For col = startCol To endCol
        val = sheet.Cells(row, col).Value
        If Not IsEmpty(val) Then
            HasDataInRange = True
            Exit Function
        End If
    Next
    HasDataInRange = False
End Function

Function FindHeaderRange(sheet, headerRow)
    Dim usedRange, firstCol, lastCol, col, cellValue
    Set usedRange = sheet.UsedRange
    firstCol = 0
    lastCol = 0
    For col = 1 To usedRange.Columns.Count
        cellValue = sheet.Cells(headerRow, col).Value
        If Not IsEmpty(cellValue) Then
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

Function FindHeader(sheet, headerColor)
    Dim usedRange, row, col, cell
    Set usedRange = sheet.UsedRange
    For row = 1 To Min(20, usedRange.Rows.Count)
        If sheet.Rows(row).Hidden Then
        Else
            For col = 1 To usedRange.Columns.Count
                Set cell = sheet.Cells(row, col)
                If cell.Interior.Color = headerColor Then
                    Set FindHeader = FindHeaderRange(sheet, row)
                    Exit Function
                End If
            Next
        End If
    Next
    Set FindHeader = Nothing
End Function

Sub RestructureSheet(sheet, headerRange)
    Dim headerRow, headerStartCol, headerEndCol, usedRange, lastRow, headerHeight
    Dim shapesMap, dataBlocks, row, block, startRow, endRow
    Dim blockRowMap, shift, i, sourceRange, insertRow, rowOffset
    Dim srcRow, dstRow, col, sourceCell, dupCell, formula, colLetter, aboveRow
    Dim emptyRow, headerInsertRow, targetRange, targetCell, finalMap, newRow
    Dim origRow, shapesInfo, found, s, name, col0, item, rowKey, lst, deltaTop, deltaLeft, newShape

    headerRow = headerRange.Row
    headerStartCol = headerRange.Column
    headerEndCol = headerRange.Column + headerRange.Columns.Count - 1
    Set usedRange = sheet.UsedRange
    lastRow = usedRange.Row + usedRange.Rows.Count - 1
    headerHeight = sheet.Rows(headerRow).RowHeight

    Set shapesMap = MapShapesByRow(sheet)
    Set dataBlocks = CreateObject("System.Collections.ArrayList")

    row = headerRow + 1
    Do While row <= lastRow
        If sheet.Rows(row).Hidden Then
            row = row + 1
        ElseIf HasDataInRange(sheet, row, headerStartCol, headerEndCol) Then
            startRow = row
            endRow = row
            If row + 1 <= lastRow And Not sheet.Rows(row + 1).Hidden And _
                    HasFormulas(sheet, row + 1, headerStartCol, headerEndCol) Then
                endRow = row + 1
                row = row + 2
            Else
                row = row + 1
            End If
            dataBlocks.Add Array(startRow, endRow)
        Else
            row = row + 1
        End If
    Loop

    If dataBlocks.Count = 0 Then
        Log "No data rows found after header"
        Exit Sub
    End If

    Set blockRowMap = CreateObject("Scripting.Dictionary")
    shift = 0
    For i = 0 To dataBlocks.Count - 1
        block = dataBlocks(i)
        startRow = block(0)
        endRow = block(1)
        blockRowMap.Add startRow, startRow + shift
        shift = shift + (endRow - startRow + 1) + 1
    Next

    sheet.Application.ScreenUpdating = False
    sheet.Application.Calculation = -4135

    For i = dataBlocks.Count - 1 To 0 Step -1
        block = dataBlocks(i)
        startRow = block(0)
        endRow = block(1)
        Set sourceRange = sheet.Range(startRow & ":" & endRow)
        sourceRange.Copy
        insertRow = endRow + 1
        sheet.Rows(insertRow).Insert -4121
        For rowOffset = 0 To (endRow - startRow)
            srcRow = startRow + rowOffset
            dstRow = insertRow + rowOffset
            DeleteShapesInRow sheet, dstRow
            For col = headerStartCol To headerEndCol
                Set sourceCell = sheet.Cells(srcRow, col)
                Set dupCell = sheet.Cells(dstRow, col)
                If sourceCell.HasFormula Then
                    formula = sourceCell.Formula
                    If InStr(UCase(formula), "LEN(") > 0 Or InStr(UCase(formula), "ДЛСТР(") > 0 Then
                        colLetter = Split(sheet.Cells(1, col).Address(False, False), "$")(0)
                        aboveRow = dstRow - 1
                        formula = RegexReplaceLen(formula, colLetter & aboveRow)
                    End If
                    dupCell.Formula = formula
                End If
            Next
            If shapesMap.Exists(srcRow) Then
                CopyShapesForRow sheet, shapesMap(srcRow), dstRow
            End If
        Next
        emptyRow = endRow + (endRow - startRow + 1) + 1
        sheet.Rows(emptyRow).Insert -4121
        sheet.Rows(emptyRow).Clear
        sheet.Rows(emptyRow).RowHeight = 15

        If i > 0 Then
            headerInsertRow = startRow
            DeleteShapesInRow sheet, headerInsertRow
            sheet.Rows(headerInsertRow).Insert -4121
            sheet.Rows(headerInsertRow).Clear
            headerRange.Copy
            Set targetRange = sheet.Range(sheet.Cells(headerInsertRow, headerStartCol), sheet.Cells(headerInsertRow, headerEndCol))
            targetRange.PasteSpecial -4104
            sheet.Rows(headerInsertRow).RowHeight = headerHeight
        End If
    Next

    Log "=== [SHAPE POSTFIX] Проверка восстановления shape на актуальных строках ==="
    Set finalMap = MapShapesByRow(sheet)
    For Each origRow In shapesMap.Keys
        shapesInfo = shapesMap(origRow)
        If blockRowMap.Exists(origRow) Then
            newRow = blockRowMap(origRow)
        Else
            newRow = origRow
        End If
        found = False
        If finalMap.Exists(newRow) Then
            For i = 0 To finalMap(newRow).Count - 1
                item = finalMap(newRow)(i)
                name = item(1)
                col0 = item(2)
                Dim j, item2
                For j = 0 To shapesInfo.Count - 1
                    item2 = shapesInfo(j)
                    If item2(1) = name And item2(2) = col0 Then
                        found = True
                    End If
                Next
            Next
        End If
        If Not found Then
            For i = 0 To shapesInfo.Count - 1
                item2 = shapesInfo(i)
                name = item2(1)
                col0 = item2(2)
                For Each rowKey In finalMap.Keys
                    lst = finalMap(rowKey)
                    For j = 0 To lst.Count - 1
                        item = lst(j)
                        If item(1) = name And item(2) = col0 Then
                            deltaTop = item(0).Top - item(0).TopLeftCell.Top
                            deltaLeft = item(0).Left - item(0).TopLeftCell.Left
                            Set targetCell = sheet.Cells(newRow, col0)
                            Set newShape = item(0).Duplicate
                            newShape.Top = targetCell.Top + deltaTop
                            newShape.Left = targetCell.Left + deltaLeft
                        End If
                    Next
                Next
            Next
        End If
    Next

    sheet.Application.CutCopyMode = False
    sheet.Application.Calculation = -4105
    sheet.Application.ScreenUpdating = True
End Sub

Function CanProcessV2(sheet, headerColor)
    Dim usedRange, row, col, cell, yellowHeadersCount
    Set usedRange = sheet.UsedRange
    If usedRange Is Nothing Then
        CanProcessV2 = False
        Exit Function
    End If
    yellowHeadersCount = 0
    For row = 1 To Min(50, usedRange.Rows.Count)
        If sheet.Rows(row).Hidden Then
        Else
            For col = 1 To Min(10, usedRange.Columns.Count)
                Set cell = sheet.Cells(row, col)
                If cell.Interior.Color = headerColor And Not IsEmpty(cell.Value) Then
                    yellowHeadersCount = yellowHeadersCount + 1
                    Exit For
                End If
            Next
        End If
    Next
    CanProcessV2 = (yellowHeadersCount >= 2)
End Function

Function FindAllBlocks(sheet, usedRange, headerColor)
    Dim blocks, currentRow, lastRow, colsCount
    Dim isHeader, col, cell, block, dataGroups, currentGroup
    Dim isNextHeader, hasData
    Set blocks = CreateObject("System.Collections.ArrayList")
    currentRow = 1
    lastRow = usedRange.Row + usedRange.Rows.Count - 1
    colsCount = usedRange.Columns.Count
    Do While currentRow <= lastRow
        If sheet.Rows(currentRow).Hidden Then
            currentRow = currentRow + 1
        Else
            isHeader = False
            For col = 1 To Min(10, colsCount)
                Set cell = sheet.Cells(currentRow, col)
                If cell.Interior.Color = headerColor And Not IsEmpty(cell.Value) Then
                    isHeader = True
                    Exit For
                End If
            Next
            If isHeader Then
                Set block = CreateObject("Scripting.Dictionary")
                block("header_row") = currentRow
                Set dataGroups = CreateObject("System.Collections.ArrayList")
                block("data_groups") = dataGroups
                currentRow = currentRow + 1
                Set currentGroup = CreateObject("System.Collections.ArrayList")
                Do While currentRow <= lastRow
                    If sheet.Rows(currentRow).Hidden Then
                        currentRow = currentRow + 1
                    Else
                        isNextHeader = False
                        For col = 1 To Min(10, colsCount)
                            Set cell = sheet.Cells(currentRow, col)
                            If cell.Interior.Color = headerColor And Not IsEmpty(cell.Value) Then
                                isNextHeader = True
                                Exit For
                            End If
                        Next
                        If isNextHeader Then
                            If currentGroup.Count > 0 Then dataGroups.Add currentGroup
                            Exit Do
                        End If
                        hasData = False
                        For col = 1 To colsCount
                            If Not IsEmpty(sheet.Cells(currentRow, col).Value) Or sheet.Cells(currentRow, col).HasFormula Then
                                hasData = True
                                Exit For
                            End If
                        Next
                        If Not hasData Then
                            If currentGroup.Count > 0 Then
                                dataGroups.Add currentGroup
                                Set currentGroup = CreateObject("System.Collections.ArrayList")
                            End If
                            currentRow = currentRow + 1
                            Exit Do
                        Else
                            currentGroup.Add currentRow
                            currentRow = currentRow + 1
                        End If
                    End If
                Loop
                If currentGroup.Count > 0 Then dataGroups.Add currentGroup
                If dataGroups.Count > 0 Then blocks.Add block
            Else
                currentRow = currentRow + 1
            End If
        End If
    Loop
    Set FindAllBlocks = blocks
End Function

Sub CopyShapesInRange(sheet, startRow, endRow, targetStartRow)
    On Error Resume Next
    Dim targetEndRow, idx, shape, shapeRow, rowOffset, targetCell, newTop, newLeft, posKey
    Dim shapesCount, existingPositions, newShape
    targetEndRow = targetStartRow + (endRow - startRow)
    For idx = 1 To sheet.Shapes.Count
        Set shape = sheet.Shapes(idx)
        shapeRow = shape.TopLeftCell.Row
        If targetStartRow <= shapeRow And shapeRow <= targetEndRow Then Exit Sub
    Next
    shapesCount = sheet.Shapes.Count
    Set existingPositions = CreateObject("Scripting.Dictionary")
    For idx = 1 To shapesCount
        Set shape = sheet.Shapes(idx)
        shapeRow = shape.TopLeftCell.Row
        If startRow <= shapeRow And shapeRow <= endRow And Not sheet.Rows(shapeRow).Hidden Then
            rowOffset = shapeRow - startRow
            Set targetCell = sheet.Cells(targetStartRow + rowOffset, shape.TopLeftCell.Column)
            newTop = targetCell.Top + (shape.Top - shape.TopLeftCell.Top)
            newLeft = targetCell.Left + (shape.Left - shape.TopLeftCell.Left)
            posKey = CStr(Round(newLeft,2)) & ":" & CStr(Round(newTop,2))
            If Not existingPositions.Exists(posKey) Then
                shape.Copy
                sheet.Paste
                Set newShape = sheet.Shapes(sheet.Shapes.Count)
                newShape.Top = newTop
                newShape.Left = newLeft
                existingPositions.Add posKey, True
            End If
        End If
    Next
End Sub

Sub DuplicateBlockRows(sheet, block, usedRange, group)
    Dim colsCount, groupSize, insertRow, i, sourceRow, targetRow, col, cell, formula
    Dim colLetter, refRow
    colsCount = usedRange.Columns.Count
    If group.Count = 0 Then Exit Sub
    groupSize = group.Count
    insertRow = group(groupSize - 1) + 1
    For i = 1 To groupSize
        sheet.Rows(insertRow).Insert -4121
    Next
    For i = 0 To groupSize - 1
        sourceRow = group(i)
        targetRow = insertRow + i
        sheet.Rows(sourceRow).Copy
        sheet.Rows(targetRow).PasteSpecial -4104
        For col = 1 To colsCount
            Set cell = sheet.Cells(targetRow, col)
            If cell.HasFormula Then
                formula = cell.Formula
                If InStr(UCase(formula), "LEN(") > 0 Or InStr(UCase(formula), "ДЛСТР(") > 0 Then
                    colLetter = Split(sheet.Cells(1, col).Address(False, False), "$")(0)
                    If i > 0 Then
                        refRow = targetRow - 1
                    Else
                        refRow = targetRow
                    End If
                    formula = RegexReplaceLen(formula, colLetter & refRow)
                    cell.Formula = formula
                End If
            End If
        Next
    Next
    CopyShapesInRange sheet, group(0), group(groupSize - 1), insertRow
    sheet.Application.CutCopyMode = False
End Sub

Sub ProcessSheetV2(sheet, headerColor)
    Log "Processing sheet '" & sheet.Name & "' with V2 method"
    Dim usedRange, blocks, totalGroups, i, block, processedGroups, j, group
    Set usedRange = sheet.UsedRange
    Set blocks = FindAllBlocks(sheet, usedRange, headerColor)
    If blocks.Count = 0 Then
        Log "No data blocks found"
        Exit Sub
    End If
    sheet.Application.ScreenUpdating = False
    sheet.Application.Calculation = -4135
    totalGroups = 0
    For i = 0 To blocks.Count - 1
        Set block = blocks(i)
        totalGroups = totalGroups + block("data_groups").Count
    Next
    processedGroups = 0
    For i = blocks.Count - 1 To 0 Step -1
        Set block = blocks(i)
        For j = block("data_groups").Count - 1 To 0 Step -1
            Set group = block("data_groups")(j)
            DuplicateBlockRows sheet, block, usedRange, group
            processedGroups = processedGroups + 1
        Next
    Next
    sheet.Application.Calculation = -4105
    sheet.Application.ScreenUpdating = True
    Log "Processed " & blocks.Count & " blocks"
End Sub

Sub ProcessSheetV1(sheet, filename, headerColor)
    Dim sheetName, headerRange
    sheetName = sheet.Name
    Log "Excel " & filename & " - Sheet '" & sheetName & "' - searching for header..."
    Set headerRange = FindHeader(sheet, headerColor)
    If Not headerRange Is Nothing Then
        Log "Excel " & filename & " - Sheet '" & sheetName & "' - found header at row " & headerRange.Row & ", duplicating rows..."
        RestructureSheet sheet, headerRange
    Else
        Log "Excel " & filename & " - Sheet '" & sheetName & "' - no header found, skipping."
    End If
End Sub

Sub ProcessFile(filePath, headerColor, dryRun)
    Dim sourceFile, outputFolder, outputFile, excel, wb, totalSheets, sheetIndex, sheet
    sourceFile = fso.GetFile(filePath)
    outputFolder = fso.BuildPath(fso.GetParentFolderName(filePath), "Deeva")
    If Not fso.FolderExists(outputFolder) Then fso.CreateFolder outputFolder
    outputFile = fso.BuildPath(outputFolder, sourceFile.Name)
    Log "Starting processing: " & filePath
    If Not dryRun Then
        Log "Copying file to: " & outputFile
        fso.CopyFile filePath, outputFile, True
        Set excel = CreateObject("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        Set wb = excel.Workbooks.Open(outputFile)
        totalSheets = wb.Sheets.Count
        For sheetIndex = 1 To totalSheets
            Set sheet = wb.Sheets(sheetIndex)
            Log "Processing sheet " & sheetIndex & "/" & totalSheets & ": '" & sheet.Name & "'"
            If CanProcessV2(sheet, headerColor) Then
                Log "Using V2 method for sheet '" & sheet.Name & "'"
                ProcessSheetV2 sheet, headerColor
            Else
                Log "Using V1 method for sheet '" & sheet.Name & "'"
                ProcessSheetV1 sheet, sourceFile.Name, headerColor
            End If
            Log "Sheet '" & sheet.Name & "' - Done."
        Next
        Log "Saving file..."
        wb.Save
        wb.Close False
        excel.Quit
        Log "Successfully saved to: " & outputFile
    Else
        Log "[DRY RUN] Would save to: " & outputFile
    End If
End Sub

Dim filePath, headerColor, dryRun
If WScript.Arguments.Count < 1 Then
    WScript.Echo "Usage: cscript excel_processor.vbs <filepath> [headerColor] [dryRun]"
    WScript.Quit 1
End If
filePath = WScript.Arguments(0)
If WScript.Arguments.Count >= 2 Then
    headerColor = CLng(WScript.Arguments(1))
Else
    headerColor = 65535
End If
If WScript.Arguments.Count >= 3 Then
    dryRun = (LCase(WScript.Arguments(2)) = "true")
Else
    dryRun = False
End If

ProcessFile filePath, headerColor, dryRun
