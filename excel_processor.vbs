Option Explicit
Dim filePath, headerColor
If WScript.Arguments.Count < 2 Then
    WScript.StdErr.WriteLine "Usage: cscript excel_processor.vbs <file> <header_color>"
    WScript.Quit 1
End If
filePath = WScript.Arguments.Item(0)
headerColor = CLng(WScript.Arguments.Item(1))
Dim xl, wb
On Error Resume Next
Set xl = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    WScript.StdErr.WriteLine "Cannot create Excel: " & Err.Description
    WScript.Quit 1
End If
On Error GoTo 0
xl.Visible = False
xl.DisplayAlerts = False
xl.ScreenUpdating = False
On Error GoTo Cleanup
Set wb = xl.Workbooks.Open(filePath)
Dim sh
For Each sh In wb.Worksheets
    ProcessSheet sh, headerColor
Next
wb.Save
Cleanup:
If Err.Number <> 0 Then
    WScript.StdErr.WriteLine "Error: " & Err.Description
End If
If Not wb Is Nothing Then wb.Close False
If Not xl Is Nothing Then xl.Quit
If Err.Number <> 0 Then WScript.Quit 1

Sub ProcessSheet(s, headerColor)
    Dim usedRange, lastRow, row
    Set usedRange = s.UsedRange
    If usedRange Is Nothing Then Exit Sub
    lastRow = usedRange.Row + usedRange.Rows.Count - 1
    Dim headerRow
    headerRow = FindHeader(s, headerColor)
    If headerRow = 0 Then Exit Sub
    row = headerRow + 1
    Do While row <= lastRow
        If s.Rows(row).Hidden Then
            row = row + 1
        ElseIf HasData(s, row, usedRange.Columns.Count) Then
            s.Rows(row).Copy
            s.Rows(row + 1).Insert
            s.Rows(row + 2).Insert
            s.Rows(row + 2).ClearContents
            row = row + 3
            lastRow = lastRow + 2
        Else
            row = row + 1
        End If
    Loop
End Sub

Function HasData(s, r, colsCount)
    Dim c
    For c = 1 To colsCount
        If s.Cells(r, c).Value <> "" Then
            HasData = True
            Exit Function
        End If
    Next
    HasData = False
End Function

Function FindHeader(s, headerColor)
    Dim usedRange, r, c
    Set usedRange = s.UsedRange
    If usedRange Is Nothing Then
        FindHeader = 0
        Exit Function
    End If
    For r = 1 To 20
        If r > usedRange.Rows.Count Then Exit For
        If Not s.Rows(r).Hidden Then
            For c = 1 To usedRange.Columns.Count
                If s.Cells(r, c).Interior.Color = headerColor Then
                    FindHeader = r
                    Exit Function
                End If
            Next
        End If
    Next
    FindHeader = 0
End Function

