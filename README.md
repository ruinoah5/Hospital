Hospital
========
Public currentRow As Integer
Public finalRow As Integer
Public col As Integer
Public cellValue As String

Public cgcol As Integer
Public cgcellValue As String

Public subjectcol As Integer
Public subjectcellValue As String

Public pastIndex As Integer

Sub addRow(row As Integer)

    Cells(row, col).Offset(1).EntireRow.Insert
    finalRow = finalRow + 1
    currentRow = currentRow + 1
    

End Sub

Sub hosCgSub(startIndex As Integer, stopIndex As Integer)
    Cells(currentRow - 1, col).Value = Mid(cellValue, startIndex, stopIndex)
    Cells(currentRow - 1, cgcol).Value = cgcellValue
    Cells(currentRow - 1, subjectcol).Value = subjectcellValue
    
    pastIndex = stopIndex + 1
    
End Sub

Sub slash(startIndex As Integer)
    Cursor = InStr(startIndex, cellValue, "/")
    If Cursor <> 0 Then
        Call addRow(currentRow)
        Call hosCgSub(startIndex, Cursor - 1)
        slash (Cursor + 1)
    End If
        Call hosCgSub(pastIndex, 15)
End Sub

Sub ExpandBtn_Click()


currentRow = 1
finalRow = 5

col = 5
cgcol = 3
subjectcol = 4
pastIndex = 1

While currentRow <= finalRow
    cellValue = Worksheets("Sheet1").Cells(currentRow, col).Value
    cgcellValue = Worksheets("Sheet1").Cells(currentRow, cgcol).Value
    subjectcellValue = Worksheets("Sheet1").Cells(currentRow, subjectcol).Value
    
    slash (1)

    currentRow = currentRow + 1

Wend

End Sub





