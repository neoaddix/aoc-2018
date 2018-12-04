# aoc-2018
Advent of Code 2018 - Excel creativity


## Day 1

```
Sub copy_columns()

Dim col As Integer
Dim row As Integer
Dim EvalTarget As Long
Dim ws As Worksheet, EvalRange As Range

For col = 3 To 150

    For row = 1 To 951
        Cells(row, col).Select
        Selection.Copy
        Cells(row, col).Offset(0, 1).Select
        ActiveSheet.Paste
        EvalTarget = Cells(row, col).Offset(0, 1).Value
        Set EvalRange = Range(Cells(1, 2), Cells(951, col).Offset(0, 1))

   If WorksheetFunction.CountIf(EvalRange, EvalTarget) > 1 Then
        MsgBox EvalTarget & " already exists on this sheet."
        Exit Sub
    End If

    Next row
Next col

End Sub
```
