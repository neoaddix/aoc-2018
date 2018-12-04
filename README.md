# aoc-2018
Advent of Code 2018 - Excel creativity


## --- Day 1: Chronal Calibration ---

The first part of the assignment was quite easy in Excel: just copy all the numbers in 1 column and in the next column add a formula calculate each next value `B2=A2+B1`

The second part is unpredictable because we need to calculate which number is repeated first. I could not see a way to predict when this would happen, so I just created a macro that keeps on doing the above calculation, checks if the number is already present in the previous results and when that happens shows an alert window. This took Excel about 2 hours ;-)

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
