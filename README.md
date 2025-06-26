# Randomize VBA Macro

## Overview

`Randomize` is an Excel VBA macro that helps you create a randomized, proportionate sample of rows from your dataset based on unique values in a user-selected column (called the RPG column). It colors selected rows and deletes the rest, providing an easy way to downsample large datasets while preserving representation across groups.

## Features

- User inputs the column letter containing group identifiers.
- Identifies all unique groups in the selected column.
- Randomly marks and retains a proportional sample of rows.
- Deletes rows not included in the sample.
- Works on the active worksheet.

## How to Use

1. Open your Excel file with data and a header row.
2. Press `Alt + F11` to open the VBA Editor.
3. Insert a new module and paste the macro code.
4. Run the `Randomize` macro.
5. When prompted, enter the letter of the column that contains the grouping values.
6. The macro will randomly sample rows proportionally and delete the unselected rows.

## Important Notes

- Make sure your data includes a header row (starting at row 1).
- The macro operates on the active worksheet.
- Always back up your data before running macros that delete rows.
- The fixed value `800` in the macro controls the sampling size — you can adjust it in the code to fit your needs.

## Code

```vba
Sub Randomize()
    Dim LastRow As Long, Rng As Range, List As Object
    Dim lTotalRPGs As Long
    Dim sUserCol As String
    Dim shName As String
    Dim MyNumberTemp As Double, MyNumber As Double
    Dim i As Long
    
    sUserCol = Application.InputBox("Enter the RPG column (letter)", "Select column", "A", , , , , 2)
    shName = ActiveSheet.Name
    
    LastRow = Cells(Rows.Count, sUserCol).End(xlUp).Row
    
    Set List = CreateObject("Scripting.Dictionary")
    For Each Rng In Range(sUserCol & "2:" & sUserCol & LastRow)
        If Not List.Exists(Rng.Value) Then List.Add Rng.Value, Nothing
    Next
    
    lTotalRPGs = List.Count
    
    MyNumberTemp = Round(800 / lTotalRPGs, 0)
    MyNumber = Round(LastRow / lTotalRPGs / MyNumberTemp, 0)
    
    For i = 2 To LastRow Step MyNumber + 1
        Worksheets(shName).Rows(i).Interior.ColorIndex = 4
    Next
    
    Worksheets(shName).Cells(1, 1).Select
    
    For i = LastRow To 2 Step -1
        If Worksheets(shName).Rows(i).Interior.ColorIndex <> 4 Then
            Worksheets(shName).Rows(i).Delete
        End If
    Next
    
    Worksheets(shName).Cells(1, 1).Select
End Sub
