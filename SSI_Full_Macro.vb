Sub ssi_full()
    'Insert Blank Rows'
        'Function to instert blank rows between 1 and 18'
            Rows("1:18").EntireRow.Insert
'--------------------------------------------------------------------------------------------------'
    'Insert Blank Columns'
        'Function to instert blank rows between 1 and 18'
            Columns("A:B").EntireColumn.Insert
'--------------------------------------------------------------------------------------------------'
    'Count Total Number of Surveys' 
        ' count total in column h skipping blanks '
            surveys = ActiveSheet.Range("F20:F5000").Cells.SpecialCells(xlCellTypeConstants).Count
        ' set variables '
            Dim last_row As Long
            last_row = Cells(Rows.Count, 7).End(xlUp).Row
        ' set variable and have subtract admin cells'
            total_num = (last_row - 19)
        ' display total in message and in cell '
            MsgBox ("Total Number of Surveys" + Str(surveys))
            Range("g18").Value = surveys
'---------------------------------------------------------------------------------------------------'
    'Find Sum of Survey Scores'
        'Set variable ' 
            Dim rng As Range
        'assign the range of cells
            Set rng = Range("F20:F5000")
        'use the range in the formula '
            Total = WorksheetFunction.Sum(rng)
        'release the range object
            Set rng = Nothing
        ' place value in cell '
            Range("f18").Value = Total
'-----------------------------------------------------------------------------------------'
    'Find the Average Survey Score'
        Average = (Total / surveys)
        Range("H18").Value = Average
'-------------------------------------------------------------------------------------------'   
    'Conditional Formatting for total SSI score'
        'Define Range'
            Dim MyRange As Range
            Set MyRange = Range("F20:F5000")
        'Delete Existing Conditional Formatting from Range
            MyRange.FormatConditions.Delete
        'Apply Conditional Formatting
            MyRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, _
                Formula1:="=100", Formula2:="=850"
            MyRange.FormatConditions(1).Interior.Color = RGB(255, 0, 0)
'-------------------------------------------------------------------------------------------'   
    'Conditional Formatting for each score '
        'Define Range'
            Dim ScoreRange As Range
            Set ScoreRange = Range("G20:N5000")
        'Delete Existing Conditional Formatting from Range
            ScoreRange.FormatConditions.Delete
        'Apply Conditional Formatting
            ScoreRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, _
                Formula1:="=1", Formula2:="=7"
            ScoreRange.FormatConditions(1).Interior.Color = RGB(255, 100, 71)
'-----------------------------------------------------------------------------------------'
End Sub