Sub Macro2()
'
' Macro2 Macro
'

'
End Sub
Sub csi_survey_count()
    ' count total in column h skipping blanks '
    surveys = ActiveSheet.Range("G21:G5000").Cells.SpecialCells(xlCellTypeConstants).Count
    ' set variables '
    Dim last_row As Long
    last_row = Cells(Rows.Count, 9).End(xlUp).Row
    ' set variable and have subtract admin cells'
    total_num = (last_row - 19)
    ' display total in message and in cell '
    MsgBox ("Total Number of Surveys" + Str(surveys))
    Range("H19").Value = surveys
End Sub
Sub GetSelectedSheetsName()
    Dim ws As Worksheet
 
    For Each ws In ActiveWindow.SelectedSheets
         MsgBox ws.Name
    Next ws
 
End Sub
Sub ssi_survey_count()
    ' count total in column h skipping blanks '
    surveys = ActiveSheet.Range("H20:H5000").Cells.SpecialCells(xlCellTypeConstants).Count
    ' set variables '
    Dim last_row As Long
    last_row = Cells(Rows.Count, 9).End(xlUp).Row
    ' set variable and have subtract admin cells'
    total_num = (last_row - 19)
    ' display total in message and in cell '
    MsgBox ("Total Number of Surveys" + Str(surveys))
    Range("I18").Value = surveys
End Sub
Sub InsertPageBreaks()
    'set variables '
    Dim I As Long, J As Long
    J = ActiveSheet.Cells(Rows.Count, "F").End(xlUp).Row
    ' For loop through data looking for changes '
    For I = J To 2 Step -1
        If Range("F" & I).Value <> Range("F" & I - 1).Value Then
            ActiveSheet.HPageBreaks.Add Before:=Range("F" & I)
        End If
    Next I
End Sub

Sub Conditional_Format()

Dim rg As Range
Dim cond1 As FormatCondition, cond2 As FormatCondition, cond3 As FormatCondition
Set rg = Range("H20", Range("H20").End(xlDown))

'clear any existing conditional formatting
rg.FormatConditions.Delete

'define the rule for each conditional format
Set cond1 = rg.FormatConditions.Add(xlCellValue, xlGreater, "=$a$1000")
Set cond2 = rg.FormatConditions.Add(xlCellValue, xlLess, "=$a$1000")
Set cond3 = rg.FormatConditions.Add(xlCellValue, xlEqual, "=$a$1000")

'define the format applied for each conditional format

With cond1
.Interior.Color = vbWhite
.Font.Color = vbWhite
End With

With cond2
.Interior.Color = vbRed
.Font.Color = vbWhite
End With

With cond3
.Interior.Color = vbWhite
.Font.Color = vbRed
End With

End Sub
‘Define Range
Dim MyRange As Range

Sub Conditional_Formatting_Total()

    'Define Range'
    Dim MyRange As Range
    Set MyRange = Range("H20:H5000")

    'Delete Existing Conditional Formatting from Range
    MyRange.FormatConditions.Delete

    'Apply Conditional Formatting
    MyRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, _
        Formula1:="=100", Formula2:="=850"
    MyRange.FormatConditions(1).Interior.Color = RGB(255, 0, 0)
End Sub

Sub Conditional_Formatting_Scores()

    'Define Range'
    Dim MyRange As Range
    Set MyRange = Range("I20:N5000")

    'Delete Existing Conditional Formatting from Range
    MyRange.FormatConditions.Delete

    'Apply Conditional Formatting
    MyRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, _
        Formula1:="=1", Formula2:="=7"
    MyRange.FormatConditions(1).Interior.Color = RGB(255, 0, 0)
End Sub




