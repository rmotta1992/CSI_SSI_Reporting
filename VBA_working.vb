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

Sub Conditional_Formatting_Total_SSI()

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

Sub Conditional_Formatting_Scores_SSI()

    'Define Range'
    Dim MyRange As Range
    Set MyRange = Range("I20:N5000")

    'Delete Existing Conditional Formatting from Range
    MyRange.FormatConditions.Delete

    'Apply Conditional Formatting
    MyRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, _
        Formula1:="=1", Formula2:="=7"
    MyRange.FormatConditions(1).Interior.Color = RGB(255, 100, 71)
End Sub
Sub Conditional_Formatting_Total_CSI()

    'Define Range'
    Dim MyRange As Range
    Set MyRange = Range("E20:E5000")

    'Delete Existing Conditional Formatting from Range
    MyRange.FormatConditions.Delete

    'Apply Conditional Formatting
    MyRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, _
        Formula1:="=100", Formula2:="=875"
    MyRange.FormatConditions(1).Interior.Color = RGB(255, 0, 0)
End Sub

Sub Conditional_Formatting_Scores_CSI()
    'Define Range'
    Dim MyRange As Range
    Set MyRange = Range("H20:L5000")

    'Delete Existing Conditional Formatting from Range
    MyRange.FormatConditions.Delete

    'Apply Conditional Formatting
    MyRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, _
        Formula1:="=1", Formula2:="=7"
    MyRange.FormatConditions(1).Interior.Color = RGB(255, 100, 100)
End Sub
Sub sbInsertingRows()
    Rows("1:18").EntireRow.Insert
End Sub
Sub calculate_csi_total()
  
    ' Create variables for the Price, Tax, Quantity, and Total
    Dim CSI As Double
    

    ' Retrieve and store the data values in each variable
    CSI = Range("E20:E5000").Value
    Tax = Range("C2").Value
    Quantity = Range("D2").Value

    ' Calculate the total by using each of the variables
    Total = Price * (1 + Tax) * Quantity

    ' Create a Message Box for the Total and insert into cell
    MsgBox ("Your total is $" + Str(Total))
    Range("E2").Value = Total
End Sub
    