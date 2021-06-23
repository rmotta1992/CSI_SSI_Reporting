Sub ssi_full()
'----------------------------------------------------------------------------------------------------'
    ' Convert scores to numbers '
        ' set range and use with stement '
            With Range("F2:L50000")
                .NumberFormat = "General"
                .Value = .Value
            End With
'----------------------------------------------------------------------------------------------------'
    'Insert Blank Rows'
        'Function to instert blank rows between 1 and 18'
            Rows("1:18").EntireRow.Insert
'----------------------------------------------------------------------------------------------------'
    'Insert Blank Columns'
        'Function to instert blank rows between 1 and 18'
            Columns("A:B").EntireColumn.Insert
'----------------------------------------------------------------------------------------------------'
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
            Range("I18").Value = surveys
'------------------------------------------------------------------------------------------------------'
    'Find Sum of Survey Scores'
        'Set variable '
            Dim rng As Range
        'assign the range of cells
            Set rng = Range("H20:H5000")
        'use the range in the  formula
            Total = WorksheetFunction.Sum(rng)
        'release the range object
            Set rng = Nothing
        ' place value in cell '
            Range("J18").Value = Total
'-------------------------------------------------------------------------------------------------------'
        With Range("A1:Z18")
            .Interior.Color = vbWhite
        End With
'-------------------------------------------------------------------------------------------------------'
    'Find the Average Survey Score'
        Average = (Total / surveys)
        Range("L18").Value = Average
'--------------------------------------------------------------------------------------------------------'
    'Conditional Formatting for total SSI score'
        'Define Range'
            Dim MyRange As Range
            Set MyRange = Range("H20:H5000")
        'Delete Existing Conditional Formatting from Range
            MyRange.FormatConditions.Delete
        'Apply Conditional Formatting
            MyRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, _
                Formula1:="=100", Formula2:="=850"
            MyRange.FormatConditions(1).Interior.Color = RGB(255, 0, 0)
'----------------------------------------------------------------------------------------------------------'
    'Conditional Formatting for each score '
        'Define Range'
            Dim ScoreRange As Range
            Set ScoreRange = Range("I20:N5000")
        'Delete Existing Conditional Formatting from Range
            ScoreRange.FormatConditions.Delete
        'Apply Conditional Formatting
            ScoreRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, _
                Formula1:="=1", Formula2:="=7"
            ScoreRange.FormatConditions(1).Interior.Color = RGB(255, 100, 71)
'-----------------------------------------------------------------------------------------------------------'
    'Sort the surveys by date '
            Dim sort_date As Integer, EndRw As Integer
            ' Set variables to correct cell'
                sort_date = 20
                EndRw = Range("G65500").End(xlUp).Row
                Rows(sort_date & ":" & EndRw).Select
            ' call function '
            Selection.Sort Key1:=Range("G20"), Order1:=xlAscending
'----------------------------------------------------------------------------------------------------------'
     ' Fill cells with survey count and total score red '
            Range("I18", "J18").Interior.Color = vbRed
'----------------------------------------------------------------------------------------------------------'
     ' Fill cells with survey average yellow '
             Range("L18").Interior.Color = vbYellow
'----------------------------------------------------------------------------------------------------------'
    ' Format font for cells displaying final numbers '
        'Average Survey Score'
            With Range("L18")
                .Value = .Value
                .Font.Size = 16
                .Font.Name = "Calibri"
                .Font.Bold = True
            End With
        'Total Number of Surveys'
            With Range("I18")
                .Value = .Value
                .Font.Size = 14
                .Font.Name = "Calibri"
                .Font.Bold = True
            End With
        'Total surver scores'
            With Range("J18")
                .Value = .Value
                .Font.Size = 10
                .Font.Name = "Calibri"
                .Font.Bold = True
            End With
'----------------------------------------------------------------------------------------------------------'
    ' Add sheet formatting '
        ' quater info '
            With Range("C2")
                .Value = "July 1st - Sep 30th"
                .Font.Size = 10
                .Font.Name = "Calibri"
                .Font.Bold = True
                .Interior.Color = vbWhite
            End With
        'add in normalized behavior categories '
            With Range("O2")
                .Value = "Normalized Behaviors"
                .Font.Size = 10
                .Font.Name = "Calibri"
                .Font.Bold = True
                .Interior.Color = vbWhite
            End With
            With Range("O3")
                .Value = "Understood  My Needs"
                .Font.Size = 10
                .Font.Name = "Calibri"
                .Font.Bold = True
                .Interior.Color = RGB(192, 192, 192)
                .WrapText = True
            End With
            With Range("P3")
                .Value = "Effective use of devices"
                .Font.Size = 10
                .Font.Name = "Calibri"
                .Font.Bold = True
                .Interior.Color = RGB(192, 192, 192)
                .WrapText = True
            End With
            With Range("Q3")
                .Value = "Consultant Knowledge"
                .Font.Size = 10
                .Font.Name = "Calibri"
                .Font.Bold = True
                .Interior.Color = RGB(192, 192, 192)
                .WrapText = True
            End With
            With Range("R3")
                .Value = "Straight Answer on Price"
                .Font.Size = 10
                .Font.Name = "Calibri"
                .Font.Bold = True
                .Interior.Color = RGB(192, 192, 192)
                .WrapText = True
            End With
            With Range("S3")
                .Value = "Changed Advertised Price"
                .Font.Size = 10
                .Font.Name = "Calibri"
                .Font.Bold = True
                .Interior.Color = RGB(192, 192, 192)
                .WrapText = True
            End With
            With Range("T3")
                .Value = "Not pushy to sell vehicle"
                .Font.Size = 10
                .Font.Name = "Calibri"
                .Font.Bold = True
                .Interior.Color = RGB(192, 192, 192)
                .WrapText = True
            End With
            With Range("U3")
                .Value = "Not pushy to sell F&I products"
                .Font.Size = 10
                .Font.Name = "Calibri"
                .Font.Bold = True
                .Interior.Color = RGB(192, 192, 192)
                .WrapText = True
            End With
            With Range("V3")
                .Value = "Did not add items"
                .Font.Size = 10
                .Font.Name = "Calibri"
                .Font.Bold = True
                .Interior.Color = RGB(192, 192, 192)
                .WrapText = True
            End With
            With Range("W3")
                .Value = "No paper errors"
                .Font.Size = 10
                .Font.Name = "Calibri"
                .Font.Bold = True
                .Interior.Color = RGB(192, 192, 192)
                .WrapText = True
            End With
            With Range("X3")
                .Value = "Follow Up After Sale"
                .Font.Size = 10
                .Font.Name = "Calibri"
                .Font.Bold = True
                .Interior.Color = RGB(192, 192, 192)
                .WrapText = True
            End With
            With Range("Y3")
                .Value = "Delivered Clean Vehicle"
                .Font.Size = 10
                .Font.Name = "Calibri"
                .Font.Bold = True
                .Interior.Color = RGB(192, 192, 192)
                .WrapText = True
            End With
            With Range("Z3")
                .Value = "Explain Safety Features (coming soon)"
                .Font.Size = 10
                .Font.Name = "Calibri"
                .Font.Bold = True
                .Interior.Color = RGB(192, 192, 192)
                .WrapText = True
            End With
            'add in normalized behavior goals '
            With Range("O4")
                .Value = "95.00%"
                .Font.Size = 10
                .Font.Name = "Calibri"
                .Font.Bold = False
                .Interior.Color = vbWhite
            End With
            'add in target info, quarter info and perfet socres needed'
                With Range("E17")
                    .Value = "Q2"
                    .Font.Size = 16
                    .Font.Name = "Calibri"
                    .Font.Bold = False
                    .HorizontalAlignment = xlCenter
                End With
                With Range("E18")
                    .Value = "SSI Target= "
                    .Font.Size = 16
                    .Font.Name = "Calibri"
                    .Font.Bold = False
                    .HorizontalAlignment = xlRight
                    .Interior.Color = vbWhite
                End With
                With Range("G16")
                    .Value = "Need"
                    .Font.Size = 11
                    .Font.Name = "calibri"
                    .Font.Bold = True
                    .HorizontalAlignment = xlCenter
                    .Interior.Color = vbWhite
                End With
                With Range("G18")
                    .Value = "Perfect"
                    .Font.Size = 11
                    .Font.Name = "calibri"
                    .Font.Bold = True
                    .HorizontalAlignment = xlCenter
                    .Interior.Color = vbWhite
                End With
                With Range("F18")
                    .Value = "950"
                    .Font.Size = 16
                    .Font.Name = "calibri"
                    .HorizontalAlignment = xlLeft
                    .Font.Bold = True
                End With
                
    ' display dates and survey number '
        With Range("A20:A100")
            .Value = "July"
            .Font.Size = 11
            .Font.Name = "calibri"
            .Interior.Color = vbWhite
        End With
        With Range("A101:A200")
            .Value = "Aug"
            .Font.Size = 11
            .Font.Name = "calibri"
            .Interior.Color = vbWhite
        End With
        With Range("A201:A300")
            .Value = "Sep"
            .Font.Size = 11
            .Font.Name = "calibri"
            .Interior.Color = vbWhite
        End With
        
                
'----------------------------------------------------------------------------------------------------------'
   
'----------------------------------------------------------------------------------------------------------'
    Range("B20") = 1
    Range("B20:B" & Range("H" & Rows.Count).End(xlUp).Row).DataSeries , xlDataSeriesLinear
'----------------------------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------------------------'
    With Range("N18")
        .Value = Range("L18") - 950
        .Font.Size = 11
        .Font.Name = "calibri"
        .Font.Bold = True
    End With
   



End Sub