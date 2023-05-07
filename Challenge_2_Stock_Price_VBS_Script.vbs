
' Bootcamp: UTA-VIRT-DATA-PT-04-2023-U-LOLC-MTTH
' Module 2:VBA Scripting
' Module 2 Challenge
' 	Create a script that loops through all the stocks for one year and outputs
' 	two summary tables then performs the same script on each Excel tab in workbook

Sub Stocks()

'Declare Workbook counter variable
Dim ws_count As Integer
'Declare Stock Price variables
Dim ticker As String
Dim open_price As Double
Dim close_price As Double
Dim vol As Double
'Declare summary table row counter variable
Dim sumrow As Integer
'Declare match formula variable
Dim find As Integer
'Declare macro tracking variables
Dim start_macro As Date
Dim stop_macro As Date
Dim run_min As Integer
Dim run_sec As Integer

'Get number of worksheets
ws_count = ThisWorkbook.Worksheets.Count

'Track when macro started
start_macro = Now

'Loop to activate first worksheet and move to next worksheet when this worksheet is complete
For active_ws = 1 To ws_count

    'Activate current worksheet and set starting point
    Worksheets(active_ws).Activate
    Range("A1").Select

'------------------------------
'Populate header rows
'------------------------------
    
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    
    'Set variables
    LastRow = Cells(Rows.Count, "A").End(xlUp).Row
    vol = 0
    sumrow = 1

    
'-------------------------------------------------------
'Sort data table by ticker and date in ascending order
'-------------------------------------------------------
    
    Range("A1:G" & LastRow).Sort Key1:=Range("A1"), _
            Order1:=xlAscending, _
            Key2:=Range("A2"), _
            Order2:=xlAscending, _
            Header:=xlYes
    
'------------------------------
'Populate first summary table
'------------------------------
    
    'Loop to iterate through data table and populate first summary table
    For i = 2 To LastRow
    
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        
                'Loop to get the last row of summary table as it is being populated
                Do While Cells(sumrow, 9).Value <> ""
                    sumrow = sumrow + 1
                Loop
            
            vol = vol + Cells(i, 7).Value
            close_price = Cells(i, 6).Value
            Cells(sumrow, 9).Value = ticker
            Cells(sumrow, 10).Value = Round(close_price - open_price, 4)
            Cells(sumrow, 11).Value = Round((close_price - open_price) / open_price, 4)
            Cells(sumrow, 12).Value = vol
            vol = 0
            sumrow = 1
            
        ElseIf Cells(i, 1).Value = Cells(i + 1, 1).Value And vol = 0 Then
            ticker = Cells(i, 1).Value
            open_price = Cells(i, 3).Value
            vol = vol + Cells(i, 7).Value
            close_price = Cells(i, 6).Value
            
        Else
            vol = vol + Cells(i, 7).Value
            close_price = Cells(i, 6).Value
            
        End If
    
    Next i
    
'------------------------------
'Populate second summary table
'------------------------------
    
    'Loop to set range of first summary table to populate second summary table
    sumrow = 1
    Do While Cells(sumrow, 9).Value <> ""
        sumrow = sumrow + 1
    Loop
        
    'Populate second summary table
    sumrow = sumrow - 1
    Cells(2, 17).Value = Round(WorksheetFunction.Max(Range("K:K")), 4)
    Cells(3, 17).Value = Round(WorksheetFunction.Min(Range("K:K")), 4)
    Cells(4, 17).Value = WorksheetFunction.Max(Range("L:L"))
    find = WorksheetFunction.Match(Cells(2, 17).Value, Range("K1:K" & sumrow), 0)
    Cells(2, 16).Value = Cells(find, 9).Value
    find = WorksheetFunction.Match(Cells(3, 17).Value, Range("K1:K" & sumrow), 0)
    Cells(3, 16).Value = Cells(find, 9).Value
    find = WorksheetFunction.Match(Cells(4, 17).Value, Range("L1:L" & sumrow), 0)
    Cells(4, 16).Value = Cells(find, 9).Value
    
'------------------------------
'Format spreadsheet
'------------------------------
    
    Range("K2:K" & sumrow).NumberFormat = "#,###0.00%"
    Range("L2:L" & sumrow).NumberFormat = "#,###0.00"
    Range("Q2").NumberFormat = "#,###0.00%"
    Range("Q3").NumberFormat = "#,###0.00%"
    Range("Q4").NumberFormat = "#,###0.00"
    Sheets(active_ws).Columns("A:Z").AutoFit
        
'----------------------------------------------
'Apply conditional formatting to Yearly Change
'----------------------------------------------
       
        Range("J2").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="=0"
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        Selection.FormatConditions(1).Interior.Color = RGB(0, 255, 0)
        Selection.FormatConditions(1).StopIfTrue = False
        Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        Selection.FormatConditions(1).Interior.Color = RGB(255, 0, 0)
        Selection.FormatConditions(1).StopIfTrue = False

Next active_ws

'----------------------------------------------
'Notify user macro completed
'----------------------------------------------

    'Activate first worksheet
    Worksheets(1).Activate
    Range("A1").Select

    'Track when macro ended
    stop_macro = Now
    
    'Calculate how long macro ran
    run_min = Int((stop_macro - start_macro) * 24 * 3600) / 60
    run_sec = Int((stop_macro - start_macro) * 24 * 3600) Mod 60
    
    'Notify user macro is complete
    MsgBox ("Macro is Complete!" & vbCrLf & "Macro ran for: " & run_min & " minutes " & run_sec & " seconds.")

End Sub
