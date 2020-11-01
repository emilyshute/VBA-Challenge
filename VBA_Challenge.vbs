Sub VBA_Stocks()

'Set variables
Dim ws As Worksheet
Dim ticker As String
Dim sum_table_row As Integer
Dim LastRow As Long

'Loop through all existing worksheets
For Each ws In Worksheets

'Label Headers
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

'Set values to variables
sum_table_row = 2

'Loop through tickers
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To LastRow
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

'Insert ticker ID
    ticker = ws.Cells(i, 1).Value
    ws.Range("I" & sum_table_row).Value = ticker

'Reset Variable
    sum_table_row = sum_table_row + 1

End If
Next i
Next ws
End Sub

Sub VBA_volume()

'Set variables
Dim ws As Worksheet
Dim ticker As String
Dim sum_table_row As Integer
Dim LastRow As Long
Dim total_stock_vol As Double

'Loop through all existing worksheets
For Each ws In Worksheets

'Set values to variables
sum_table_row = 2
total_stock_vol = 0

'Loop through tickers
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To LastRow
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
'Begin column calculations
    total_stock_vol = total_stock_vol + ws.Cells(i, 7).Value
    ws.Range("L" & sum_table_row).Value = total_stock_vol

'Reset Variables
    sum_table_row = sum_table_row + 1
    total_stock_vol = 0
    
Else
total_stock_vol = total_stock_vol + ws.Cells(i, 7).Value

End If
Next i
Next ws
End Sub

Sub VBA_yearly()

'Set variables
Dim ws As Worksheet
Dim ticker As String
Dim sum_table_row As Integer
Dim LastRow As Long
Dim yearly_change As Double
Dim perc_change As Double
Dim year_open As Double
Dim year_close As Double

'Loop through all existing worksheets
For Each ws In Worksheets

sum_table_row = 2
yearly_change = 0
perc_change = 0
year_open = 0
year_close = 0

'Loop through tickers
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To LastRow
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

'Calculations
    year_open = year_open + ws.Cells(i, 3).Value
    year_close = year_close + ws.Cells(i, 6).Value
    yearly_change = year_close - year_open
    perc_change = (yearly_change / year_open)
    perc_change = perc_change * 100

    ws.Range("J" & sum_table_row).Value = yearly_change
    ws.Range("K" & sum_table_row).Value = perc_change

'Reset Variables
    sum_table_row = sum_table_row + 1
    yearly_change = 0
    perc_change = 0
    year_open = 0
    year_close = 0

Else
year_open = year_open + ws.Cells(i, 3).Value
year_close = year_close + ws.Cells(i, 6).Value

End If
Next i
Next ws
End Sub

Sub VBA_Bonus()

'Set variables
Dim ws As Worksheet
Dim greatest_inc As Double
Dim greatest_dec As Double
Dim greatest_total_vol As Double

'Loop through all existing worksheets
For Each ws In Worksheets

'Begin loops for calculations
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To LastRow
    
    greatest_inc = Application.WorksheetFunction.Max(Range("K:K"))
    ws.Cells(2, 17).Value = greatest_inc

    greatest_dec = Application.WorksheetFunction.Min(Range("K:K"))
    ws.Cells(3, 17).Value = greatest_dec

    greatest_total_vol = Application.WorksheetFunction.Max(Range("L:L"))
    ws.Cells(4, 17).Value = greatest_total_vol

Next i
Next ws
End Sub

Sub formatting()

'Set variables
Dim ws As Worksheet
Dim LastRowSum As Integer

'Loop through all existing worksheets
For Each ws In Worksheets

'Label Headers
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"

'Adjust format to include percent sign
ws.Range("K:K").NumberFormat = "0.00%"
ws.Cells(2, 17).NumberFormat = "0.00%"
ws.Cells(3, 17).NumberFormat = "0.00%"

'Expand columns
Columns("A:Q").EntireColumn.AutoFit

LastRowSum = ws.Cells(Rows.Count, 10).End(xlUp).Row
For i = 2 To LastRowSum

'set colours
    If ws.Cells(i, 10).Value < 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 3
    
    ElseIf ws.Cells(i, 10).Value > 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 4

End If
Next i
Next ws
End Sub
