Attribute VB_Name = "Module1"
Sub stocks()

Dim sheet As Worksheet

For Each sheet In ThisWorkbook.Worksheets
sheet.Activate

'Insert column headers
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Volume Traded"

'Determine how many rows are in first column and store variable
Dim rw As Long
Range("A1").Select

rw = Range("A1").End(xlDown).Row

'Declare variables for storing information
Dim row_num, volume_traded As Integer
Dim close_initial, close_final, close_diff, close_perc As Double
Dim ticker_name As String
Dim ticker_summary, yearly_change, total_volume, percentage_diff As Integer

'Set intial counters
ticker_summary = 2
yearly_change = 2
percentage_diff = 2
total_volume = 2

For i = 2 To rw
'Check if row above has different value in column 1
If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
row_num = i
ticker_name = Cells(i, 1).Value
close_initial = Cells(i, 3).Value
Cells(ticker_summary, 9).Value = ticker_name
End If

'Check if row below is different to current row
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
close_final = Cells(i, 6).Value
close_diff = close_final - close_initial
close_perc = close_diff / close_initial

'Check if difference in stock open and close value is negative or positive then insert it & fill cell red or green
If close_diff < 0 Then
Cells(yearly_change, 10).Interior.ColorIndex = 3
Cells(yearly_change, 10).Value = close_diff
Else
Cells(yearly_change, 10).Interior.ColorIndex = 4
Cells(yearly_change, 10).Value = close_diff
End If

'Insert percentage change into appropriate column
Cells(percentage_diff, 11).Value = close_perc

'Calculate total stock volume from the first appearance of the stock to current cell (last mention before new stock is encountered)
Cells(total_volume, 12) = Application.Sum(Range(Cells(row_num, 7), Cells(i, 7)))

'Begin a new row when a new stock in encountered
ticker_summary = ticker_summary + 1
yearly_change = yearly_change + 1
percentage_diff = percentage_diff + 1
total_volume = total_volume + 1

End If

Next i

'Change Percent Change Column format
With ActiveSheet
.Range("K2:K" & rw).NumberFormat = "0.00%"
End With

'Bonus
'Finding the Greatest Percentage Increase & Decrease and Greatest Total Volume Traded
'Insert Column Headers and Row Descriptions
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"

'Find the Values and Insert into appropriate Cells
Cells(2, 17).Value = WorksheetFunction.Max(Columns("K"))
Cells(3, 17).Value = WorksheetFunction.Min(Columns("K"))
Cells(4, 17).Value = WorksheetFunction.Max(Columns("L"))

'Get Ticker of values above
'Find the number of values in Percent Change Column
Dim tr As Integer
Range("K1").Select

tr = Range("K1").End(xlDown).Row

For i = 2 To tr
'Check for match between values column and percent change column
If Cells(i, 11).Value = Cells(2, 17).Value Then
Cells(2, 16).Value = Cells(i, 9).Value
ElseIf Cells(i, 11).Value = Cells(3, 17).Value Then
Cells(3, 16).Value = Cells(i, 9).Value
End If

Next i

For i = 2 To tr
'Check for match between values column and volume traded column
If Cells(i, 12).Value = Cells(4, 17).Value Then
Cells(4, 16).Value = Cells(i, 9).Value
End If

Next i

'Change Value Column Number Format to Percent for selected range
With ActiveSheet
.Range("Q2:Q3").NumberFormat = "0.00%"
End With

'AutoFit Column Widths of the Used Columns
ActiveSheet.UsedRange.EntireColumn.AutoFit
ActiveSheet.UsedRange.EntireRow.AutoFit

Next sheet
End Sub


