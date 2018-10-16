Sub ModerateStockInfo()
' Create a script that will loop through all the stocks and take the following info.

  ' Yearly change from what the stock opened the year at to what the closing price was.

  ' The percent change from the what it opened the year at to what it closed.

  ' The total Volume of the stock

  ' Ticker symbol

' You should also have conditional formatting that will highlight positive change in green and negative change in red.

'Define each stock name as a string variable
Dim StockName As String
'Define the yearly change, % change, and total volume variables based on number type
Dim YearlyChange As Double
Dim PercentChange As Double
Dim TotalVolume As Double
Dim OpenPrice As Double
Dim ClosePrice As Double
'Set the original total volume variable to 0
TotalVolume = 0
'add labels of Ticker, Yearly Change, Percent Change and Total Stock Volume to the top of two open columns to start your results table
Range("J1").Value = "Ticker"
Range("K1").Value = "Yearly Change"
Range("L1").Value = "Percentage Change"
Range("M1").Value = "Total Stock Volume"
'Define the row # in the results table where you want to put the name and value of the stock, and dim as integer
Dim ResultsRow As Integer
ResultsRow = 2
'define last row of the data, to use in For loop
LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    For j = 2 To LastRow
        'set the open price for the first ticker the value in C2
        If j = 2 Then
        OpenPrice = Cells(2, 3).Value
        End If
    'look at the stock name each row, compare it to the one below (in col A)
    'If it's different, you've hit the last record for that stock, so define the stock name variable as that particular stock
        If Cells(j + 1, 1).Value <> Cells(j, 1).Value Then
        StockName = Cells(j, 1).Value
        'Put that stock name in the next available cell in the results table
        Range("J" & ResultsRow).Value = StockName
        'add the value in the last row for the stock to the current total volume
        TotalVolume = TotalVolume + Cells(j, 7).Value
        'Put current value of the total volume for that ticker in results table
        Range("M" & ResultsRow).Value = TotalVolume
       'reset the total volume variable to 0
        TotalVolume = 0
        'set the close price as the LAST cell for that ticker in column F
        ClosePrice = Cells(j, 6).Value
        'Put difference between close and open price in summary table column K
        Range("K" & ResultsRow).Value = ClosePrice - OpenPrice
        'Put percentage change between close and open price in summary table column L
        'If open price is 0, set % to 0, otherwise calculate
            If OpenPrice = 0 Then
            Range("L" & ResultsRow).Value = "0"
            Else
            Range("L" & ResultsRow).Value = (ClosePrice - OpenPrice) / OpenPrice
            End If
        'Change style of percentage change to percent
        Range("L" & ResultsRow).Style = "Percent"
        'format cells to show red for negative values and green for positive values
        If Range("L" & ResultsRow).Value < 0 Then
            Range("K" & ResultsRow & ":L" & ResultsRow).Font.ColorIndex = 3
            ElseIf Range("L" & ResultsRow).Value > 0 Then
            Range("K" & ResultsRow & ":L" & ResultsRow).Font.ColorIndex = 4
            Else: Range("K" & ResultsRow & ":L" & ResultsRow).Font.ColorIndex = 1
        End If
        'Add a row to the summary table
        ResultsRow = ResultsRow + 1
        'set the next row's open value
        OpenPrice = Cells(j + 1, 3).Value
    'If it's not different (same stock), add the value in col 3 to the total volume variable
        Else
        TotalVolume = TotalVolume + Cells(j, 7).Value
        End If
    'move to the next row and repeat the process
    Next j
End Sub