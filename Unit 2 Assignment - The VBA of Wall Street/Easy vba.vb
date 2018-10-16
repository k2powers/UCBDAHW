Sub EasyStockInfo()
'Create a script that will loop through each year of stock data and grab the total amount of volume each stock had over the year.
'You will also need to display the ticker symbol to coincide with the total volume.

'Define each stock name as a string variable
Dim StockName As String
'Define the total volume variable based on number type
Dim TotalVolume As Double
'Set the original total volume variable to 0
TotalVolume = 0
'add labels of Ticker and Total Stock Volume to the top of two open columns to start your results table
Range("J1").Value = "Ticker"
Range("K1").Value = "Total Stock Volume"
'Define the row # in the results table where you want to put the name and value of the stock, and dim as integer
Dim ResultsRow As Integer
ResultsRow = 2
'define last row of the data, to use in For loop
LastRow = Cells(Rows.Count, 1).End(xlUp).Row
'look at the stock name each row, compare it to the one below (in col A)
    For j = 2 To LastRow
    'If it's different, you've hit the last record for that stock, so define the stock name variable as that particular stock
        If Cells(j + 1, 1).Value <> Cells(j, 1).Value Then
        StockName = Cells(j, 1).Value
        'Put that stock name in the next available cell in the results table
        Range("J" & ResultsRow).Value = StockName
        'add the value in the last row for the stock to the current total volume
        TotalVolume = TotalVolume + Cells(j, 7).Value
        'set the value for the cell in the results table in the row # to be equal to the total volume to print it there
        Range("K" & ResultsRow).Value = TotalVolume
       'Add a row to the summary table
        ResultsRow = ResultsRow + 1
       'reset the total volume variable to 0
        TotalVolume = 0
    'If it's not different (same stock), add the value in col 3 to the total volume variable
        Else
        TotalVolume = TotalVolume + Cells(j, 7).Value
        End If
    'move to the next row and repeat the process
    Next j
End Sub
