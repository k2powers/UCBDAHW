Sub ChallengeHardStockInfo()
' Create a script that will loop through all the stocks and take the following info.

  ' Yearly change from what the stock opened the year at to what the closing price was.

  ' The percent change from the what it opened the year at to what it closed.

  ' The total Volume of the stock

  ' Ticker symbol

' You should also have conditional formatting that will highlight
' positive change in green and negative change in red.

' Your solution will also be able to locate the stock with the
' "Greatest % increase", "Greatest % Decrease" and "Greatest total volume"

'INSTRUCTIONS:
' Set up to run on all worksheets by counting the number of worksheets
' Declare Current as a worksheet object variable.
Dim Current As Worksheet
' Loop through all of the worksheets in the active workbook
For Each Current In Worksheets
Current.Activate
    
    'Define stock name variable as string
    Dim StockName As String
    
    ' Define the yearly change, % change, and total volume variables based on number type
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalVolume As Double
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    
    ' Set the original total volume variable to 0
    TotalVolume = 0
    
    ' add labels of Ticker, Yearly Change, Percent Change and Total Stock
    ' Volume to the top of two open columns to start your results table
    Range("J1").Value = "Ticker"
    Range("K1").Value = "Yearly Change"
    Range("L1").Value = "Percentage Change"
    Range("M1").Value = "Total Stock Volume"
    
    ' Define the row # in the results table where you want to put the
    ' name and value of the stock, and dim as integer
    Dim ResultsRow As Integer
    ResultsRow = 2
    
    ' define last row of the data, to use in For loop
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
        For j = 2 To LastRow
            
            ' set the open price for the first ticker the value in C2
            If j = 2 Then
            OpenPrice = Cells(2, 3).Value
            End If
        
        ' look at the stock name in each row, compare it to the one below (in col A)
        ' If it' s different, you' ve hit the last record for that stock, so define
        ' the stock name variable as that particular stock
            If Cells(j + 1, 1).Value <> Cells(j, 1).Value Then
            StockName = Cells(j, 1).Value
            
            ' Put that stock name in the next available cell in the results table
            Range("J" & ResultsRow).Value = StockName
            
            ' add the value in the last row for the stock to the current total volume
            TotalVolume = TotalVolume + Cells(j, 7).Value
            
            ' Put current value of the total volume for that ticker in results table
            Range("M" & ResultsRow).Value = TotalVolume
           
           ' reset the total volume variable to 0
            TotalVolume = 0
            
            ' set the close price as the LAST cell for that ticker in column F
            ClosePrice = Cells(j, 6).Value
            
            
            ' Put difference between close and open price in summary table column K
            Range("K" & ResultsRow).Value = ClosePrice - OpenPrice
            
            ' Put percentage change between close and open price in summary table column L
            ' If open price is 0, set % to 0, otherwise calculate
                If OpenPrice = 0 Then
                Range("L" & ResultsRow).Value = "0"
                Else
                Range("L" & ResultsRow).Value = (ClosePrice - OpenPrice) / OpenPrice
                End If
            
            ' Change style of percentage change to percent
            Range("L" & ResultsRow).Style = "Percent"
            
            ' format cells to show red for negative values and green for positive values
            If Range("L" & ResultsRow).Value < 0 Then
                Range("K" & ResultsRow & ":L" & ResultsRow).Interior.ColorIndex = 3
                ElseIf Range("L" & ResultsRow).Value > 0 Then
                Range("K" & ResultsRow & ":L" & ResultsRow).Interior.ColorIndex = 4
                Else: Range("K" & ResultsRow & ":L" & ResultsRow).Interior.ColorIndex = 1
            End If
            
            ' Add a row to the summary table
            ResultsRow = ResultsRow + 1
            
            ' set the next row' s open value
            OpenPrice = Cells(j + 1, 3).Value
            
            ' If it' s not different (same stock), add the value in col 3 to the total volume variable
            Else
            TotalVolume = TotalVolume + Cells(j, 7).Value
            End If
        
        ' move to the next row and repeat the process
        Next j
        
    ' Add Greatest table with row headers for "Greatest % increase"(P2),
    '"Greatest % Decrease"(P3)and "Greatest total volume" (P4)
        Range("P2") = "Greatest % Increase"
        Range("P3") = "Greatest % Decrease"
        Range("P4") = "Greatest Total Volume"
    
    ' Add column headers for "Ticker" (Q1) and "Value" (R1)
        Range("Q1") = "Ticker"
        Range("R1") = "Value"
    
    ' Create variables for the following items:
    
        ' Greatest % Increase Ticker Name
        Dim BigIncName As String
        ' Greatest % Increase Value
        Dim BigIncValue As Double
        ' Greatest % Decrease Ticker Name
        Dim BigDecName As String
        ' Greatest % Decrease Value
        Dim BigDecValue As Double
        ' Greatest Total Volume Name
        Dim TotalVolName As String
        ' Greatest Total Volume Value
        Dim TotalVolValue As Double
    
    ' Reuse LastRow variable for results table
    LastRow = Cells(Rows.Count, 10).End(xlUp).Row
    
    ' Find Greatest ...:
        'Loop through each row of Totals table, compare each value to the
        'one before, if it finds one larger, set variables to ticker name
        'and value, when it finds next larger, update variables
        'At last row of table, print values in Greatest table
        For j = 2 To LastRow
            
            ' set the variables for Greatest... to match the first ticker's values in results table,
            If j = 2 Then
                BigIncName = Cells(2, 10).Value
                BigIncValue = Cells(2, 12).Value
                BigDecName = Cells(2, 10).Value
                BigDecValue = Cells(2, 12).Value
                TotalVolName = Cells(2, 10).Value
                TotalVolValue = Cells(2, 13).Value
            End If
            
            ' look at the stock name in each row, compare it to the one below (in col J. If diff, that's last record for that stock, set ticker variable
            If Cells(j + 1, 12).Value > BigIncValue Then
                BigIncName = Cells(j + 1, 10).Value
                BigIncValue = Cells(j + 1, 12).Value
            ElseIf Cells(j + 1, 12).Value < BigDecValue Then
                BigDecName = Cells(j + 1, 10).Value
                BigDecValue = Cells(j + 1, 12).Value
            End If
            
            If Cells(j + 1, 13).Value > TotalVolValue Then
                TotalVolName = Cells(j + 1, 10).Value
                TotalVolValue = Cells(j + 1, 13).Value
            End If
        
        Next j
       
        ' Put that stock info in the greatest increase row in the Greatest table
        Range("Q2").Value = BigIncName
        Range("R2").Value = BigIncValue
        Range("Q3").Value = BigDecName
        Range("R3").Value = BigDecValue
        Range("Q4").Value = TotalVolName
        Range("R4").Value = TotalVolValue
        
        ' Change format style
        Range("R2:R3").Style = "Percent"
    
    'Repeat for next worksheet through last worksheet
Next Current

End Sub



