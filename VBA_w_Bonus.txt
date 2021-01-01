Sub stock_filter()

    'Loop Through All Worksheets

For Each ws In Worksheets

    'Set a string variable for Ticker

Dim Ticker As String

    'set a double variable for YearlyChange

Dim YearlyChange As Double

    'set a double variable for PercentChange

Dim PercentChange As Double

    'Set a variant variable for TotalStockVolume (Long did not work for some reason)

Dim TotalStockVolume As Variant

    'Set an initial starting row for summary information which will be looped as the ticker changes while being looped. call it SummaryRow

TotalStockVolume = 0
    
    'start the summary row on the second row because we have headers in the first row
    
Dim SummaryRow As Integer

SummaryRow = 2

'Add headers for the summary table

ws.Range("I1").Value = "Ticker"

ws.Range("J1").Value = "Yearly Change"

ws.Range("K1").Value = "Percent Change"

ws.Range("L1").Value = "Stock Volume"

'Assign a value to YearOpen to grab the first instance of the opening price which will be reassigned further in the loop

YearOpen = ws.Cells(2, 3).Value

        'create a loop (i) that loops through to the end of the last row on the sheet

    For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row

                'create a conditional statement that determines IF the next Ticker is different. If so, THEN do the following

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then

                'Determine current Ticker name

            Ticker = ws.Cells(i, 1).Value

                'Assign a columnn value to Ticker in the respective SummaryRow
            
            ws.Range("I" & SummaryRow).Value = Ticker

                'determine the YearlyChange value by subtracting the closing price of the Ticker's year from the opening price of that year. (closing price - opening price)
                
            YearClose = ws.Cells(i, 6).Value
            
            YearChange = YearClose - YearOpen

                'Assign a column value to YearlyChange in the respective SummaryRow

            ws.Range("J" & SummaryRow).Value = YearChange

                'Determine the PercentChange by dividing YearlyChange by the initial opening price. We can not divide by 0 so we need to create a short conditional statement to avoid 0.

                    If YearOpen <> 0 Then

                    PercentChange = (YearClose - YearOpen) / YearOpen

                    End If

                'Assign a column value to PercentChange in the respective SummaryRow

            ws.Range("K" & SummaryRow).Value = PercentChange

                'Add to the current TotalStockVolume
            
            TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value

                'Assign a columnn value to TotalStockVolume in the respective SummaryRow

            ws.Range("L" & SummaryRow).Value = TotalStockVolume
            
            'Reassign value to YearOpen
            
            YearOpen = ws.Cells(i + 1, 3).Value

            'Add 1 to the SummaryRow

            SummaryRow = SummaryRow + 1

            'Reset TotalStockVolume

            TotalStockVolume = 0

            'Else, add to the TotalStockVolume

        Else

            TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value

        End If

    Next i

'format YearlyChange column with +/-, green/red cell indicator

For i = 2 To ws.Cells(Rows.Count, 10).End(xlUp).Row
    
    If ws.Cells(i, 10).Value >= 0 Then

    ws.Cells(i, 10).Interior.Color = RGB(0, 255, 0)

    Else

    ws.Cells(i, 10).Interior.Color = RGB(255, 0, 0)

    End If

Next i
    
    'Create a table indicating the max/min percent change and the stock with the highest total Volume

    ' we need 3 variables, Max, Min, and GreatestVolume. each set equal to 0

Max = 0

Min = 0

GreatestVolume = 0

        'create a loop for the summary table column starting and 2 and counting from the bottom up.

    For i = 2 To ws.Cells(Rows.Count, 11).End(xlUp).Row

            'add a conditional statement checking if the cell in the loop is greater than Max (or 0).

        If ws.Cells(i, 11) > Max Then

            'if it is greater than Max, then Max is equal to that cell.

        Max = ws.Cells(i, 11).Value

            'assign a cell value equal to Max to store the new Max value.

        ws.Range("P2").Value = Max

            'grab the ticker by assigning the proportionate (i) value to the column where Tickers live and assign it a new cell in the summary table
        
        ws.Range("O2").Value = ws.Cells(i, 9)
        
        End If

            'new if statement for minimum. this will look just like the maximum but this time the cell value is less than Min.

        If ws.Cells(i, 11).Value < Min Then

            'set Min equal to the cell value within the loop

        Min = ws.Cells(i, 11).Value

            'give Min a cell to live in
        
        ws.Range("P3").Value = Min
        
            'grab the ticker

        ws.Range("O3").Value = ws.Cells(i, 9)

        End If

            'new conditional statement for GreatestVolume, just like the Max conditional statement

        If ws.Cells(i, 12).Value > GreatestVolume Then
        
            'I'm starting to see a patern.

        GreatestVolume = ws.Cells(i, 12).Value

            'give GreatestColumn a place to live in the summary table

        ws.Range("P4").Value = GreatestVolume
        
            'assign the ticker value

        ws.Range("O4").Value = ws.Cells(i, 9)

        End If

Next i
    
    'format the following ranges to give percentage symbols

ws.Range("P2:P3").NumberFormat = "0.00%"

ws.Range("K:K").NumberFormat = "0.00%"

    'assign written values for the summary table.

ws.Range("N2").Value = "Greatest % Increase"

ws.Range("N3").Value = "Greatest % Decrease"

ws.Range("N4").Value = "Greatest Total Volume"

ws.Range("O1").Value = "Ticker"

ws.Range("P1").Value = "Value"

    'on to the next worksheet

Next ws

End Sub