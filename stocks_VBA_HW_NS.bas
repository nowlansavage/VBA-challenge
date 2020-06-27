Attribute VB_Name = "Module1"
Sub stocks()
'define variables
Dim stock_volume As Double
Dim year_start As Double
Dim year_end As Double
Dim percent_change As Double
Dim greatest_percent_change As Double
Dim lowest_percent_change As Double
Dim greatest_stock_volume As Double
Dim ticker_symbol As String


'Set up summary row to keep track where stock summary should be entered
Dim summary_row As Integer


'Loop through all sheets
For Each ws In Worksheets

    'set labels for summary
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"


    'set labels for greatest
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"
    
    'Set the first item as the starting price for each sheet
    year_start = ws.Range("F2").Value
    
    'In each worksheet set summary row to 2
    summary_row = 2
    
    'In each worksheet set all greatest marks to zero
    greatest_percent_change = 0
    lowest_percent_change = 0
    greatest_stock_volume = 0

    'Loop through all stock data
    For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row

        'Check is next cell is in same stock ticker. If not then..
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            'set summary row
            summary_row = (ws.Cells(Rows.Count, 9).End(xlUp).Row + 1)
            
            'set year_end value
            year_end = ws.Cells(i, 6).Value
            
            'calculate and display yearly change
            ws.Range("J" & summary_row).Value = year_end - year_start
            
            'set color of cell green for growth and red for loss
            If ws.Range("J" & summary_row).Value > 0 Then
                ws.Range("J" & summary_row).Interior.ColorIndex = 4
            ElseIf ws.Range("J" & summary_row).Value < 0 Then
                ws.Range("J" & summary_row).Interior.ColorIndex = 3
            Else
            End If
            
            'calculate and display percent yearly change and check for zeros in denominator
            If year_start = 0 Then
                ws.Range("K" & summary_row).Value = "N/A"
            Else
                percent_change = ((year_end - year_start) / year_start)
                ws.Range("K" & summary_row).Value = percent_change
            End If
            
            'check for biggest increase percent change
            If percent_change > greatest_percent_change Then
                greatest_percent_change = percent_change
                ws.Range("O2").Value = ws.Cells(i, 1).Value
                ws.Range("P2").Value = greatest_percent_change
                'change format to percent
                ws.Range("P2").NumberFormat = "0.00%"
            Else
            End If
            
            'check for biggest decrease percent change
            If percent_change < lowest_percent_change Then
                lowest_percent_change = percent_change
                ws.Range("O3").Value = ws.Cells(i, 1).Value
                ws.Range("P3").Value = lowest_percent_change
                'change format to percent
                ws.Range("P3").NumberFormat = "0.00%"
            Else
            End If
            
            'change format to percent
            ws.Range("K" & summary_row).NumberFormat = "0.00%"
            
            'set ticker symbol
            ticker_symbol = Cells(i, 1).Value
            ws.Range("I" & summary_row).Value = ticker_symbol
       
            'calculate and set stock volume
            stock_volume = stock_volume + ws.Cells(i, 7)
            ws.Range("L" & summary_row).Value = stock_volume
            
            'check for biggest stock volume
            If stock_volume > greatest_stock_volume Then
                greatest_stock_volume = stock_volume
                ws.Range("O4").Value = ws.Cells(i, 1).Value
                ws.Range("P4").Value = greatest_stock_volume
            Else
            End If
       
            'Set stock volume back to zero for next stock
            stock_volume = 0
           
            
            'Set cell i+1 to year start
            year_start = ws.Cells(i + 1, 6).Value
            
        'If it is the same then
        Else
            'update stock volume
            stock_volume = stock_volume + ws.Cells(i, 7)
        
        End If
    Next i
Next ws
End Sub
