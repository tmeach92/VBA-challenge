Sub Stock_Script():
    
For Each ws In Worksheets

    ws.Activate

    'assign variables
    Dim ticker As String
    Dim ticker_value As Integer
    ticker_value = 0
    Dim open_value As Double
    open_value = 0
    Dim close_value As Double
    close_value = 0
    Dim yearly_change As Double
    yearly_change = 0
    Dim percent_change As Double
    percent_change = 0
    Dim total_stock_volume_value As Double
    total_stock_volume_value = 0
    
    'Determine last row
    lastRowState = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    'create table headers
        Cells(1, 10).Value = "Ticker"
        Cells(1, 11).Value = "Yearly Change"
        Cells(1, 12).Value = "Percent Change"
        Cells(1, 13).Value = "Total Stock Volume"
        
    'loop through all tickers
    For i = 2 To lastRowState
    
        'ticker symbol
        ticker = Cells(i, 1).Value
        
        'opening value
        If open_value = 0 Then
            open_value = Cells(i, 3).Value
        End If
        
        'total stock volume for each ticker
        total_stock_volume = total_stock_volume + Cells(i, 7).Value
        
        'see if still in the same ticker
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            ticker_value = ticker_value + 1
            Cells(ticker_value + 1, 10).Value = ticker
        
            'end of the year close value
            close_value = Cells(i, 6)
        
            'yearly change
            yearly_change = close_value - open_value
            Cells(ticker_value + 1, 11).Value = yearly_change
        
            ' Yearly change > 0 green, < 0 red and = 0 yellow
            If yearly_change > 0 Then
                Cells(ticker_value + 1, 11).Interior.ColorIndex = 4
            ElseIf yearly_change < 0 Then
                Cells(ticker_value + 1, 11).Interior.ColorIndex = 3
            Else
                Cells(ticker_value + 1, 11).Interior.ColorIndex = 6
            End If
            
            'percent change
            percent_change = (yearly_change / open_value)
            Cells(ticker_value + 1, 12).Value = Format(percent_change, "percent")
            
            'total stock volume
            Cells(ticker_value + 1, 13).Value = total_stock_volume
            
            'reset open value and total stock volume for each ticker
            open_value = 0
            total_stock_volume = 0
            
            
        End If
        
    Next i
    
Next ws

End Sub
