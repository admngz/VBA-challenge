# VBA-challenge




Sub yearly_stock_report():
'variables
    Dim ticker As String
    Dim ticker_symbol As Integer
    Dim lastrow As Long
    Dim annual_open As Double
    Dim annual_close As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim total_stock As Double
    
'loop'n thru ws - yay for loops
    For Each ws In Worksheets
        ws.Activate
    
'titles
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

'values
    lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row
   
   'tickerscript
   'ticker symbol = 2
   'For i = 2 To lastrow
   'If cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        'ticker = Cells(i, 1).Value
    'Range("H" & ticker_symbol).Value = ticker
    'ticker_symbol = ticker_symbol + 1
    'End if
    'Next i
    
    For i = 2 To lastrow
        ticker = Cells(i, 1).Value
    'conditional for open
        If annual_open = 0 Then
            annual_open = Cells(i, 3).Value
        End If
    'total_stock - define
        total_stock = total_stock + Cells(i, 7).Value
    ' <> finding differences thru tickers
        If Cells(i + 1, 1).Value <> ticker Then
            ticker_symbol = ticker_symbol + 1
            Cells(ticker_symbol + 1, 9) = ticker
        'close- define
            annual_close = Cells(i, 6)
        'yearly change equation
            yearly_change = annual_close - annual_open
        'location
            Cells(ticker_symbol + 1, 10).Value = yearly_change
           'green highlight
            If yearly_change > 0 Then
                Cells(ticker_symbol + 1, 10).Interior.ColorIndex = 4
            ' red highlight
            ElseIf yearly_change < 0 Then
                Cells(ticker_symbol + 1, 10).Interior.ColorIndex = 3
            ' yellow highlight
            Else
                Cells(ticker_symbol + 1, 10).Interior.ColorIndex = 6
            End If
        
            ' conditional for percent change
            If annual_open = 0 Then
                percent_change = 0
            Else
                percent_change = (yearly_change / annual_open)
            End If
            
            
            'NOTE: look up how to change to %????
            Cells(ticker_symbol + 1, 11).Value = Format(percent_change, "Percent")
            annual_open = 0
            
            
            
            ' total stock
            Cells(ticker_symbol + 1, 12).Value = total_stock
            total_stock = 0
        End If
    Next i
    Next ws
    End Sub
