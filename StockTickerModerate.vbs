Sub StockTicker()

    Dim ticker_symbol as String
    
    Dim open_price as Double
    open_price = 0

    Dim close_price as Double
    close_price = 0
    
    Dim change as Double
    change = 0

    dim change_percent as Double
    change_percent = 0

    Dim last_row as LongLong

    last_row = Cells(Rows.Count, 1).End(xlUp).Row

    ticker_symbol = Cells(2,1).Value
    open_price = Cells(2,3).Value
    close_price = cells(2,6).Value
    trade_date = Cells(2,2).Value    


    For Each ws in Worksheets

        For i = 2 to last_row
            if ws.cells(i+1, 1).value <> ws.cells(i,1).value then
                ticker_row = ticker_row + 1
                ticker_symbol = ws.cells(i,1).value
                ws.cells(ticker_row, "I").value = ticker_symbol
            
                close_price = ws.Cells(i,6).value
                change = close_price - open_price
            
            elseif open_price <> 0 then
                change = (change / open_price) * 100


            end if

            i = i + 1
        next j

    Next ws

End Sub
