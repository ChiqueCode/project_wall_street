Sub Wall_Street():
For Each ws In Worksheets
    'declare all your variables
    Dim ticker As String
    Dim yearly_change As Double
    Dim percentage_change As Double
    Dim total_stock_value As Double
    Dim storage_rows As Integer
    Dim open_price As Double
    Dim close_price As Double
    Dim last_row As Double
    Dim i As Double

'print column header for ticker
ws.Cells(1, 9).Value = "Ticker"
'print column header for yearly change
ws.Cells(1, 10).Value = "Yearly Change"
'print column header for percentage change
ws.Cells(1, 11).Value = "Percentage Change"
'print column header for total stock volume
ws.Cells(1, 12).Value = "Total Stock Value"
ws.Cells(1, 18).Value = "Close Price"
ws.Cells(1, 19).Value = "Open Price"

'For Each ws In Worksheets

    'save default values for all static variables
    storage_rows = 2
    'save last row of current worksheet
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    'grab the WorkSheetName for some reason
    WorkSheetName = ws.Name
    'set total stock volume to 0
    total_stock_value = 0
    open_price = ws.Cells(2, 3).Value

    'For loop (2nd row to last row)
    For i = 2 To last_row

        'if (current row equals 2)
        'If ws.Cells(i, 1).Value = ws.Cells(2, 1).Value Then
            
            'save & print name
            'ticker = ws.Cells(i, 1).Value
            'save total open from current row
            
            'TODO save total volume of stock
            'total_stock_value = total_stock_value + ws.Cells(i, 7).Value

        'else if (current row name does not equal the next row name)
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            'set the ticker name
            ticker = ws.cells(i, 1).Value 
            'save total close from current row
            close_price = ws.Cells(i, 6).Value
            'save the yearly change of total close - open
            yearly_change = close_price - open_price
            'save & print final total volume of stock
            total_stock_value = total_stock_value + ws.Cells(i, 7).Value
            'print ticker name in a store
            ws.Range("I" & storage_rows).Value = ticker
            'print ticker in a total volume 
            'printing out values
            ws.Range("J" & storage_rows).Value = yearly_change
            ws.Range("K"& storage_rows).Value = percentage_change    
            ws.Range("L" & storage_rows).Value = total_stock_value
            'print ticker close_price 
            ws.Range("R" & storage_rows).Value = close_price
            'print open_price
            ws.Range("S" & storage_rows).Value = open_price
            'increment store value 
            storage_rows = storage_rows + 1
            'save total volume of stock as 0
            total_stock_value = 0
            'save total close = 0
            open_price = ws.Cells(i + 1, 3).Value
            'print percentage difference between total close and total open
            'save & print name of next row
            'save total open from next row (current row +1)
            
        Else
            'save total volume of stock
            total_stock_value = total_stock_value + ws.Cells(i, 7).Value

    End If
    Next i

    'save next worksheet name

'End worksheet for each loop
Next ws
End Sub
