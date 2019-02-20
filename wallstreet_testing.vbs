Sub wall_street():
'Loop through all sheets 

For Each ws in Worksheets

'Declaring variables
Dim WorksheetName as String
Dim ticker As String
Dim total As Double
Dim store_total_row_number As Long
Dim i As Long

'Name Ticker & Total Stock Volume
ws.Cells(1, 9).Value = "Ticker" 
ws.Cells(1, 10).Value = "Total Stock Volume"

'Determine the last Row
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Grab the WorkSheetName 
WorksheetName = ws.name

'set an initial value to 0 for holding the total
total = 0

'keep the track of the location for every tick
store_total_row_number = 2

For i = 2 To LastRow

    'check if we are still within the same tick
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

        'set the ticker name
        ticker = ws.Cells(i, 1).Value

        'if column 1 is a new ticker, push total
        total = total + ws.Cells(i, 7).Value

        'print tick name in a store_total_row_numberRange
        ws.Range("I" & store_total_row_number).Value = ticker

        'print tick in a store_total_row_number
        ws.Range("J" & store_total_row_number).Value = total

        'increment store total row number
        store_total_row_number = store_total_row_number + 1

        'reset total to 0
        total = 0

    Else
    total = total + ws.Cells(i, 7).Value

    End If
Next i
Next ws
End Sub
