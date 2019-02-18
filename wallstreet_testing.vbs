Sub wall_street():

'Declaring variables
Dim ticker As String
Dim total As Double
Dim store_total_row_number As Long
Dim i As Long

'set an initial value to 0 for holding the total
total = 0

'keep the track of the location for every tick
store_total_row_number = 2

'looping through the rows
With Worksheets("A")
' calculating last row
last_row = .Cells(.Rows.Count, "A").End(xlUp).Row
Set Rng = .Range("A2:A" & last_row)
End With

For i = 2 To last_row

    'check if we are still within the same tick
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

        'set the ticker name
        ticker = Cells(i, 1).Value

        'if column 1 is a new ticker, push total
        total = total + Cells(i, 7).Value

        'print tick name in a store_total_row_numberRange
        Range("I" & store_total_row_number).Value = ticker

        'print tick in a store_total_row_number
        Range("J" & store_total_row_number).Value = total

        'increment store total row number
        store_total_row_number = store_total_row_number + 1

        'reset total to 0
        total = 0

    Else
    total = total + Cells(i, 7).Value

    End If
Next i
End Sub
