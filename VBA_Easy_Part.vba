Sub A2015()
Dim current_ticker As String
Dim next_ticker As String
Dim Total_Stock_Volume As Double
Dim cc_row As Integer
cc_row = 2
Cells(1, 9).Value = "ticker"
Cells(1, 10).Value = "total stock volume"


LRow = Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To LRow
    current_ticker = Cells(i, 1).Value
    next_ticker = Cells(i + 1, 1).Value

    If current_ticker <> next_ticker Then
        Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
        Range("I" & cc_row).Value = current_ticker
        Range("J" & cc_row).Value = Total_Stock_Volume
        Total_Stock_Volume = 0
        cc_row = cc_row + 1
    Else
        Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value

    End If

Next i

End Sub
