Sub Moderate_Stock()
Dim WS As Worksheet
    For Each WS In Worksheets
    WS.Activate
        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        Dim Open_Price As Double
        Dim Close_Price As Double
        Dim Yearly_Change As Double
        Dim Ticker_Name As String
        Dim Percent_Change As Double
        Dim Total_Stock_Volume As Double
        Total_Stock_Volume = 0
        Dim current_ticker As String
        Dim next_ticker As String
        cc_row = 2
        Open_Price = Cells(2, 3).Value


        For i = 2 To LastRow
            current_ticker = Cells(i, 1).Value
            next_ticker = Cells(i + 1, 1).Value
            If current_ticker <> next_ticker Then
                Ticker_Name = Cells(i, 1).Value
                Cells(cc_row, 9).Value = Ticker_Name
                Close_Price = Cells(i, 6).Value
                Yearly_Change = Close_Price - Open_Price
                Cells(cc_row, 10).Value = Yearly_Change
                If (Open_Price = 0 And Close_Price = 0) Then
                    Percent_Change = 0
                ElseIf (Open_Price = 0 And Close_Price <> 0) Then
                    Percent_Change = 1
                Else
                    Percent_Change = Yearly_Change / Open_Price
                    Cells(cc_row, 11).Value = Percent_Change
                    Cells(cc_row, 11).NumberFormat = "0.00%"
                End If
                Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
                Cells(2, 12).Value = Total_Stock_Volume
                cc_row = cc_row + 1
                Open_Price = Cells(i + 1, 3)
                Total_Stock_Volume = 0
            Else
                Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
            End If
        Next i
        Yearly_Change_LastRow = WS.Cells(Rows.Count, 10).End(xlUp).Row
        For j = 2 To Yearly_Change_LastRow
            If (Cells(j, 10).Value > 0 Or Cells(j, 10).Value = 0) Then
                Cells(j, 10).Interior.ColorIndex = 10
            ElseIf Cells(j, 10).Value < 0 Then
                Cells(j, 10).Interior.ColorIndex = 3
            End If
        Next j

    Next WS


End Sub
