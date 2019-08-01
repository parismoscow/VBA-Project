Attribute VB_Name = "Module2"

Sub Hard_Moderate_Easy_Stock()
    ' Loop thru all sheets
    ' --------------------------------------------
Dim WS As Worksheet
    For Each WS In Worksheets
    WS.Activate
        ' Determine the Last Row
        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row

        ' Add Heading
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        'Create Variable
        Dim Open_Price As Double
        Dim Close_Price As Double
        Dim Yearly_Change As Double
        Dim Ticker_Name As String
        Dim Percent_Change As Double
        Dim Total_Stock_Volume As Double
        Total_Stock_Volume = 0
        Dim cc_row As Double
        Dim greatest_total_Volume As Double
        cc_row = 2


        'Set Initial Open Price
        Open_Price = Cells(2, 3).Value
         ' Loop through all ticker symbol

        For i = 2 To LastRow
         ' Check same ticker symbol
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                ' Set Ticker name
                Ticker_Name = Cells(i, 1).Value
                Cells(cc_row, 9).Value = Ticker_Name
                ' Set Close Price
                Close_Price = Cells(i, 6).Value
                ' Add Yearly Change
                Yearly_Change = Close_Price - Open_Price
                Cells(cc_row, 10).Value = Yearly_Change
                ' Add Percent Change
                If (Open_Price = 0 And Close_Price = 0) Then
                    Percent_Change = 0
                ElseIf (Open_Price = 0 And Close_Price <> 0) Then
                    Percent_Change = 1
                Else
                    Percent_Change = Yearly_Change / Open_Price
                    Cells(cc_row, 11).Value = Percent_Change
                    Cells(cc_row, 11).NumberFormat = "0.00%"
                End If
                ' Add Total_Stock_Volume
                Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
                Cells(cc_row, 12).Value = Total_Stock_Volume
                ' Add one to the summary table row
                cc_row = cc_row + 1
                ' reset the Open Price
                Open_Price = Cells(i + 1, 3)
                ' reset the Total_Stock_Volume
                Total_Stock_Volume = 0
            'if cells are the same ticker
            Else
                Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
            End If
        Next i

        ' Determine the Last Row of Yearly Change per WS
        YearlyChangeLastRow = WS.Cells(Rows.Count, 9).End(xlUp).Row
        ' Set the Cell Colors
        For j = 2 To YearlyChangeLastRow
            If (Cells(j, 10).Value > 0 Or Cells(j, 10).Value = 0) Then
                Cells(j, 10).Interior.ColorIndex = 10
            ElseIf Cells(j, 10).Value < 0 Then
                Cells(j, 10).Interior.ColorIndex = 3
            End If
        Next j

        'set greatest% increase, greatest % decrease, greatest total Volume
        Cells(2, 15).Value = "greatest % increase"
        Cells(3, 15).Value = "greatest % decrease"
        Cells(4, 15).Value = "greatest total Volume"
        Cells(1, 16).Value = "ticker"
        Cells(1, 17).Value = "value"

        'loop thru each rows to get greatest value and its ticker
        For x = 2 To YearlyChangeLastRow
            If Cells(x, 11).Value = WorksheetFunction.Max(WS.Range("K2:K" & YearlyChangeLastRow)) Then
                Cells(2, 16).Value = Cells(x, 9).Value
                Cells(2, 17).Value = Cells(x, 11).Value
                Cells(2, 17).NumberFormat = "0.00%"

            ElseIf Cells(x, 11).Value = WorksheetFunction.Min(WS.Range("K2:K" & YearlyChangeLastRow)) Then
                Cells(3, 16).Value = Cells(x, 9).Value
                Cells(3, 17).Value = Cells(x, 11).Value
                Cells(3, 17).NumberFormat = "0.00%"
            ElseIf Cells(x, 12).Value = WorksheetFunction.Max(WS.Range("L2:L" & YearlyChangeLastRow)) Then
                Cells(4, 16).Value = Cells(x, 9).Value
                Cells(4, 17).Value = Cells(x, 12).Value
            End If
        Next x


    Next WS

End Sub
