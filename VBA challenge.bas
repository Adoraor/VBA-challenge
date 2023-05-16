Attribute VB_Name = "Module1"
Sub assignment2()

For Each ws In Worksheets
        summary_row = 2
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
  
        Lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        Dim open_price As Double

        For i = 2 To Lastrow
            If ws.Cells(i + 1, 1) <> ws.Cells(i, 1).Value Then
                ticker_name = ws.Cells(i, 1).Value
                yearly_change = ws.Cells(i, 6).Value - open_price
                percent_change = yearly_change / open_price
                stock_volume = stock_volume + ws.Cells(i, 7).Value

                ws.Range("I" & summary_row).Value = ticker_name
                ws.Range("J" & summary_row).Value = yearly_change
                ws.Range("J" & summary_row).NumberFormat = "0.00"
                ws.Range("K" & summary_row).Value = percent_change
                ws.Range("K" & summary_row).NumberFormat = "0.00%"
                ws.Range("L" & summary_row).Value = stock_volume

                If yearly_change > 0 Then
                    ws.Range("J" & summary_row).Interior.Color = vbGreen
                Else
                    ws.Range("J" & summary_row).Interior.Color = vbRed
                End If

                summary_row = summary_row + 1
                yearly_change = 0
                percent_change = 0
                stock_volume = 0
                open_price = 0
            Else
                If open_price = 0 Then
                    open_price = ws.Cells(i, 3).Value
                End If

                stock_volume = stock_volume + ws.Cells(i, 7).Value
            
End If
        
Next i
  
'Part 2

ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
 ws.Cells(2, 15).Value = "Greatest % increase"
 ws.Cells(3, 15).Value = "Greatest % decrease"
 ws.Cells(4, 15).Value = "Greatest total volume"

summary_lastrow = ws.Cells(ws.Rows.Count, 9).End(xlUp).Row

Dim max_volume As Double
max_volume = 0

For i = 2 To summary_lastrow

If ws.Cells(i, 11).Value = WorksheetFunction.Max(ws.Range("K2:K" & summary_lastrow).Value) Then
ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
ws.Cells(2, 17).NumberFormat = "0.00%"

ElseIf ws.Cells(i, 11).Value = WorksheetFunction.Min(ws.Range("K2:K" & summary_lastrow).Value) Then
ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
ws.Cells(3, 17).NumberFormat = "0.00%"

ElseIf ws.Cells(i, 12).Value > max_volume Then
    max_volume = ws.Cells(i, 12).Value
    ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
    ws.Cells(4, 17).Value = max_volume


End If

Next i

Next ws

End Sub
