Sub code_testing()

Dim ws As Worksheet

For Each ws In Worksheets

'Defining my variables
Dim ticker As String
Dim vol, year_open, year_close, total_stock_vol As Integer
Dim yearly_change, percent_change, greatest_increase, greatest_decrease, greatest_volume As Double
Dim lastrow, ticketrrow, lastrowticker As Long

'setting my ticker row counter
tickerrow = 2
j = 2

'setting headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Total"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
'loop
    For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
    
'to ouput each kind of ticker
    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
        ws.Cells(tickerrow, "I").Value = ws.Cells(i, 1).Value
        
'yearly change
        ws.Cells(tickerrow, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                
                If ws.Cells(tickerrow, "J").Value < 0 Then
                'setting to green
                ws.Cells(tickerrow, "J").Interior.ColorIndex = 3
Else
'setting to red
    ws.Cells(tickerrow, "J").Interior.ColorIndex = 4
    
End If
'percent change
    If ws.Cells(j, 3).Value <> 0 Then
    percent_change = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
    ws.Cells(tickerrow, 11).Value = percent_change
    ws.Cells(tickerrow, 11).NumberFormat = "0.00%"

Else
'formating into percentage
    ws.Range("Q2").NumberFormat = "0.00%"
    
End If
    ws.Cells(tickerrow, 12).Value = WorksheetFunction.Sum(ws.Range(ws.Cells(j, 7), ws.Cells(i, 7)))
    tickerrow = tickerrow + 1
    j = i + 1

End If

Next i

'data needed for functionality table
        greatest_volume = ws.Cells(2, 12).Value
        greatest_increase = ws.Cells(2, 11).Value
        greatest_decrease = ws.Cells(2, 11).Value
        
'to find the last row
        lastrowticker = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        ws.Range("Q2") = WorksheetFunction.Max(ws.Range("K2:K" & lastrowticker))
        ws.Range("Q2").NumberFormat = "0.00%"
        
        ws.Range("Q3") = WorksheetFunction.Min(ws.Range("K2:K" & lastrowticker))
        ws.Range("Q3").NumberFormat = "0.00%"
        
        ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & lastrowticker))
        
        increased_index = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & lastrowticker)), ws.Range("K2:K" & lastrowticker), 0)
        ws.Cells(2, 16).Value = ws.Cells(increased_index + 1, 9).Value
'loop
    For i = 2 To ws.Cells(Rows.Count, 9).End(xlUp).Row
    
'to find the greatest vol
    If ws.Cells(i, 12).Value > greatest_volume Then
                greatest_volume = ws.Cells(i, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
    End If
            
'to find the greatest decrease
    If ws.Cells(i, 11).Value < Greatest_Decreased Then
                Greatest_Decreased = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value

                
     End If
            
    
    Next i
    
    
    Next ws
    
End Sub