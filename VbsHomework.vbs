Sub stock()
'Looping through all the worksheets
    For Each ws In Worksheets
    
'assigning data types to variable
        Dim ticker_symbol As String
        Dim total_stock As LongLong
        Dim MyRange As Range
'total_stock variable
        total_stock = 0
        Dim summary_table_row As Integer
               
'summary table row variable
        summary_table_row = 2
        
 'creating headers throughout our worksheets
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
'This line returns last row number in lastRow Variable.
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
'starting variable for stick value
        Start = 2
        
        
        
'looping throug all the stocks
            For i = 2 To last_row
             If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                
                 'assigning ticker symbol, total stock, yearly change, and percent change
                 ticker_symbol = ws.Cells(i, 1).Value
                 total_stock = total_stock + ws.Cells(i, 7).Value
                 closing_value = ws.Cells(i, 6).Value
                 opening_value = ws.Cells(Start, 3).Value
                 yearly_change = closing_value - opening_value
                
                 
                 If opening_value = 0 Then
                    Range("K" & summary_table_row).Value = Null
                 Else
                    Range("K" & summary_table_row).Value = FormatPercent((closing_value / opening_value) - 1)
                 End If
                'printing ticker symbol, total stock, yearly change, and percent change in assigned columns
                 ws.Range("I" & summary_table_row).Value = ticker_symbol
                 ws.Range("J" & summary_table_row).Value = yearly_change
                 ws.Range("L" & summary_table_row).Value = total_stock
                 
                'conditional formating
                 Set MyRange = ws.Range("J" & summary_table_row)
                    MyRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
                    MyRange.FormatConditions(1).Interior.ColorIndex = 4
                    MyRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
                    MyRange.FormatConditions(2).Interior.ColorIndex = 3
                    
                 Set MyRange = ws.Range("K" & summary_table_row)
                    MyRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
                    MyRange.FormatConditions(1).Interior.ColorIndex = 4
                    MyRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
                    MyRange.FormatConditions(2).Interior.ColorIndex = 3
                    
                
                 
        'increase summary table row
                 summary_table_row = summary_table_row + 1
                 
        'reset ticker volume
                 total_stock = 0
                 
                 
        'increase start
                 Start = i + 1
                 
                 
             Else
                 total_stock = total_stock + Cells(i, 7).Value
             End If
             
             
    
         Next i
    Next ws
End Sub
