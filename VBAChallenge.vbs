Sub Multiple_year_stock_data():

    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    For Each ws In Worksheets
    
   
    Dim rownumber As Long
    Dim tick_counter As Long
    'Last row column A
    Dim last_row_a As Long
    Dim percent_change As Double

'Add the text on columns for Data Analisys and Table Summary --> Source: VBA Activity 07

        ws.Range("i1") = "Ticker"
        ws.Range("j1") = "Yearly Change"
        ws.Range("k1") = "Percent Change"
        ws.Range("l1") = "Total Stock Volume"
        ws.Range("p1") = "Ticker"
        ws.Range("q1") = "Value"
        ws.Range("o2") = "Greatest & Increase"
        ws.Range("o3") = "Greatest & Decrease"
        ws.Range("o4") = "Greatest & Total Volume"
 
 'Create a counter for <ticker> for first row"
        tick_counter = 2
 
 'Set the row number 2
        rownumber = 2
    
 'Determine last row --> Source: VBA activity 07 & https://www.wallstreetmojo.com/vba-last-row/
        last_row_a = ws.Cells(Rows.Count, 1).End(xlUp).Row

            'Loop through all rows
            For i = 2 To last_row_a
 
                'Check if we are still within the same ticker, if it is not...
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'Add value in column i9
                ws.Cells(tick_counter, 9).Value = ws.Cells(i, 1).Value
                
                'Calculate and place Yearly Change in column j10
                'Example of first cell --> Yearly change = F2 (last change value) - c2 (first value)
                ws.Cells(tick_counter, 10).Value = ws.Cells(i, 6).Value - ws.Cells(rownumber, 3).Value
                
                
                    'Conditional formating
                    If ws.Cells(tick_counter, 10).Value < 0 Then
                
                    'Set cell background color to red
                    ws.Cells(tick_counter, 10).Interior.ColorIndex = 3
                
                    Else
                
                    'Set cell background color to green
                    ws.Cells(tick_counter, 10).Interior.ColorIndex = 4
                
                    End If
                
                    'Calculate and place percent change in column k11
                    'Example of first cell --> percent_chance = F2(last change value) - c2 (fisrt value) / c2 (fisrt value)
                    If ws.Cells(rownumber, 3).Value <> 0 Then
                    percent_change = ((ws.Cells(i, 6).Value - ws.Cells(rownumber, 3).Value) / ws.Cells(rownumber, 3).Value)
                
                    'Percent formating --> Source: https://www.techonthenet.com/excel/formulas/format_string.php
                    'percencen format = percent_change * 100
                    ws.Cells(tick_counter, 11).Value = Format(percent_change, "Percent")
                  
                    Else
                    ws.Cells(tick_counter, 11).Value = Format(0, "Percent")
                
                    End If
                    
                    'Calculate and place total volume in column l12 --> Source: https://excelchamps.com/vba/sum/
                    ws.Cells(tick_counter, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(rownumber, 7), ws.Cells(i, 7)))
                    
                
            'Increase ticker_counter by 1
            tick_counter = tick_counter + 1
                
            'Set new start row of the ticker block
            rownumber = i + 1
                
            End If
    Next i
    
 Next ws
 
 End Sub
