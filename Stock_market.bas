Attribute VB_Name = "Módulo1"
Sub Stock_market()



For Each ws In Worksheets


tickerrow = 1
first_open = ws.Cells(2, 3).Value
v1 = ws.Cells(2, 7).Row



   
    

    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To lastRow
           
            If ws.Cells(i, 1) <> ws.Cells(i + 1, 1) Then
            
              'Ticker name
              ws.Cells(tickerrow + 1, 9).Value = ws.Cells(i, 1).Value
              
              'Yearly value
                            
              ws.Cells(tickerrow + 1, 10).Value = ws.Cells(i, 6) - first_open
              If ws.Cells(tickerrow + 1, 10).Value > 0 Then
                ws.Cells(tickerrow + 1, 10).Interior.ColorIndex = 4
                Else
                ws.Cells(tickerrow + 1, 10).Interior.ColorIndex = 3
              End If
                              
              
              'Percent Change
              If first_open = 0 Then
                ws.Cells(tickerrow + 1, 11).Value = 0
              Else
                ws.Cells(tickerrow + 1, 11).Value = 1 - (ws.Cells(i, 6) / first_open)
                
              End If
                
              'Total Volume
              v2 = ws.Cells(i, 7)
              ws.Cells(tickerrow + 1, 12).Formula = "=SUM(G" & v1 & ": G" & i & ")"
                
              
              tickerrow = tickerrow + 1
              first_open = ws.Cells(i + 1, 3)
              ws.Cells(tickerrow, 11).Style = "Percent"
              v1 = ws.Cells(i + 1, 7).Row
                           
              
              
              
            End If
        Next i
        
    Next ws
        
    


End Sub

