Attribute VB_Name = "Module1"
Sub Stock_stats()


Dim ws As Worksheet

For Each ws In Worksheets
    
output_row = 2
LastRow = ActiveSheet.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    
    For row_ = 2 To LastRow
        
'   Find unique tickers and place in output table
        If ws.Cells(row_ + 1, 1).Value <> ws.Cells(row_, 1).Value Then
            ws.Cells(output_row, 9).Value = ws.Cells(row_, 1).Value
            output_row = output_row + 1
            
'           Calculate the Total Volume amount
                volume_ = volume_ + ws.Cells(row_, 7).Value
                ws.Cells(output_row - 1, 12).Value = volume_
                   volume_ = 0
                    
'           Locate first row of opening cost and last closing cost of the year
                open_row = ws.Range("A:A").Find(what:=ws.Cells(row_, 1).Value, _
                    after:=ws.Range("A1")).Row
                close_row = ws.Range("A:A").Find(what:=ws.Cells(row_, 1).Value, _
                    after:=ws.Range("A1"), SearchDirection:=xlPrevious).Row
                    
'           Calculate Yearly Change amount
                y_change = ws.Cells(close_row, 6).Value - ws.Cells(open_row, 3).Value
                ws.Cells(output_row - 1, 10).Value = y_change
                    y_change = 0
                   
'           Calculate the Percent Change amount
                change_ = (ws.Cells(close_row, 6).Value - ws.Cells(open_row, 3).Value) / _
                    ws.Cells(open_row, 3)
                ws.Cells(output_row - 1, 11).Value = change_
                
                    change_ = 0
                
'           Format Percent Change column with colors
                y_change = ws.Cells(output_row - 1, 10).Value
                If y_change > 0 Then
                    ws.Cells(output_row - 1, 10).Interior.Color = vbGreen
                    ElseIf y_change < 0 Then
                        ws.Cells(output_row - 1, 10).Interior.Color = vbRed
                        Percent = ws.Range("K:K").Value
                
                increase = WorksheetFunction.Max(Percent)
                    max_row = WorksheetFunction.Match(increase, Percent, 0)
                    ws.Range("O2").Value = ws.Cells(max_row, 9).Value
                    ws.Range("P2").Value = increase
                
                decrease = WorksheetFunction.Min(Percent)
                   min_row = WorksheetFunction.Match(decrease, Percent, 0)
                   ws.Range("O3").Value = ws.Cells(min_row, 9).Value
                    ws.Range("P3").Value = decrease
                   
                ws.Range("P2:P3").NumberFormat = "0.00%"
                
              t_volume = ws.Range("L:L").Value
            
                 high_volume = WorksheetFunction.Max(t_volume)
                    vol_row = WorksheetFunction.Match(high_volume, t_volume, 0)
                        ws.Range("O4").Value = ws.Cells(vol_row, 9).Value
                    ws.Range("P4").Value = high_volume
                ws.Range("P4").NumberFormat = "General"
               End If
  
            
            Else
            volume_ = volume_ + ws.Cells(row_, 7).Value
           
        End If
           
    Next row_

Next ws

End Sub




