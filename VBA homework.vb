Sub ticker_function()
    
    'define everything from ws (given)
    Dim my_worksheet As Worksheet
    Dim ticker  As String
    Dim my_date As String
    Dim open_amount As Double
    Dim high As Double
    Dim low As Double
    Dim close_amount As Double
    Dim daily_vol As Long
    
    'define everything in ws (new header)
    Dim yearly_change As Double
    Dim year_close As Double
    Dim year_open As Double
    Dim percent_change As Double
    Dim output_row As Integer
    
    
'going through each worksheet
For Each my_worksheet In Worksheets
   
    my_worksheet.Cells(1, 9).Value = "Ticker"
    my_worksheet.Cells(1, 10).Value = "Yearly change"
    my_worksheet.Cells(1, 11).Value = "Percent change"
    my_worksheet.Cells(1, 12).Value = "Volume total"
    
    my_worksheet.Columns("K").NumberFormat = "0.00%"
    output_row = 2
    LastRow = my_worksheet.Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To LastRow
    ticker = my_worksheet.Cells(i, 1).Value
    daily_vol = my_worksheet.Cells(i, 7).Value
    
    volume_total = volume_total + daily_vol
        'find last unique row and read last value of year close
        If my_worksheet.Cells(i + 1, 1).Value <> my_worksheet.Cells(i, 1).Value Then
        
            year_close = my_worksheet.Cells(i, 6).Value
        
            'calculate change, % and total
            yearly_change = year_close - year_open
            
            'in case values are 0s, just assign 0 and move on
            If year_close = 0 Then
                
                percent_change = 0
            
            Else
            
                percent_change = (year_close - year_open) / year_close
            
            End If
            
            my_worksheet.Cells(output_row, 9).Value = ticker
            my_worksheet.Cells(output_row, 10).Value = yearly_change
            my_worksheet.Cells(output_row, 11).Value = percent_change
            my_worksheet.Cells(output_row, 12).Value = volume_total
                       
            If yearly_change > 0 Then
                my_worksheet.Cells(output_row, 10).Interior.Color = vbGreen
            Else
                my_worksheet.Cells(output_row, 10).Interior.Color = vbRed
            End If
            
            output_row = output_row + 1
            
        End If
        
        
        'find first unique row and read first value of year open
        If my_worksheet.Cells(i - 1, 1).Value <> my_worksheet.Cells(i, 1).Value Then
            
            'this resets volume when it get to a new ticker
            volume_total = daily_vol
            year_open = my_worksheet.Cells(i, 3).Value
        
        End If
        
        
        
        
    Next i

Next
    
End Sub
