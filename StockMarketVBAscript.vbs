Sub Week2Challenge()

    '
    Dim row As Double
    Dim lrow As Double
    Dim summary_table_row As Double
    Dim ticker_name As String
    Dim yearly_open As Double
    Dim yearly_close As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim total_volume As Double
    Dim summary_lrow As Double
    
    ' Bonus
    Dim max_percent As Double
    Dim max_percent_ticker As String
    Dim min_percent As Double
    Dim min_percent_ticker As String
    Dim max_volume As Double
    Dim max_volume_ticker As String
    
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Sheets
    
        ' Set variable lrow to the number of rows in the sheet
        lrow = ws.Cells(Rows.Count, 1).End(xlUp).row
        
        ' Initialize the next available row in the first summary table to 2
        summary_table_row = 2
        
        ' Initialize the running total for total_volume for each ticker symbol to 0
        total_volume = 0
        
        '  Create headers for summary table and format as bold/centered
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("I1:L1").Font.Bold = True
        ws.Range("I1:L1").HorizontalAlignment = xlCenter
        
        ' For loop to iterate through all rows in the worksheet
        For row = 2 To lrow
            
            ' Checks if current cell is first of its kind to set yearly_open
            If ws.Cells(row, 1).Value <> ws.Cells(row - 1, 1).Value Then
            
            ' Sets yearly_open equal to the opening price
                yearly_open = ws.Cells(row, "C").Value
                
            End If
            
            'Checks next row to see if there are more of the same ticker symbol
            If ws.Cells(row, 1).Value = ws.Cells(row + 1, 1).Value Then
            
                 'Update running volume total
                total_volume = total_volume + ws.Cells(row, "G").Value
            
            ' Takes effect if ticker in current cell is the last of its kind
            Else
            
                ' Adds ticker symbol to next empty row in summary table Ticker column
                ws.Cells(summary_table_row, "I").Value = ws.Cells(row, 1).Value
                
                'Update running volume total
                total_volume = total_volume + ws.Cells(row, "G").Value
    
                '  Sets yearly_close equal to closing price
                yearly_close = ws.Cells(row, "F").Value
                
                'Calculate yearly change
                yearly_change = (yearly_close - yearly_open)
    
                'Update Yearly Change column in summary table
                ws.Cells(summary_table_row, "J").Value = yearly_change
                
                'Update Percent Change column in summary table
                ws.Cells(summary_table_row, "K").Value = yearly_change / yearly_open
                
                'Update Total Volume column in summary table
                ws.Cells(summary_table_row, "L").Value = total_volume
                
                ' Reset yearly_change, yearly_open, and total_volume for next ticker
                yearly_change = 0
                yearly_open = 0
                total_volume = 0
                
                ' Updates next empty cell in summary table
                summary_table_row = summary_table_row + 1
                
            End If
                
        Next row
        
        ' Assign variable for last row number in summary table
        summary_lrow = ws.Cells(Rows.Count, "I").End(xlUp).row
        
        'Format Percent Change column as percents
        ws.Range("K2:K" & summary_lrow).NumberFormat = "0.00%"
        
        'Conditional formatting for Percent Change Column
            For row = 2 To summary_lrow
            
                ' Check if the value is positive
                If ws.Cells(row, "J").Value > 0 Then
                
                    '  Color interior of positive percentages green
                    ws.Cells(row, "J").Interior.ColorIndex = 4
                    
                ' Check if value is negative
                ElseIf ws.Cells(row, "J").Value < 0 Then
                
                    '  Color interior of negative percentages red
                    ws.Cells(row, "J").Interior.ColorIndex = 3
                    
                End If
                
            Next row
        
        ' BONUS
        'Set Summary Table 2 for Greatest %Increase/%Decrease/Total Volume and "Ticker"/"Value" Headers
        ws.Cells(2, "O").Value = "Greatest % Increase"
        ws.Cells(3, "O").Value = "Greatest % Decrease"
        ws.Cells(4, "O").Value = "Greatest Total Volume"
        ws.Cells(1, "P").Value = "Ticker"
        ws.Cells(1, "Q").Value = "Value"
        
        ' Initialize min and max values to 0
        min_percent = 0
        max_percent = 0
        max_volume = 0
        
        
        ' For loop to iterate through the rows of the summary table
        For row = 2 To summary_lrow
    
            
            '  Check if each cell in the Percent Change column is GREATER THAN the current MAX percent change
            If ws.Cells(row, "K").Value > max_percent Then
            
                ' Update max_percent to the new greatest percent change
                max_percent = ws.Cells(row, "K").Value
                
                ' Update max_percent_ticker to the ticker symbol of the new max value
                max_percent_ticker = ws.Cells(row, "I").Value
                
            ' Check if each cell in the Percent Change column is LESS THAN the current MIN percent change
            ElseIf ws.Cells(row, "K").Value < min_percent Then
            
                '  Update min_percent to the new lowest percent change
                min_percent = ws.Cells(row, "K").Value
                
                '  Update min_percent_ticker to the ticker symbol of the new min value
                min_percent_ticker = ws.Cells(row, "I").Value
                
            End If
            
            ' Check if each cell in the Total Stock Volume column is greater than the current MAX volume
            If ws.Cells(row, "L").Value > max_volume Then
            
                ' Update max_volume to the new greatest total volume amount
                max_volume = ws.Cells(row, "L").Value
                
                ' Update max_volume_ticker to the ticker symbol of the new max value
                max_volume_ticker = ws.Cells(row, "I").Value
                
            End If
            
        Next row
        
        '  Assign all ticker symbols and values to the new summary table
        ws.Cells(2, "P").Value = max_percent_ticker
        ws.Cells(2, "Q").Value = max_percent
        ws.Cells(3, "P").Value = min_percent_ticker
        ws.Cells(3, "Q").Value = min_percent
        ws.Cells(4, "P").Value = max_volume_ticker
        ws.Cells(4, "Q").Value = max_volume
        
        
        
        
        ' 'Format Percent Change column as percents
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        
        '  Auto-fit all used columns
        ws.UsedRange.EntireColumn.AutoFit
        
    Next ws
    
End Sub