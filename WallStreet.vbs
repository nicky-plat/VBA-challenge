Sub WallStreetTest():

'Declare variables
    Dim current_row As Long
    Dim last_row As Long
    Dim summary_row As Long
    Dim total_volume As Variant
    Dim open_price As Double
    Dim close_price As Double
    Dim yearly_change As Double
    Dim percent_change As Long
    Dim ticker_data As String
    Dim starting_row As Double
    
    
    
For Each ws In Worksheets
    
    'Initialize variables
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    summary_row = 2
    total_volume = 0
    yearly_change = 0
    percent_change = 0
    
    'Create Headers in Summary Table
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ws.Columns("I:L").Font.Bold = True
    ws.Columns("I:L").EntireColumn.AutoFit
        
    
    'Iterate through the worksheet from row 2 to last row
    For current_row = 2 To last_row
    
        If ticker_data <> ws.Range("A" & current_row).Value And ws.Range("C" & current_row).Value <> 0 Then
        
            ticker_data = ws.Range("A" & current_row).Value
            
            open_price = ws.Range("C" & current_row).Value
            
            
        End If
    
    
        If ws.Cells(current_row + 1, 1) <> ws.Cells(current_row, 1) Then
            
            
            'If ws.Range("C2").Value <= 0 Then
            
            total_volume = total_volume + ws.Cells(current_row, 7).Value
            ws.Range("L" & summary_row).Value = total_volume
            
            close_price = ws.Range("F" & current_row).Value
            
            'Calculate Yearly Change
            yearly_change = close_price - open_price
            
                'Check if opening price is 0
                If open_price = 0 Then
                    percent_change = 0
                    ws.Range("K" & summary_row).Value = percent_change
                Else
                
                    'Calculate Percent Change
                    percent_change = yearly_change / open_price * 100
                    
                    
                    'Print ticker to summary table
                    ws.Range("I" & summary_row).Value = ws.Cells(current_row, 1).Value
                    
                    'Print yearly change to summary table
                    ws.Range("J" & summary_row).Value = yearly_change
                    
                    'Print percent change to summary table
                    ws.Range("K" & summary_row).Value = percent_change
                End If
            
            'Conditional formatting for ticker
            If ws.Range("J" & summary_row).Value < 0 Then
                ws.Range("J" & summary_row).Interior.ColorIndex = 3
                
            ElseIf ws.Range("J" & summary_row).Value > 0 Then
                ws.Range("J" & summary_row).Interior.ColorIndex = 4
                
            End If
            
            
            'reset ticker variables for next ticker
            open_price = 0
            total_volume = 0
            summary_row = summary_row + 1
            
            
        Else
        
            'add the volume for that day to my total volume
            total_volume = total_volume + ws.Cells(current_row, 7).Value
            
            
        End If
        
    
    
    Next current_row
    
    
Next ws

End Sub
