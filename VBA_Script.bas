Attribute VB_Name = "Module1"
Sub alphabetical_testing()

    'define variable data type
   Dim last_row As Long
   Dim open_price As Double
   Dim yearly_change As Double
   Dim pct_change As Double
   Dim total_volume As LongLong
   Dim summary_row As Long
   Dim column As Integer
   
   For Each ws In ActiveWorkbook.Worksheets
   
    
   ' initialize variable values
   summary_row = 2
   total_volume = 0
   open_price = ws.Cells(2, 3).Value
  
   
   
    
   'display column labels
   ws.Cells(1, 9).Value = "Ticker"
   ws.Cells(1, 10).Value = "Yearly Change"
   ws.Cells(1, 11).Value = "Percent Change"
   ws.Cells(1, 12).Value = "Total Volume"
   
   ws.Cells(1, 16).Value = "Ticker"
   ws.Cells(1, 17).Value = "Value"
   ws.Cells(2, 15).Value = "Greatest % Increase"
   ws.Cells(3, 15).Value = "Greatest % Decrease"
   ws.Cells(4, 15).Value = "Greatest Total Volume"
   
   
   'For understanding last_row = Last Row

   last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'NESTED LOOPS:
    'Loop through column 1
    For i = 2 To last_row
    
    
        If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
        
            total_volume = total_volume + Cells(i, 7).Value
        
        Else
        
            total_volume = total_volume + ws.Cells(i, 7).Value
            
            yearly_change = ws.Cells(i, 6).Value - open_price
            
            pct_change = yearly_change / open_price
            
            'display the results
            ws.Cells(summary_row, 9).Value = ws.Cells(i, 1).Value
            ws.Cells(summary_row, 10).Value = yearly_change
            ws.Cells(summary_row, 11).Value = pct_change
            ws.Cells(summary_row, 11).NumberFormat = "0.00%"
            ws.Cells(summary_row, 12).Value = total_volume
            
            
            If yearly_change > 0 Then
            ws.Cells(summary_row, 10).Interior.ColorIndex = 4
            Else
            ws.Cells(summary_row, 10).Interior.ColorIndex = 3
            End If
            
            
            're-initialize
            open_price = CDbl(ws.Cells(i + 1, 3).Value)
            summary_row = summary_row + 1
          
        
        End If
        
    Next i
    
   last_summary_row = ws.Cells(Rows.Count, "I").End(xlUp).Row
   
   
   greatest_total_volume = -1E+55
   
   volume_ticker = ""
   
   
   greatest_percent_increase = -1E+57
    
    greatest_increase_ticker = ""
    
    
    greatest_percent_decrease = 1E+61
    
    greatest_decrease_ticker = ""
    
    
   'For loop
   
   For j = 2 To last_summary_row
   
   
   current_volume = ws.Cells(j, "L").Value
   
   current_percentage = ws.Cells(j, "K").Value
   
   
   If current_volume > greatest_total_volume Then
   
   greatest_total_volume = current_volume
   
   volume_ticker = ws.Cells(j, "I").Value
   
   
   End If
   
   
   If current_percentage > greatest_percent_increase Then
   
   greatest_percent_increase = current_percentage
   
   greatest_increase_ticker = ws.Cells(j, "I").Value
   
      
   End If
   
   
   If current_percentage < greatest_percent_decrease Then
   
   greatest_percent_decrease = current_percentage
   
   greatest_decrease_ticker = ws.Cells(j, "I").Value
   
      
   End If
   
   
   
   Next j
       
    ws.Cells(4, "p").Value = volume_ticker
    ws.Cells(4, "q").Value = greatest_total_volume
    
    ws.Cells(2, "p").Value = greatest_increase_ticker
    ws.Cells(2, "q").Value = greatest_percent_increase
    
    ws.Cells(3, "p").Value = greatest_decrease_ticker
    ws.Cells(3, "q").Value = greatest_percent_decrease
    
    
    
      

 
Next


End Sub




