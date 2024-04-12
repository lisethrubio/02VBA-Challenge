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
   
    
   ' initialize variable values
   summary_row = 2
   total_volume = 0
   open_price = Cells(2, 3).Value
  
   
   
    
   'display column labels
   Cells(1, 9).Value = "Ticker"
   Cells(1, 10).Value = "Yearly Change"
   Cells(1, 11).Value = "Percent Change"
   Cells(1, 12).Value = "Total Volume"
   
   Cells(1, 16).Value = "Ticker"
   Cells(1, 17).Value = "Value"
   Cells(2, 15).Value = "Greatest % Increase"
   Cells(3, 15).Value = "Greatest % Decrease"
   Cells(4, 15).Value = "Greatest Total Volume"
   
   
   'For understanding last_row = Last Row

   last_row = Cells(Rows.Count, 1).End(xlUp).Row

    'NESTED LOOPS:
    'Loop through column 1
    For i = 2 To last_row
    
    
        If Cells(i, 1).Value = Cells(i + 1, 1).Value Then
        
            total_volume = total_volume + Cells(i, 7).Value
        
        Else
        
            total_volume = total_volume + Cells(i, 7).Value
            
            yearly_change = Cells(i, 6).Value - open_price
            
            pct_change = yearly_change / open_price
            
            'display the results
            Cells(summary_row, 9).Value = Cells(i, 1).Value
            Cells(summary_row, 10).Value = yearly_change
            Cells(summary_row, 11).Value = pct_change
            Cells(summary_row, 12).Value = total_volume
            
            
            're-initialize
            open_price = CDbl(Cells(i + 1, 3).Value)
            summary_row = summary_row + 1
          
        
        End If
        
    Next i
    
    MsgBox (summary_row)
    
    For j = 2 To summary_row
     
    Next j
    
    For n = 2 To 90
    
        If Cells(i, column).Value > 0 Then
    
        Cells(i, column).Interior.ColorIndex = 3
        
        Else
        
        Cells(i, column).Interior.ColorIndex = 5
        
        
        End If
    
    Next cell
    
    



End Sub



'1. loop through worksheets (at the beginning)
'2. Loop through all the rows and place conditionals (within the loop)
'3. formatting (last step)

