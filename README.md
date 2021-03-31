# VBA-challenge
 Sub Stock_Data():

    ' Set initial variable for worksheet
 Dim ws As Worksheet

    'Activate worksheets so that we can run program through each worksheet
 Set ws = ActiveSheet

    'Use For command to ensure process takes place in each worksheet in workbook
  For Each ws In Worksheets

    'Create column labels for the variables
  ws.Cells(1, 9).Value = "Ticker"
  ws.Cells(1, 10).Value = "Yearly Change"
  ws.Cells(1, 11).Value = "Percent Change"
  ws.Cells(1, 12).Value = "Total Stock Volume"
  ws.Cells(2, 15).Value = "Greatest % Increase"
  ws.Cells(3, 15).Value = "Greatest % Decrease"
  ws.Cells(4, 15).Value = "Greatest Total Volume"
  ws.Cells(1, 16).Value = "Ticker"
  ws.Cells(1, 17).Value = "Value"

    'Set initial variables for ticker selections, the changes from open to close, percent changes, open and close
  Dim ticker As String
  Dim yearly_change As Double
  Dim percent_change As Double
  Dim year_opn As Double
  Dim year_cls As Double
  
    'Set initial variables for greatest percent increase, decrease and total along with their ticker names
  Dim greatest_percent_increase As Double
  greatest_percent_increase = 0
  Dim greatest_percent_decrease As Double
  greatest_percent_decrease = 0
  Dim greatest_total As LongLong
  greatest_total = 0
  Dim gpi_ticker As String
  Dim gpd_ticker As String
  Dim gt_ticker As String

    'Set initial variables for volume total and start at 0 for adding totals between different tickers
  Dim vol_total As LongLong
  vol_total = 0

    'Set initial variables and starting point for tracking ticker in loops
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

    'Set initial variables for using last row
  Dim LastRow As Long

    'Define last row in order to count and locate the last row in a column
  LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'Define where to grab year open
  year_opn = ws.Cells(2, 3).Value

    'Resize the columns to properly fit the text
  ws.Columns("A:Q").AutoFit
 
     'Check the entire ticker row
   For i = 2 To LastRow
     
       'If we are at the end of the first ticker row, so end of year, then proceed below
   If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
             
        'Define where to grab ticker and year close amount
   ticker = ws.Cells(i, 1).Value
   year_cls = ws.Cells(i, 6).Value
            
          'Define how to calculate yearly change
   yearly_change = year_cls - year_opn
        
           'Define how to calculate volume total (add volume to volume total)
   vol_total = vol_total + ws.Cells(i, 7).Value
        
            'Define the location to place ticker, yearly change, and volume total
   ws.Range("I" & Summary_Table_Row).Value = ticker
        
   ws.Range("J" & Summary_Table_Row).Value = yearly_change
          
   ws.Range("L" & Summary_Table_Row).Value = vol_total
        
             'Check that if yearly change is not equal to 0 change, then proceed with calculating percent change
   If yearly_change <> 0 And year_opn <> 0 Then
   percent_change = yearly_change / year_opn
                    
              'Cannot divide by 0, so if we have a change from 0 it is 100 percent change
   ElseIf yearly_change <> 0 And year_opn = 0 Then
   percent_change = 100
                   
            ' If no change, it is 0 percent change
   Else
   percent_change = 0
        
   End If
        
         'Define the location to put percent change, and turn it into a number format of %
  ws.Range("K" & Summary_Table_Row).Value = percent_change
  ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
            
         'Color negative changes in red
  If (yearly_change < 0) Then
  ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                   
         'Color no or positive changes in green
  ElseIf (yearly_change >= 0) Then
  ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
   
   End If
           
         'Continue on to next row
   Summary_Table_Row = Summary_Table_Row + 1
              
           'find the new year open for next ticker
   year_opn = ws.Cells(i + 1, 3).Value
              
           'Find the percent change higher than 0 and use as the greatest percent increase, then continue to next loop, and replace if higher. Grab the ticker from that row also
   If percent_change > greatest_percent_increase Then
   greatest_percent_increase = percent_change
   gpi_ticker = ticker
              
           'Find the percent change lower than 0 and use as the greatest percent decrease, then continue to next loop, and replace if lower. Grab the ticker from that row also
   ElseIf percent_change < greatest_percent_decrease Then
   greatest_percent_decrease = percent_change
   gpd_ticker = ticker
                  
            'Find the total volume greater than 0 and use as greatest total, and then continue to next loop, and replace if higher. Grab the ticker from that row also
   ElseIf vol_total > greatest_total Then
   greatest_total = vol_total
   gt_ticker = ticker
                
   End If
           
           'Define where to put the variables on the right side
   ws.Cells(2, 16).Value = gpi_ticker
   ws.Cells(3, 16).Value = gpd_ticker
   ws.Cells(4, 16).Value = gt_ticker
   ws.Cells(2, 17).Value = greatest_percent_increase
   ws.Cells(3, 17).Value = greatest_percent_decrease
   ws.Cells(4, 17).Value = greatest_total
            
          'turn these cells into a %
   ws.Cells(2, 17).NumberFormat = "0.00%"
   ws.Cells(3, 17).NumberFormat = "0.00%"
            
          'reset volume total for new ticker
   vol_total = 0
     
          'If we are not at the end of the first ticker row, so end of year, then proceed below
  Else
              
           'add the volume in cell it is viewing to the volume total
  vol_total = vol_total + ws.Cells(i, 7).Value
        
   End If

  Next i
    
 Next ws
      
End Sub


