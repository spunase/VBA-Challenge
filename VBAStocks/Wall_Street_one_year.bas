Attribute VB_Name = "Wall_Street_one_year"
Sub Wall_Street_one_year()
    ' This program aggregates the Ticker percent change and yearly change for one year data from year 2014
    ' AUTHOR: Surabhi Mukati
    ' DATE  : 06-22-2020

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim open_val As Double
    Dim close_val As Double
    Dim percent_change As Double
    Dim total_vol As Currency
    Dim ticker As String
    Dim ticker_summary_row As Integer
      
    ' Initialize the variable
    total_vol = 0
     
    ' Turnoff automatic calculations an screen updating for faster processing
    With Application
       .Calculation = xlCalculationManual
       .ScreenUpdating = False
    End With

    Set ws = Sheets("2014")
    

     ' Compute the last unused row in the current worksheet
     lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
     
     ' Sort the date fields in ascending order as the code relies on that - this is optional when the data is already sorted
     Worksheets(ws.Name).Sort.SortFields.Add Key:=Range("B2:B" & lastRow), _
      SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
     
     ' Sort the ticker fields in ascending order as the code relies on that - this is optional when the data is already sorted
     Worksheets(ws.Name).Sort.SortFields.Add Key:=Range("A2:A" & lastRow), _
      SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
      
     
     ' Populate the headers of the summary table
     
      ws.Range("I1").Value = "Ticker"
      ws.Range("J1").Value = "Yearly Change"
      ws.Range("K1").Value = "Percent Change"
      ws.Range("L1").Value = "Total Stock Volume"
     
      ticker_summary_row = 2
         
         
      
   
      ' Loop through all records in the current worksheet to look for Open Value, Close Value and Aggregate Total volume per ticker
        
      For i = 2 To lastRow
      
        If i = 2 Then
             open_val = ws.Cells(2, 3).Value
        End If
        
        ' When the ticker is about to change, grab the close value for the current ticker and open value for the next ticker
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
           ' Save the unique ticker value in the summary table
           ticker = ws.Cells(i, 1).Value
           ws.Range("I" & ticker_summary_row).Value = ticker
           
           total_vol = total_vol + ws.Cells(i, 7).Value
           ws.Range("L" & ticker_summary_row).Value = total_vol   ' Save the total volume for this ticker
           ws.Columns("L:L").NumberFormat = "0"                   ' Remove currency formatting
           total_vol = 0                                          ' Reset the variable that stores total volume
              
            
           close_val = ws.Cells(i, 6)
             
           yearly_change = close_val - open_val
           
           ' Check for a divide by zero error in case the data has zeroes
           
           If open_val = 0 Then
                percent_change = 0
           Else
                percent_change = yearly_change / open_val
                
           End If
           
           ws.Range("J" & ticker_summary_row).Value = yearly_change
           If yearly_change > 0 Then
               ws.Range("J" & ticker_summary_row).Interior.ColorIndex = 4
               
           Else
               ws.Range("J" & ticker_summary_row).Interior.ColorIndex = 3
               
           End If
               
            
           ws.Range("K" & ticker_summary_row).Value = percent_change
           ws.Columns("K:K").NumberFormat = "0.00%"               ' Show as Percentage
            
             
           open_val = ws.Cells(i + 1, 3)                          ' Store the open value for the upcoming ticker
            
           ticker_summary_row = ticker_summary_row + 1
           
        Else
           total_vol = total_vol + ws.Cells(i, 7).Value
           
        End If
        
        
      Next i
          
          
      ' Restore parameters
      With Application
       .Calculation = xlCalculationAutomatic
       .ScreenUpdating = True
      End With
     
End Sub
