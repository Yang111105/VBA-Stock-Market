
Sub Stock_market_analysis()

  For Each ws In Worksheets
  
      ' Find out the number of rows
      Dim lastrow As Double
      lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
      'MsgBox (lastrow)
      
      ' Set the frame for the summary table
      ws.Range("I1").Value = "Ticker"
      ws.Range("J1").Value = "Yearly Change"
      ws.Range("K1").Value = "Percent Change"
      ws.Range("L1").Value = "Total Stock Volme"
      
      Dim summary_table_row As Integer
      summary_table_row = 2
      
      ' Set initial variables for the summary table
      Dim ticker As String
      Dim yearly_open As Double
      Dim yearly_close As Double
      Dim yearly_change As Double
      Dim pct_change As Double
      Dim stock_volume As Double
      
      ' Set initial values
      yearly_open = ws.Cells(2, 3).Value
      yearly_close = ws.Cells(2, 6).Value
      stock_volume = 0
      
      ' Loop through all tickers to summarize change/volume per ticker
      Dim i As Double
      
      For i = 2 To lastrow
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
          
          ticker = ws.Cells(i, 1).Value
    
          stock_volume = stock_volume + ws.Cells(i, 7).Value
          
          yearly_close = ws.Cells(i, 6).Value
          
          yearly_change = yearly_close - yearly_open
          
          ws.Range("I" & summary_table_row).Value = ticker
              
          ws.Range("J" & summary_table_row).Value = yearly_change
              
          ws.Range("L" & summary_table_row).Value = stock_volume
              
              If yearly_open <> 0 Then
              
              pct_change = yearly_change / yearly_open
              
              ws.Range("K" & summary_table_row).Value = pct_change
              
              Else
              
              ws.Range("K" & summary_table_row).Value = 0
              
              End If
          
          ' Reset stock volume and open/close
          stock_volume = 0
          
          yearly_open = ws.Cells(i + 1, 3).Value
          
          yearly_close = ws.Cells(i + 1, 6).Value
                  
          summary_table_row = summary_table_row + 1
    
        Else
    
          stock_volume = stock_volume + ws.Cells(i, 7).Value
    
        End If
        
    Next i
      
      ' Conditional Formatting
      For j = 2 To summary_table_row
      
        If ws.Cells(j, 10).Value > 0 Then
        
          ws.Cells(j, 10).Interior.ColorIndex = 4
        
        ElseIf ws.Cells(j, 10).Value < 0 Then
        
          ws.Cells(j, 10).Interior.ColorIndex = 3
          
        End If
        
      Next j
      
      ' Adjust the format for summary table
      ws.Columns("K:K").NumberFormat = "0.00%"
      
      ws.Columns("I:L").AutoFit
  
  Next ws
  
End Sub
  
' ----------------------------Bunos Question----------------------------------

Sub Stock_Market_Bonus_Analysis()
  
  For Each ws In Worksheets
  
      ' Set the frame of the additional analysis table
      ws.Range("P1").Value = "Ticker"
      ws.Range("Q1").Value = "Value"
      ws.Range("O2").Value = "Greatest % Increase"
      ws.Range("O3").Value = "Greatest % Decrease"
      ws.Range("O4").Value = "Greatest Total Volume"
      
      ' Find out number of summary table rows
      Dim lastsummaryrow As Double
      lastsummaryrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
      ' MsgBox (lastsummaryrow)
      
      Dim additional_table_row As Integer
      additional_table_row = 2
      
      ' Set initial variables for the additional analysis table
      Dim greatest_ticker_increase As String
      Dim greatest_ticker_decrease As String
      Dim greatest_ticker_volume As String
      Dim greatest_increase As Double
      Dim greatest_decrease As Double
      Dim greatest_volume As Double
      
      ' Set initial value for greatest increase/decrease variables
      greatest_increase = ws.Cells(2, 11).Value
      greatest_decrease = ws.Cells(2, 11).Value
      greatest_volume = ws.Cells(2, 12).Value
      
      ' Loop through all distinctive tickers to identify the ones with greates increase/decrease
      Dim j As Double
      For j = 2 To lastsummaryrow
    

        If ws.Cells(j + 1, 11).Value > greatest_increase Then
          
          greatest_ticker_increase = ws.Cells(j + 1, 9).Value
    
          greatest_increase = ws.Cells(j + 1, 11).Value
      
        ElseIf ws.Cells(j + 1, 11).Value < greatest_decrease Then
          
          greatest_ticker_decrease = ws.Cells(j + 1, 9).Value
          
          greatest_decrease = ws.Cells(j + 1, 11).Value
        
        ElseIf ws.Cells(j + 1, 12).Value > greatest_volume Then
          
          greatest_ticker_volume = ws.Cells(j + 1, 9).Value
    
          greatest_volume = ws.Cells(j + 1, 12).Value
    
        End If
        
      Next j
            
      ' Print the additional analysis table
        ws.Range("P2").Value = greatest_ticker_increase
        
        ws.Range("P3").Value = greatest_ticker_decrease
        
        ws.Range("P4").Value = greatest_ticker_volume
        
        ws.Range("Q2").Value = greatest_increase
        
        ws.Range("Q3").Value = greatest_decrease
        
        ws.Range("Q4").Value = greatest_volume
      
      ' Adjust the format
      ws.Range("Q2", "Q3").NumberFormat = "0.00%"
      
      ws.Range("Q4").NumberFormat = "0.0000E+00"
      
      ws.Columns("O:Q").AutoFit
    
  Next ws
  
End Sub

