Attribute VB_Name = "Module1"
Sub stocks()
  ' Loop through each sheet
  For Each ws In Worksheets

      ' Define variables to capture ticker name, yearly change and  percent
      Dim ticker_name As String
      Dim yearly_change As Double
      Dim percent As Double
    
      ' Define variable to capture total stock volume, initialize at zero
      Dim ticker_total As Double
      ticker_total = 0
      
      ' Define variables to capture row number for the first and the last record for that ticker
      Dim first_ticker As Double
      Dim last_ticker As Double
    
      ' Keep track of the location for each ticker name in the summary table, data starts at row 2
      Dim summary_table_row As Integer
      summary_table_row = 2
      
      ' Define variable to capture the last row with data
      Dim last_row As Double
      last_row = ws.Cells(Rows.count, 1).End(xlUp).Row
      
      ' Populate headers for the summary table
      ws.Cells(1, 9).Value = "Ticker"
      ws.Cells(1, 10).Value = "Yearly Change"
      ws.Cells(1, 11).Value = "Percent Change"
      ws.Cells(1, 12).Value = "Total Stock Volume"
    
          ' Loop through all ticker records
          For i = 2 To last_row
        
            ' Check to see if the ticker is on its last record by determining if the next record is something different, if this is the case then execute the following
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
              ' Capture the ticker name
              ticker_name = ws.Cells(i, 1).Value
        
              ' Add to total stock volume
              ticker_total = ticker_total + ws.Cells(i, 7).Value
              
              ' Grab row number of last ticker
              last_ticker = i
              
              ' Calculate yearly change
              yearly_change = ws.Cells(last_ticker, 6).Value - ws.Cells(first_ticker, 3).Value
              
              ' Calculate percent
              If ws.Cells(first_ticker, 6).Value <> 0 Then
                percent = (ws.Cells(last_ticker, 6).Value - ws.Cells(first_ticker, 3).Value) / ws.Cells(first_ticker, 6).Value
              End If
        
              ' Print the ticker name, yearly change, percent change and total stock volume in the Summary Table
              ws.Range("I" & summary_table_row).Value = ticker_name
              ws.Range("J" & summary_table_row).Value = yearly_change
              ws.Range("K" & summary_table_row).Value = Format(percent, "Percent")
              ws.Range("L" & summary_table_row).Value = ticker_total
                          
              ' Add one to the summary table row
              summary_table_row = summary_table_row + 1
              
              ' Reset the total stock volume
              ticker_total = 0
        
            ' If the ticker is not on its last record, execute the following
            Else
        
              ' Sum up the total stock volume for that ticker
              ticker_total = ticker_total + ws.Cells(i, 7).Value
              
              ' Grab row number of first ticker
              If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                first_ticker = i
              End If
        
            End If
        
          Next i
          
        For j = 2 To last_row
        
            ' Format yearly change green for positive, red for negative
            If ws.Cells(j, 10).Value >= 0 Then
              ws.Cells(j, 10).Interior.ColorIndex = 4
            Else
              ws.Cells(j, 10).Interior.ColorIndex = 3
            End If
            
        Next j
          
  Next ws

End Sub
