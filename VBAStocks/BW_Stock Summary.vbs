Sub stock_summary()
  
  For Each ws In Worksheets

  Dim ticker_symbol As String

  Dim volume As Double
  volume = 0
  
  Dim first_px As Double
  first_px = 0
  
  Dim last_px As Double
  last_px = 0

  ws.Range("I1").Value = "Ticker"
  ws.Range("J1").Value = "Yearly Change"
  ws.Range("K1").Value = "Percent Change"
  ws.Range("L1").Value = "Total Stock Volume"
  
  Dim summary_table_row As Integer
  summary_table_row = 2
  
  Dim lastrow As Long
  lastrow = ws.Cells(Rows.Count, 1).End(xlUp).row
  
  Dim row As Long
  
  For row = 2 To lastrow

    If ws.Cells(row - 1, 1).Value <> ws.Cells(row, 1).Value Then
      
      ' Set the opening price
      first_px = ws.Cells(row, 3).Value
    
      ' Adjust if opening price equal to 0 (pull next non-zero price)
      If first_px = 0 Then
        For i = row To lastrow
            If ws.Cells(i + 1, 3).Value <> 0 And ws.Cells(i, 3).Value = 0 Then
            first_px = ws.Cells(i + 1, 3).Value
            End If
        Next i
      End If
          
    End If
        
    ' Check if we are still within the same ticker, if it is not...
    If ws.Cells(row + 1, 1).Value <> ws.Cells(row, 1).Value Then

      ' Set the ticker symbol
      ticker = ws.Cells(row, 1).Value

      ' Set the last price
      last_px = ws.Cells(row, 6).Value

      ' Add to the volume total
      volume = volume + ws.Cells(row, 7).Value

      ' Print the ticker in the summary table
      ws.Range("I" & summary_table_row).Value = ticker

      ' Print the yearly change to the summary table
      ws.Range("J" & summary_table_row).Value = last_px - first_px

        ' Format the yearly change in the summary table
        If ws.Range("J" & summary_table_row).Value > 0 Then
            ws.Range("J" & summary_table_row).Interior.ColorIndex = 4
        Else
            ws.Range("J" & summary_table_row).Interior.ColorIndex = 3
        End If
      
      ' Print the percent change to the summary table
      ws.Range("K" & summary_table_row).Value = (last_px - first_px) / first_px

      ' Format the percent change in the summary table
      ws.Range("K" & summary_table_row).NumberFormat = "0.00%"
      
      ' Print the volume to the summary table
      ws.Range("L" & summary_table_row).Value = volume

      ' Add one to the summary table row
      summary_table_row = summary_table_row + 1
      
      ' Reset the volume
      volume = 0
      
    ' If the cell immediately following a row is the same ticker...
    Else

      ' Add to the volume
      volume = volume + ws.Cells(row, 7).Value

    End If

  Next row

Next ws

End Sub