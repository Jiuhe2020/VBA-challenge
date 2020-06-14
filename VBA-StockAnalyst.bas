Attribute VB_Name = "Module1"
Sub StockAnalyst()
    
    For Each ws In Worksheets
    
    ' Set the title for each column (from I to L)
    ws.Range("I1:L1").Value = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")
    
    ' Set i as row to check the entire data set
    Dim i As Variant
    ' Set j as the row to enter each ticker in the output table
    Dim j As Integer
    ' Set total to store the Total Stock Volume
    Dim total As Variant
    ' Set beginofyear to store the open price at the beginning of the year of each ticker
    Dim beginofyear As Variant
    ' Set endofyear to store the closing price at the end of the year of each ticker
    Dim endofyear As Variant
    ' Set PercentChange to calculate the Percent Change and format it into %
    Dim PercentChange As Variant
    ' Set LastRow to store the total number of rows containing data
    Dim LastRow As Long
    
    j = 2
    total = 0
    beginofyear = ws.Cells(j, 6).Value
    ' Calculate the total number of rows to set the range of i
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To LastRow
      If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
          ' Enter a ticker in column I
          ws.Cells(j, 9).Value = ws.Cells(i, 1).Value
          ' Find the closing price at the end of the year
          endofyear = ws.Cells(i, 6).Value
          ' Calculate the Yearly Change and enter column J
          ws.Cells(j, 10).Value = endofyear - beginofyear
            ' Conditional formatting the color of the cell based on Yearly Change being positive or negative
            If ws.Cells(j, 10).Value > 0 Then
              ws.Cells(j, 10).Interior.ColorIndex = 4
              ElseIf ws.Cells(j, 10).Value < 0 Then
              ws.Cells(j, 10).Interior.ColorIndex = 3
            End If
            
              ' Check if beginofyear is 0
              If beginofyear = 0 Then
                    ws.Cells(j, 11).Value = "N/A"
                    Else
                    ' If it is not 0, calculate the Percent Change and enter column K
                    PercentChange = ws.Cells(j, 10).Value / beginofyear
                    ws.Cells(j, 11).Value = Format(PercentChange, "0.00%")
                End If
          ' Add up the stock volume
          total = total + ws.Cells(i, 7).Value
          ws.Cells(j, 12).Value = total
          ' Rest the total for the next ticker
          total = 0
          ' Find the open price at the beginning of the year for the next ticker
          beginofyear = ws.Cells(i + 1, 6).Value
          j = j + 1
          Else
          total = total + ws.Cells(i, 7).Value
          
        End If
    Next i
    
    
    ' BONUS
    ' Find the "Greatest % increase", "Greatest % decrease" and "Greatest total volume" of each year
    Dim MaxIncrease As Variant
    Dim MaxDecrease As Variant
    Dim MaxTotal As Variant
    Dim R1, R2, R3 As Integer
    
    ws.Range("P1:Q1").Value = Array("Ticker", "Value")
    ws.Range("O2").Value = "Greatest % increase"
    ws.Range("O3").Value = "Greatest % decrease"
    ws.Range("O4").Value = "Greatest total volume"
    
    
    ' Find the "Greatest % increase"
    MaxIncrease = WorksheetFunction.Max(ws.Range("K:K"))
    ws.Cells(2, 17).Value = Format(MaxIncrease, "0.00%")
    ' Find the row of the cell that contains MaxIncrease
    R1 = WorksheetFunction.Match(MaxIncrease, ws.Range("K:K"), 0)
    ws.Cells(2, 16).Value = ws.Cells(R1, 9).Value
    
    ' Find the "Greatest % decrease"
    MaxDecrease = WorksheetFunction.Min(ws.Range("K:K"))
    ws.Cells(3, 17).Value = Format(MaxDecrease, "0.00%")
    ' Find the row of the cell that contains MaxDecrease
    R2 = WorksheetFunction.Match(MaxDecrease, ws.Range("K:K"), 0)
    ws.Cells(3, 16).Value = ws.Cells(R2, 9).Value
    
    ' Find the "Greatest total volume"
    MaxTotal = WorksheetFunction.Max(ws.Range("L:L"))
    ws.Cells(4, 17).Value = MaxTotal
    ' Find the row of the cell that contains MaxTotal
    R3 = WorksheetFunction.Match(MaxTotal, ws.Range("L:L"), 0)
    ws.Cells(4, 16).Value = ws.Cells(R3, 9).Value
    
    ' Autofit to display data
    ws.Columns("I:Q").AutoFit
    
    
    Next ws
End Sub

