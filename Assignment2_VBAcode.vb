Sub ticker()
'set worksheet name
For Each ws In Worksheets
    
    'Set a variable to hold the worksheets
    Dim worksheetname As String
    
    'set an initial variable to hold the ticker names
    Dim ticker_name As String

    'set an initial variable to hold the ticker volumes
    Dim ticker_total As Double
    
    'Set the ticker_total to start at 0
    ticker_total = 0

    'keep track of where we want the computer to put the data
    Dim summary_table_row As Integer
    
    'let the computer know that we want it to start with the second row when inputting the data
    summary_table_row = 2

    'get the row number of the last row with data
    RowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
   
    'check to see that it's counting the rows correctly
    MsgBox (RowCount)
    
    'add the headers "Ticker" and "Total Stock Volume" to collumns L and M
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Total Stock Volume"
'create a loop that goes through the data starting with the second row
For i = 2 To RowCount
    
    'check to make sure we're still keeping track of the same ticker, and if it isn't then...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
    'set the ticker name
     ticker_name = ws.Cells(i, 1).Value
    
    'add to the ticker total
     ticker_total = ticker_total + ws.Cells(i, 7).Value
    
    'tell the computer where to put the ticker names in the summary tabel
     ws.Range("I" & summary_table_row).Value = ticker_name
    
    'tell the computer where to put the ticker totals
     ws.Range("J" & summary_table_row).Value = ticker_total
    
    'tell the computer to go to the next row after moving onto the next ticker name
     summary_table_row = summary_table_row + 1
    
    'reset the total to zero when starting with a new ticker name
     ticker_total = 0

'otherwise, if the cell immediately following has the same ticker
Else
    
    'keep on adding them up
     ticker_total = ticker_total + ws.Cells(i, 7).Value
    
    End If
  
  'move onto the next row in the worksheet
  Next i

'move onto the next worksheet in the workbook
Next ws

End Sub
