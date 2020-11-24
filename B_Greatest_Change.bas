Attribute VB_Name = "B_Greatest_Change"
Sub GreatestChange()


' Declare Current as a worksheet object variable.
Dim ws As Worksheet

' Loop through all of the worksheets in the active workbook.
  For Each ws In Worksheets


'Set Lables
    ws.Range("O2") = "Greatest % increase"
    ws.Range("O3") = "Greatest % decrease"
    ws.Range("O4") = "Greatest total volume"
    ws.Range("P1") = "Ticker"
    ws.Range("Q1") = "Value"
    
'set width
    ws.Columns("O:O").ColumnWidth = 20
    ws.Columns("P:P").ColumnWidth = 10
    ws.Columns("Q:Q").ColumnWidth = 15

'Find Value for Greatest % increase, then find matching ticker
    
    ws.Cells(2, 17).Value = Application.Max(ws.Range("L:L"))
    ws.Cells(2, 16).Value = Application.Index(ws.Range("J:J"), Application.Match(ws.Range("Q2"), ws.Range("L:L"), 0))
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
    

'Find Value for Greatest % decrease, then find matching ticker
    
    ws.Cells(3, 17).Value = Application.Min(ws.Range("L:L"))
    ws.Cells(3, 16).Value = Application.Index(ws.Range("J:J"), Application.Match(ws.Range("Q3"), ws.Range("L:L"), 0))

    
    
'Find Value for Greatest total volume, then find matching ticker
    
    ws.Cells(4, 17).Value = Application.Max(ws.Range("M:M"))
    ws.Cells(4, 16).Value = Application.Index(ws.Range("J:J"), Application.Match(ws.Range("Q4"), ws.Range("M:M"), 0))

  
' This line displays the worksheet name in a message box.
 
 Next ws

End Sub

