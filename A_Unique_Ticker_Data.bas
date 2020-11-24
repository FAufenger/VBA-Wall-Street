Attribute VB_Name = "A_Unique_Ticker_Data"

Sub ForLoop_Ticker()
        
'Label headers

    Range("J1") = "Ticker"
    Range("K1") = "Yearly Change"
    Range("L1") = "Percent Change"
    Range("M1") = "Total Stock Volume"
    
 'Cell Size
    Columns("J:J").ColumnWidth = 15
    Columns("K:K").ColumnWidth = 15
    Columns("M:M").ColumnWidth = 17
    Columns("L:L").ColumnWidth = 18
        
Dim Column As Integer
Dim Row As Long
Dim i As Long
Dim Start As Long

LastRow = Cells(Rows.Count, "A").End(xlUp).Row
Column = 1
Row = 2
Start = 2
Total = 0

     For i = 2 To LastRow
     
       If Cells(i, Column).Value <> Cells(i + 1, Column).Value Then
           
           
          'Pull all unique tickers from column A
          Cells(Row, 10).Value = (Cells(i, Column).Value)
          
          'Yearly change
          Yearly_Change = Cells(i, 6).Value - Cells(Start, 3).Value
          
          Cells(Row, 11).Value = Yearly_Change
          
               'Percent Change able to work with / 0
               If Cells(Start, 3) = 0 Then
                  Cells(Row, 12) = 0 & "%"
    
               Else
                  Cells(Row, 12) = ((Yearly_Change / Cells(Start, 3).Value) * 100) & "%"

               End If
         
          'include value for last cell in unique ticker
          Total = Total + Cells(i, 7).Value
          
          'write in total volume
          Cells(Row, 13).Value = Total
            
          'move down a row for every value entered
          Row = Row + 1
          
          'start value changes with each unique ticker value
          Start = i + 1
          
          'reset total for new uique ticker
          Total = 0
          
        Else
            'add total volume while looping
            Total = Total + Cells(i, 7).Value
    
       End If
       
    
    Next i


LastTicker = Cells(Rows.Count, "J").End(xlUp).Row
Row = 2

    For i = 2 To LastTicker
 
       If Cells(i, 11).Value > 0 Then
  
            Cells(i, 11).Interior.ColorIndex = 4
  
       Else
            Cells(i, 11).Interior.ColorIndex = 3
  
       End If
     
    Next i
  
End Sub

