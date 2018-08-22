Sub TickerVolume()

Dim ticker As String
Dim vol_total As Double
    vol_total = 0
Dim sum_table_row As Integer
    sum_table_row = 2
Dim r As Double

    
LastRow = Cells(Rows.Count, 1).End(xlUp).Row
Range("I1").Value = "Ticker"
Range("J1").Value = "Total Volume"

'Start the loop
For r = 2 To LastRow
   If Cells(r + 1, 1).Value <> Cells(r, 1).Value Then
   
        'Grab ticker symbol
        ticker = Cells(r, 1).Value
   
        'place ticker in summary table
         Cells(sum_table_row, 9).Value = ticker
   
        'add to volume total
         vol_total = vol_total + Cells(r, 7)
    
        'place volume total in summary table
        Cells(sum_table_row, 10).Value = vol_total
    
        'move to next line in summary table
        sum_table_row = sum_table_row + 1
    
        'reset volume total
        vol_total = 0
    
    Else
    
        'add to volume total
         vol_total = vol_total + Cells(r, 7)
    
    End If
        
Next r


End Sub
