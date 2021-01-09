Attribute VB_Name = "Module1"

Sub ALPHABET1():

        Dim WS_Count As Integer
         Dim m As Integer
         
         WS_Count = ActiveWorkbook.Worksheets.Count

         ' Begin the loop.
    For m = 1 To WS_Count
    
      ' defining all variables directly after the worksheet loop starts
      
          
            Dim tickernew As String
            
            Dim StockOpen As Integer
            Dim StockClose As Integer
        
'-----------------------------------------------------------------------


Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

Dim Vol_total As Double
Vol_total = 0

'send to last row
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

        
    '-----------MAIN LOOP ------------
For I = 2 To lastrow



  If Cells(I - 1, 1).Value <> Cells(I, 1).Value Then
    
        StockOpen = Cells(I, 3).Value
   

'if the TICKER BELOW is different
        ElseIf Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
     '----------------------------------------
            StockClose = Cells(I, 6).Value
            
            tickernew = Cells(I, 1).Value
    
            Vol_total = Vol_total + Cells(I, 7).Value
            
            StockTotal = StockClose - StockOpen
                        
                        If StockOpen > 0 Then
            Percentage = (StockClose / StockOpen) - 1
                           End If
            
      Range("I" & Summary_Table_Row).Value = tickernew
      Range("L" & Summary_Table_Row).Value = Vol_total
     
            
            Vol_total = 0
            
        Range("J" & Summary_Table_Row).Value = StockTotal
        Range("K" & Summary_Table_Row).Value = Percentage
 
     Summary_Table_Row = Summary_Table_Row + 1
           
            
    'if the TICKER ABOVE is different
  
         Else
               
        Vol_total = Vol_total + Cells(I + 1, 7).Value



        End If
        Range("K2:K2836").NumberFormat = "0.0%"
        Range("J2:J2836").NumberFormat = "0.00"

        

        Next I
        
         
        If StockTotal >= 0 Then
        
            Range("J2:J2836").Interior.ColorIndex = 4
            
            Else
                 Range("J2:J2836").Interior.ColorIndex = 3
        End If
        
    Next m
End Sub
