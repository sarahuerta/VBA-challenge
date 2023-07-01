# VBA-challenge
Sub Stock_Market():

 Dim Ticker As String
 Dim Opening As Double
 Dim Closing As Double
 Dim Percent_Change As Double
 Dim Yearly_Change As Double
 Dim Row_Table As Long
 Dim Last_Row As Long
 Dim Stock_Volume As Double
 Dim ws As Worksheet

For Each ws In Worksheets
 
 
 Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row
 

 ws.Cells(1, 9).Value = "Ticker"
 ws.Cells(1, 10).Value = "Yearly Change"
 ws.Cells(1, 11).Value = "Percent Change"
 ws.Cells(1, 12).Value = "Stock Volume"
 
Opening = 0
 Closing = 0
 Row_Table = 2
 Stock_Volume = 0
 Yearly_Change = 0
 Percent_Change = 0

 For i = 2 To Last_Row
  

  If Opening = 0 Then
   Opening = ws.Cells(i, 3).Value
  End If
  

  If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
  
   
    Ticker = ws.Cells(i, 1).Value
    ws.Range("I" & Row_Table).Value = Ticker
    
   
    
    Closing = ws.Cells(i, 6).Value
   
 
    Yearly_Change = Closing - Opening
    ws.Range("J" & Row_Table).Value = Yearly_Change
    ws.Range("J" & Row_Table).NumberFormat = "$0.00"
   
   
     If Closing > Opening Then
     ws.Range("J" & Row_Table).Interior.ColorIndex = 50
     ElseIf Closing < Opening Then
      ws.Range("J" & Row_Table).Interior.ColorIndex = 53
     Else: ws.Range("J" & Row_Table).Interior.ColorIndex = 44
     End If
    

    If Opening = 0 Then
     Percent_Change = 0
    Else
     Percent_Change = (Yearly_Change / Opening)
    End If
    ws.Range("K" & Row_Table).Value = Percent_Change
    ws.Range("K" & Row_Table).NumberFormat = "0.00%"
   
    Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
    ws.Range("L" & Row_Table).Value = Stock_Volume
    

    Row_Table = Row_Table + 1
    
    Opening = 0
    Closing = 0
    Row_Table = 2
    Stock_Volume = 0
    Yearly_Change = 0
    Percent_Change = 0
    

   Else: Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
  

  End If
   

 Next i


Next ws
  
End Sub
