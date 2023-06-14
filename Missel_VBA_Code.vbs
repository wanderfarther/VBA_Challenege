Sub Changes()
' adding the worksheets as variables '
Dim ws As Worksheet

For Each ws In Worksheets

'Defining all my variables'

Dim LR As Long
Dim TickerName As String
Dim TickerSpot As Integer
Dim Table_Row As Integer
Dim Volume As Double
Dim Opening As Double
Dim Closing As Double
Dim Least_Increase As Double
Dim Great_Increase As Double


LR = Cells(Rows.Count, 1).End(xlUp).Row
TickerSpot = 2
Table_Row = 2
Opening = 0

'This will add your header to each sheet for the new info'
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    For I = 2 To LR
     
        'This will cycle through the data on each sheet to identify the ticker symbol and put it in the correct cell'
          
       If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
        
        ws.Range("I" & TickerSpot).Value = TickerName
        TickerSpot = TickerSpot + 1
        
        'Puts the volume measure in the correct cell for display'
        ws.Cells(Table_Row, 12).Value = Volume
        
        Volume = 0
       
       'This will find the lastest closing numbers'
        Closing = ws.Cells(I, 6)
        ws.Cells(Table_Row, 10).Value = Closing - Opening
            'This formats the cells with colors'
            If ws.Cells(Table_Row, 10).Value > 0 Then
                ws.Cells(Table_Row, 10).Interior.ColorIndex = 4
                Else: ws.Cells(Table_Row, 10).Interior.ColorIndex = 3
                End If
        'This formats the percentage column'
        ws.Cells(Table_Row, 11).NumberFormat = "00.00%"
        'This figures out the percentage change'
        If Closing > Opening Then
                ws.Cells(Table_Row, 11).Value = (Closing - Opening) / Opening
                Else
                ws.Cells(Table_Row, 11).Value = -((Opening - Closing) / Opening)
                End If
                
        'Reseting the opening value and moving the inputs to the next row in our table'
        Table_Row = Table_Row + 1
        Opening = 0
        
       Else
           Volume = Volume + ws.Range("G" & I) 'Adds the total volume together'
           TickerName = ws.Cells(I, 1).Value
           
           'This identifies the value of the opening price of a ticker'
           
           If Opening = 0 Then
            Opening = ws.Cells(I, 3)
            End If
            
       End If
        
    Next I

'This will add your header to each sheet for the new info'
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"

'This will give the least increased ticker'
Least_Increase = ws.Cells(2, 11).Value

   For j = 3 To LR
        
        If Least_Increase > ws.Cells(j, 11).Value Then
        Least_Increase = ws.Cells(j, 11).Value
        TickerName = ws.Cells(j, 9)
        Else
        Least_Increase = Least_Increase
        End If
                
    Next j
        
    ws.Range("P3").NumberFormat = "00.00%"
    ws.Range("P3").Value = Least_Increase
    ws.Range("O3").Value = TickerName
    
'This will give the greatest increase ticker'
Great_Increase = ws.Cells(2, 11).Value

   For j = 3 To LR
        
        If Great_Increase < ws.Cells(j, 11).Value Then
        Great_Increase = ws.Cells(j, 11).Value
        TickerName = ws.Cells(j, 9)
        Else
        Great_Increase = Great_Increase
        End If
                
    Next j
        
    ws.Range("P2").NumberFormat = "00.00%"
    ws.Range("P2").Value = Great_Increase
    ws.Range("O2").Value = TickerName

'This will give the greatest volume ticker'
Volume = ws.Cells(2, 12).Value

   For j = 3 To LR
        
        If Volume < ws.Cells(j, 12).Value Then
        Volume = ws.Cells(j, 12).Value
        TickerName = ws.Cells(j, 9)
        Else
        Volume = Volume
        End If
                
    Next j
        
    ws.Range("P4").Value = Volume
    ws.Range("O4").Value = TickerName
        
Next ws

End Sub
