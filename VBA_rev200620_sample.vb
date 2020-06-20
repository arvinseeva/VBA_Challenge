Sub heading():

'for each worksheet

For Each ws In Worksheets


    ' provide headers
    
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Price Difference"
    ws.Cells(1, 11).Value = "Total Percentage (in %)"
    ws.Cells(1, 12).Value = "Total Stocks"
    ws.Cells(1, 13).Value = "Opening Price"
    ws.Cells(1, 14).Value = "Closing Price"
    ws.Cells(2, 16).Value = "Greatest Profit"
    ws.Cells(3, 16).Value = "Greatest Loss"
    ws.Cells(4, 16).Value = "Greatest Volume"
    ws.Cells(1, 17).Value = "Ticker"
    ws.Cells(1, 18).Value = "Value"
    
    ' list the tickers
    
    Dim a As Long
    Dim lastrow As Long
    Dim tickername As String
    Dim tickerloc As Integer
    Dim volume As Long
    
    
    
    'determine last row
    
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'define date
    
       
    Dim startyear As Long
    Dim endyear As Long
    
    

    If ws.Cells(2, 2).Value = 20160101 Then
    startyear = 20160101
    endyear = 20161230
    
    ElseIf ws.Cells(2, 2).Value = 20150101 Then
    startyear = 20150101
    endyear = 20151230
    
    ElseIf ws.Cells(2, 2).Value = 20140101 Then
    startyear = 20140101
    endyear = 20141230
    
    End If
    
        
    'list out tickers and total volume at year end
    
    
    tickerloc = 2
    volume = 0
    
    
    For a = 2 To lastrow
        
        If ws.Cells(a + 1, 1).Value <> ws.Cells(a, 1).Value Then
        tickername = ws.Cells(a, 1).Value
        
        volume = volume + ws.Cells(a, 7).Value
        
        ws.Range("I" & tickerloc).Value = tickername
        
        ws.Range("L" & tickerloc).Value = volume
        
        tickerloc = tickerloc + 1
        
                                
        End If
        
                    
            
            
            
    Next a
    
   
    
       
    'inserting opening price
    
    
    
    Dim lastrow2 As Long
    Dim c As Long
    
    a = 2
    
    
    
    lastrow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    For c = 2 To lastrow2
        
        For a = 2 To 1000
        
        If ws.Cells(c, 9).Value = ws.Cells(a, 1).Value And ws.Cells(a, 2).Value = startyear Then
            ws.Range("M" & c).Value = ws.Cells(a, 3).Value
            
                 
        End If
        
        
        Next a
        
    Next c
    
    'insert closing price
    
    
    Dim d As Long
    
    a = 2
    
    
    For d = 2 To lastrow2
        
        For a = 2 To 1000
        
            
        If ws.Cells(d, 9).Value = ws.Cells(a, 1).Value And ws.Cells(a, 2).Value = endyear Then
            ws.Range("N" & d).Value = ws.Cells(a, 6).Value
            
           
        End If
        
        
        Next a
        
    Next d
    
    'calculate the price difference from opening price and closing price
    
    
    Dim e As Long
    
    
    
    For e = 2 To lastrow2
    
        If IsEmpty(ws.Cells(e, 14).Value) Or IsEmpty(ws.Cells(e, 13).Value) Then
        ws.Cells(e, 10).Value = 0
        
        
        
        Else: ws.Cells(e, 10).Value = ws.Cells(e, 14).Value - ws.Cells(e, 13).Value
         
        
        
        
        
        
        End If
        
        

     
        
             
     
    Next e
    
    'calculate annual percentage difference
    
    
      
    Dim f As Long
    
    For f = 2 To lastrow2
    
        If ws.Cells(f, 10).Value = 0 Then
        
        ws.Cells(f, 11).Value = "Not full year"
        
    
        
        Else: ws.Cells(f, 11).Value = ws.Cells(f, 10).Value / ws.Cells(f, 13).Value * 100
        
        
        
            If ws.Cells(f, 10).Value > 0 Then
        
        'colour cell for profit or loss
        
            ws.Cells(f, 10).Interior.ColorIndex = 4
        
            Else: ws.Cells(f, 10).Interior.ColorIndex = 3
        
            End If
        
        
        
        
            
        
        End If
            
        
    Next f
    
        
    'determine greatest profit or loss or volume
    
    
 
    
    Dim g As Long
    Dim h As Long
    Dim k As Long
    Dim greatprofit As Long
    Dim tickerprofit As Long
    Dim tickerloss As Long
    Dim greatloss As Long
    Dim greatvolume As Long
    Dim tickervolume As Long
    
    
    
    greatprofit = 0.01
    greatloss = -0.01
    greatvolume = 1
    
    
    For g = 2 To lastrow2
    
        If ws.Cells(g, 10).Value > greatprofit Then
        ws.Cells(2, 18).Value = ws.Cells(g, 10).Value
        ws.Cells(2, 17).Value = ws.Cells(g, 9).Value
        greatprofit = ws.Cells(g, 10).Value
        
        End If
              

               
    Next g
    
    h = 2
    
    For h = 2 To lastrow2
    
        If ws.Cells(h, 10).Value < greatloss Then
        ws.Cells(3, 18).Value = ws.Cells(h, 10).Value
        ws.Cells(3, 17).Value = ws.Cells(h, 9).Value
        greatloss = ws.Cells(h, 10).Value
        
        End If
    
    Next h
    
 
    k = 2
    
    For k = 2 To lastrow2
    
        If ws.Cells(k, 12).Value > greatvolume Then
        ws.Cells(4, 18).Value = ws.Cells(k, 12).Value
        ws.Cells(4, 17).Value = ws.Cells(k, 9).Value
        greatvolume = ws.Cells(k, 12).Value
        
        End If
        
    Next k
    


Next ws

    

End Sub






