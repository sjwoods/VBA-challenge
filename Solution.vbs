Sub ticker()
    
    Dim ticker As String
    
    Dim totalstockvolume As Double
    
    Dim startval As Double
    
    Dim endval As Double
    
    For Each ws In Worksheets
    
    startval = ws.Cells(2, 3).Value
    
    endval = 0
    
    totalstockvolume = 0
    
    Dim tickerTotalRow As Integer
    tickerTotalRow = 2
    
    'instead of 800000, need to find last row
    
    For Row = 2 To 800000
        
        
        If ws.Cells(Row + 1, 1).Value <> ws.Cells(Row, 1).Value Then
            
            endval = ws.Cells(Row, 6).Value
            
            yearchange = endval - startval
            
            If startval <> 0 Then
        
                percentchange = yearchange / startval
            
            Else
        
                percentchange = 0
        
            End If
        
        
        'ElseIf (Cells(Row + 1, 1).Value <> Cells(Row, 1).Value) Then
            
            totalstockvolume = totalstockvolume + ws.Cells(Row, 7).Value
            
            ticker = ws.Cells(Row, 1).Value
            
            ws.Range("H" & tickerTotalRow).Value = ticker
            
            ws.Range("K" & tickerTotalRow).Value = totalstockvolume
            
            ws.Range("I" & tickerTotalRow).Value = yearchange
            
            ws.Range("J" & tickerTotalRow).Value = percentchange
            
            'after calculations of percentchange, if statement for color code
        If (yearchange > 0) Then
            
            ws.Range("I" & tickerTotalRow).Interior.ColorIndex = 4
            
            Else: ws.Range("I" & tickerTotalRow).Interior.ColorIndex = 3
        
        End If
        
            
            totalstockvolume = 0
            
            tickerTotalRow = tickerTotalRow + 1
            
            yearchange = 0
            
            startval = ws.Cells(Row + 1, 3).Value
        
        Else
            totalstockvolume = totalstockvolume + ws.Cells(Row, 7).Value
        End If
    Next Row
    Next ws
End Sub

