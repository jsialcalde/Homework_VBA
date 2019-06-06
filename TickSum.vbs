Sub TickSum()
    
    Dim Ticker As String
    Dim Vol As Single
    Dim VolTotal As Single
    Dim sht As Worksheet
    Dim LastRow As Long
    Dim Summary_Table_Row As Integer
    Dim s_Close As Single
    Dim s_Start As Single
    Dim s_End As Single
    Dim sCount As Long
    Dim I As Long
    ' Begin the loop.
    For Each ws In Worksheets
        
        
        'Label Summary Table
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        'Set Initial Volume Total to 0
        VolTotal = 0
        
        ' For each sheet, create a summary table to hold the results
        ' Keep track of the location for stock ticker in the summary table
        
        Summary_Table_Row = 2
        
        
        ' Find last row of each sheet
        LastRow = ws.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row
        
        ' Set counter to 1 to count how many times a stock symbol appeared
        sCount = 1
        
        ' Loop through each row
        For I = 2 To LastRow
        
          ' For each row, find the ticker symbol
            Ticker = ws.Cells(I, 1).Value
                
            ' For each row, find the volume
            Vol = ws.Cells(I, 7).Value

            ' For each row, find the closing price
            s_Close = ws.Cells(I, 6).Value
            
            
            'Check if we are still looking at the same Ticker Symbol, if not:
            If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
                
                ' If the ticker symbol in next iteration is different, set closing price for end of year
                s_End = s_Close
                
                ' Add volume to running total
                VolTotal = VolTotal + Vol
                
                ' Print Ticker Symbol to summary table
                ws.Range("I" & Summary_Table_Row).Value = Ticker
                
                ' Print Running Total to Summary table
                ws.Range("L" & Summary_Table_Row).Value = VolTotal
                
                
                ' Set Closing price for beginning of each year of stock
                ' Set to 1 if = 0 (to avoid 0 division)
                If (ws.Cells(I - sCount + 1, 6) = 0) Then
                    s_Start = 1
                Else
                    s_Start = ws.Cells(I - sCount + 1, 6)
                
                End If
                
                ' Calculate difference in closing price for beginning and end of year
                ' Print difference in closing price for beginning and end of year
                ' Set color to green if price change was above 0, else red
                ws.Range("J" & Summary_Table_Row).Value = s_End - s_Start
                
                If (ws.Range("J" & Summary_Table_Row).Value < 0) Then
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                    
                    Else
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                End If
                                
                ' Calculate % difference in closing price for beginnning and end of year
                ' Print % difference in closing price for beginning and end of year
                ws.Range("K" & Summary_Table_Row).Value = (s_End - s_Start) / s_Start
                ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                
                ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
                
                ' Reset running total to 0
                VolTotal = 0
                
                ' Reset counter to 1
                sCount = 1
            
            Else
                ' Add volume to running total
                VolTotal = VolTotal + Vol
                sCount = sCount + 1
            End If
                
               
        
        Next I
        
    Next ws
        
        
End Sub


