Attribute VB_Name = "Module1"
'Part 1:
'   Ticker Symbol

Sub StockTickerAnalysis():


' Part 1 - Setting variables and columns/data
    '1. Set For Loop for Worksheets
        Dim ws As Worksheet
        For Each ws In Worksheets
    
    '2. Set variables
        Dim TickerName As String
        
        Dim YearlyChange As Double
            YearlyChange = 0
            
        Dim PercentChange As Double
        
        Dim TotalStockVolume As Double
            TotalStockVolume = 0
            
        Dim LastRow As Long
             LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
             
        Dim SummaryRow As Long
            SummaryRow = 2

        Dim OpenValue As Long
            OpenValue = 2
        
    '3. Set place for columns/data
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
'Part 2 - Initialize the loop!
    '1. What does "For I" mean? Start at Row 2 and cycle through until it gets to the last row. We use LastRow because there are too many rows to count and I don't know if it changes per worksheet.
        For i = 2 To LastRow
        
    '2. What we will do if the ticker information is not the same?
            
           'Start the If/Then Statement what what to do when tickers change
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                 
                'Identify where the tickers are
                TickerName = ws.Cells(i, 1).Value
                
                'Identify where total stock volume equals to
                
                TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
                
                'Identify where stock volume data will go
                
                ws.Range("L" & SummaryRow) = TotalStockVolume
                
               'Identify where the ticker names will go
                
                ws.Range("I" & SummaryRow) = TickerName
                
                'Identify what yearly change equals to
                
                YearlyChange = ws.Cells(i, 6).Value - ws.Cells(OpenValue, 3).Value
                
                'Identify where the yearly change data will go and format
                 
                 ws.Range("J" & SummaryRow) = YearlyChange
                 
                 ws.Range("J" & SummaryRow).NumberFormat = "0.00"
                
                'Identify what percent change is equal to
                
                    If ws.Cells(OpenValue, 3).Value = 0 Then
                        PercentChange = 0
                    Else
                    PercentChange = YearlyChange / ws.Cells(OpenValue, 3).Value
                    
                    End If
                 
               'Identify where the percentage data will go and format
                
                ws.Range("k" & SummaryRow) = PercentChange
                
                ws.Range("k" & SummaryRow).NumberFormat = "0.00%"
                
                'Identify to cycle through each row of data
                
                SummaryRow = SummaryRow + 1
                
               'Reset the total stock volume
                
                TotalStockVolume = 0
                
               'Once you get to the next unique ticker, the open value row will need to increase an increment
                
                OpenValue = i + 1
                
               ' Reset the yearly change
                
                YearlyChange = 0
            
       '3. if the tickers are the same, it this will add the stock volume together.
            Else
            TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
            
           'Need to end the if statement so hat the loop knows where to end
            End If
            
        
       '4. Conditional formatting
            If ws.Cells(SummaryRow, 10).Value >= 0 Then
               ws.Cells(SummaryRow, 10).Interior.ColorIndex = 4
            
            Else
                ws.Cells(SummaryRow, 10).Interior.ColorIndex = 3
                
            End If
        
    'Indicate for the loop to move on to the next iteration
    
    Next i
    
Next ws

End Sub

