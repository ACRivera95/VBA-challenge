# VBA-challenge

In order to succesfully write this code I used the following links/tools


  
    
    ' Loop through each worksheet in the workbook :
    
    https://support.microsoft.com/en-us/topic/macro-to-loop-through-all-worksheets-in-a-workbook-feef14e3-97cf-00e2-538b-5da40186e2b0
    
    For Each ws In ThisWorkbook.Worksheets
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        
       'I tried to use the for method but my excel kept crashing so I google how to read columns 
        
       ' Dim i As Long
       ' For i = lastRow To 1 Step -1
        '    If Not IsEmpty(ws.Cells(i, "A").Value) Then
         '       lastRow = i
          '      Exit For
           ' End If
        ' Next i
        
        ' Create a dictionary : https://excelmacromastery.com/vba-dictionary/
        Set tickers = CreateObject("Scripting.Dictionary")
        
        ' Loop through each cell in columns A, B, C, F, and in order to find all needed information
    
        For Each tickerCell In ws.Range("A2:A" & lastRow)
        ' Offset by # of columns needed to get the corresponding data
        I didn't really know how to use offset properly so I guide my self with : https://www.wallstreetmojo.com/vba-offset/
            Set dateCell = tickerCell.Offset(0, 1)
            Set openingCell = tickerCell.Offset(0, 2)
            Set closingCell = tickerCell.Offset(0, 5)
            
            
            Dim year As Long 'this was used in order to properly read the years in the date.
            year = Left(dateCell.Value, 4)
            
            ' This part of the code was used to verify if the ticker was already in the list in order to add ir or leave it.
            ' https://www.youtube.com/watch?v=WvXOYeanIj8 this guy was a great help.
            If Not tickers.Exists(tickerCell.Value) Then
             
                Set tickers(tickerCell.Value) = CreateObject("Scripting.Dictionary")
                tickers(tickerCell.Value)("Year") = year
                tickers(tickerCell.Value)("OpeningPrice") = openingCell.Value
            End If
            
            tickers(tickerCell.Value)("ClosingPrice") = closingCell.Value
            tickers(tickerCell.Value)("TotalVolume") = tickers(tickerCell.Value)("TotalVolume") + tickerCell.Offset(0, 6).Value
        Next tickerCell
        
        ' Write extracted data to the new columns
        ws.Range("M:P").ClearContents ' Clear previous data in columns M to P
        ws.Range("M1").Value = "Ticker"
        ws.Range("N1").Value = "Yearly Change"
        ws.Range("O1").Value = "Percentage Change"
        ws.Range("P1").Value = "Total Volume"
        
        Dim rowIndex As Long
        rowIndex = 2
        For Each Key In tickers.keys
            ' Calculate the yearly change
            yearlyChange = tickers(Key)("ClosingPrice") - tickers(Key)("OpeningPrice")
            
            ' Calculate the percentage change
            Dim percentageChange As Double
            If tickers(Key)("OpeningPrice") <> 0 Then
                percentageChange = yearlyChange / tickers(Key)("OpeningPrice") * 100
            Else
                percentageChange = 0
            End If

            ws.Cells(rowIndex, "M").Value = Key
            ws.Cells(rowIndex, "N").Value = yearlyChange
            ws.Cells(rowIndex, "O").Value = percentageChange
            ws.Cells(rowIndex, "P").Value = tickers(Key)("TotalVolume")
            
           
            If percentageChange > maxPercentageIncrease Then
                maxPercentageIncrease = percentageChange
                maxIncreaseTicker = Key
            End If
            
            If percentageChange < maxPercentageDecrease Then
                maxPercentageDecrease = percentageChange
                maxDecreaseTicker = Key
            End If
            
            If tickers(Key)("TotalVolume") > maxTotalVolume Then
                maxTotalVolume = tickers(Key)("TotalVolume")
                maxVolumeTicker = Key
            End If
            
            rowIndex = rowIndex + 1
        Next Key
        
        ' Print the new table : https://www.exceldemy.com/vba-excel-print/
        ws.Range("R2").Value = "Greatest % Increase"
        ws.Range("R3").Value = "Greatest % Decrease"
        ws.Range("R4").Value = "Greatest Total Volume"
        ws.Range("S2").Value = maxIncreaseTicker
        ws.Range("S3").Value = maxDecreaseTicker
        ws.Range("S4").Value = maxVolumeTicker
        ws.Range("T2").Value = maxPercentageIncrease & "%"
        ws.Range("T3").Value = maxPercentageDecrease & "%"
        ws.Range("T4").Value = maxTotalVolume
        
    Next ws
End Sub

