Sub AllStocksAnalysisRefactored()
    
    'define variable yearValue to collect input
    yearValue = InputBox("What year would you like to run the analysis on?")

    'define variables for the timer
    Dim startTime As Single
    Dim endTime  As Single
    
    'starts the timer immediately after the year is input
    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
        
        'sheet title
        Range("A1").Value = "All Stocks (" + yearValue + ")"
        
        'header row
        Cells(3, 1).Value = "Ticker"
        Cells(3, 2).Value = "Total Daily Volume"
        Cells(3, 3).Value = "Return"
        
    'activate worksheet for indicated year
    Worksheets(yearValue).Activate
        

        'initialize array with 12 positions for tickers
        Dim tickers(12) As String
        
        tickers(0) = "AY"
        tickers(1) = "CSIQ"
        tickers(2) = "DQ"
        tickers(3) = "ENPH"
        tickers(4) = "FSLR"
        tickers(5) = "HASI"
        tickers(6) = "JKS"
        tickers(7) = "RUN"
        tickers(8) = "SEDG"
        tickers(9) = "SPWR"
        tickers(10) = "TERP"
        tickers(11) = "VSLR"
        
        'create a ticker Index
        tickerIndex = 0
        
        'create three output arrays to hold the volumes
        Dim tickerVolumes(12) As Long
        Dim tickerStartingPrices(12) As Double
        Dim tickerEndingPrices(12) As Double
        
        'Get the number of rows to loop over
        Dim RowEnd As Integer
        RowEnd = Cells(Rows.Count, "A").End(xlUp).Row
        
        
        'create the for loop to initialize all the tickerVolumes to zero
        For i = 0 To 11

            tickerVolumes(i) = 0
            
        Next i

        'Loop over all the rows in the spreadsheet to obtain volume and start and end prices
        
        For i = 2 To RowEnd
   
            'increase volume for current ticker as indicated by the ticker index
            If Cells(i, 1).Value = tickers(tickerIndex) Then
                
                tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
                                           
            End If

            'check if the current row is the first row with the selected tickerIndex.
            If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
            End If

            'check if the current row is the last row with the selected ticker
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then

                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

                'and increase the tickerIndex.
                tickerIndex = tickerIndex + 1

            End If

        Next i

        'activate the worksheet to receive the analysis
        Worksheets("All Stocks Analysis").Activate
        
        'Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
        For i = 0 To 11
        
            'output the values to the table
            Cells(4 + i, 1).Value = tickers(i)
            
            Cells(4 + i, 2).Value = tickerVolumes(i)
            
            Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
                       
        Next i

        'Formatting the output table
        Worksheets("All Stocks Analysis").Activate
        Range("A3:C3").Font.FontStyle = "Bold"
        Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
        Range("B4:B15").NumberFormat = "#,##0"
        Range("C4:C15").NumberFormat = "0.0%"
        Columns("B").AutoFit
        
        dataRowStart = 4
        dataRowEnd = 15

        'color the cells to show pos or neg returns
        For i = dataRowStart To dataRowEnd

            If Cells(i, 3).Value > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
            Else
            
            Cells(i, 3).Interior.Color = vbRed
            
            End If

        Next i

    endTime = Timer

    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
    
    'prints the time in the header
    Cells(4, 11).Value = (endTime - startTime)
    
    End Sub