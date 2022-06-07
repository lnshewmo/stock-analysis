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
        
        'concatenate the yearValue to store the value as String instead of a number
        'spaces buffer for correct sentence structure
        'sheet title
        Range("A1").Value = "All Stocks (" + yearValue + ")"
        
        'Create a header row
        Cells(3, 1).Value = "Ticker"
        Cells(3, 2).Value = "Total Daily Volume"
        Cells(3, 3).Value = "Return"
        
        
    'Activate data worksheet containing prices and volumes
    Worksheets(yearValue).Activate
        

       'Initialize array of all tickers
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
       
       '1a) Create a ticker Index
       tickerIndex = tickers(i)
    
       '1b) Create three output arrays  WHY DO THESE NEED TO BE ARRAYS?
        Dim tickerVolumes(12) As Long  'use (12) sets up an array but IDK why that is needed
    
        'these didn't work:
        'tickerVolumes(tickerIndex) = Cells(i, 8).Value
        'tickerVolumes(i) = Cells(tickerIndex(i), 8).Value
              
        
        Dim tickerStartingPrices(12) As Single
    
        Dim tickerEndingPrices(12) As Single
        
        Dim RowEnd As Integer
                      
                                          
       'Get the number of rows to loop over
       RowEnd = Cells(Rows.Count, "A").End(xlUp).Row
  
            ''2a) Create a for loop to initialize the tickerVolumes to zero.
            For i = 0 To 11   'for each ticker
            
           
                 tickerVolumes(i) = 0  'first reset the volume to zero
                 
                 
             
                ''2b) Loop over all the rows in the spreadsheet.
                For j = 2 To RowEnd
                                             
                            
                    '3a) Increase volume for current ticker
                    If Cells(j, 1).Value = tickerIndex Then
                    
                                          
                    'these don't work
                    'tickerVolumes = tickerVolumes + Cells(j, 8).Value
                        'the above one will return one row of data in the table and then breaks at a divide by zero error below
                    'tickerVolumes(i) = tickerVolumes(i) + Cells(j, 8).Value
                       
                    'THE EXAMPLE SHOWS THE VARIABLE AS AN ARRAY, WHY?
                    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
                                                                                     
                        
                    End If
                    
                        
                        '3b) Check if the current row is the first row with the selected tickerIndex.
                        If Cells(j - 1, 1).Value <> tickerIndex And Cells(j, 1).Value = tickerIndex Then
                        
                            tickerStartingPrice = Cells(j, 6).Value
                        
                        End If
                        
                        '3c) check if the current row is the last row with the selected ticker
                        If Cells(j + 1, 1).Value <> tickerIndex And Cells(j, 1).Value = tickerIndex Then
                        
                            tickerEndingPrice = Cells(j, 6).Value
                            
                          '3d) Increase the tickerIndex.
                          'If the next row's ticker doesn't match, increase the tickerIndex.
                            tickerIndex = tickerIndex + 1
                            
                        End If
                       
                                                 
                Next j  'loops to the next row
                                         
                
                '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
                'isn't this just closing the for loops for i and j?
                
                
                'activate the worksheet to receive the analysis
                Worksheets("All Stocks Analysis").Activate
                
                    'output the values to the table
                    Cells(4 + i, 1).Value = tickerIndex
                    Cells(4 + i, 2).Value = tickerVolumes(tickerIndex)
                    Cells(4 + i, 3).Value = tickerEndingPrices(tickerIndex) / tickerStartingPrice(tickerIndex) - 1 'getting div0error
                    
                                      
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
            
                For i = dataRowStart To dataRowEnd
                    
                    If Cells(i, 3).Value > 0 Then
                        
                        Cells(i, 3).Interior.Color = vbGreen
                        
                    Else
                    
                        Cells(i, 3).Interior.Color = vbRed
                        
                    End If
                    
                Next i
             
                endTime = Timer
                
                MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
            
            End Sub
