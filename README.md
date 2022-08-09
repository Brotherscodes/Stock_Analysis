# Steve's Green Stock Analysis for 2017 and 2018
### The purpose of this analysis is to find the total daily volume and yearly return for each stock in 2017 and 2018.
#### Steve's parents are relying on the data results to better guide their investment decisions.

- In order to analyze a large dataset, sometimes it is necessary to refactor code to be as efficient as we can. See the refactored and original code below.


## Refactored Code:
    
    Sub AllStocksAnalysisRefactored()

    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    '1) Format the output sheet on All Stocks Analysis worksheet

    Worksheets("All Stocks Analysis").Activate
    
   
    'Create a header row
    
    
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"


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
    
    'Activate data worksheet
    
    Worksheets(yearValue).Activate
    
    
    'Get the number of rows to loop over


    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    
    '1a) Create a ticker Index
    
    Dim tickerIndex As Integer
    
    tickerIndex = 0
    
    

    '1b) Create three output arrays
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    
    For tickerIndex = 0 To 11
    

    tickerVolumes(tickerIndex) = 0
    
        
     'activate data worksheet
    
    Worksheets(yearValue).Activate
    
        
    ''2b) Loop over all the rows in the spreadsheet.
    
    
    For i = 2 To RowCount
    
    
    '3a) Increase volume for current ticker
    
    
    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        
        
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        
    'set starting price
                
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
        
    '3c) check if the current row is the last row with the selected ticker

        
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        
        
    'Ending price set

        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        
        End If
        
    'If the next row’s ticker doesn’t match, increase the tickerIndex.
         
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
         
         
    
            
    '3d Increase the tickerIndex.
            
        tickerIndex = tickerIndex + 1
        
        
        End If
        
    
    Next i
    
    
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    
    
    For i = 0 To 11
    
    'activate worksheet
        
        Worksheets("All Stocks Analysis").Activate
        
     'Ticker Row
     
        Cells(4 + i, 1).Value = tickers(i)
        
    'Sum of Volumes
        
        Cells(4 + i, 2).Value = tickerVolumes(i)
        
    'Total Return
        
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
        
    Next i
    
    Next tickerIndex
    
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

    End Sub
    
    
    
    
## Original Code:
    
    
    
    Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

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
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    

    '1b) Create three output arrays   
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero. 
    
        
    ''2b) Loop over all the rows in the spreadsheet. 
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            
            
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            
            

            '3d Increase the tickerIndex. 
            
            
        'End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        
    Next i
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

    End Sub
    
    
    
## Original vs. Refactored (time elapsed)

### Year 2017 Refactored Timer:

![VBA_Challenge_2017](https://github.com/Brotherscodes/stock-analysis/blob/main/resources/VBA_Challenge_2017.png)

### Year 2017 Original Timer:

![2017 original Runtime](https://github.com/Brotherscodes/stock-analysis/blob/main/resources/2017_Original_Runtime.png)

### Year 2018 Refactored Timer:

![VBA_Challenge_2018](https://github.com/Brotherscodes/stock-analysis/blob/main/resources/VBA_Challenge_2018.png)

### Year 2018 Original Timer:

![2018 Original Runtime](https://github.com/Brotherscodes/stock-analysis/blob/main/resources/2018_Original_Runtime.png)


### As seen above, refactoring the code allowed it to run faster and be more efficient. 



## Please reference the results below to see how each stock performed by year. 

## 2017:

![2017 Stock Return](https://github.com/Brotherscodes/stock-analysis/blob/main/resources/2017_Stock_Returns.png)

## 2018:

![2018 Stock Returns](https://github.com/Brotherscodes/stock-analysis/blob/main/resources/2018_Stock_Returns.png)


- 2017 and 2018 returned two polar different returns for these green stocks. ENPH and Run were the only two stocks that had a positive return for both years. 
- Due to the vastly opposing returns these stocks returned between the two years, either ENPH or RUN is recommended as a better stock investment as they both had positive returns.


## There are advantages and disadvantages in refactoring code.

- A common advantage is being able to take someone else's solution and apply it your problem by refactoring it and making it work in your dataset. A safe practice in refactoring code is to make a comment within your code stating where it was taken from, giving credit to its author.
- A disadvantage is when adding code from elsewhere to a working code may cause errors due to any mistakes within that code. The common practice of saving your work frequently will prohibit this from happening and allow you to start fresh from the moment your code was still working.



