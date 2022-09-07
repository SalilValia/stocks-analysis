# stocks-analysis
 for module 2 
Overview of Project
    The purpose of this anlysis was to the us the refactor on the excel VBA code we have been working on through the modules 2 activity. In order to collect cartain data in the year 2018 and decide if it was worth investing in the stock that year. We originally done that in one format in the module work. Now we are focused in increasing and getting a more efficient original code. From the data we used in the Modules. We had a data from the year 2018 that showed us 12 different stocks. Each stock had a ticker value, the date the stock was issued, the opening, closing, and adjusted closing price, the highest and lowest price, and the volume of the stock. Our goal was the retrieve the ticker, the total daily volume, and the return on each stock. 

Analysis
    For the Analysis, I started by copying the code that was needed to perform the refactor by creating the input box, chart headers, ticker array, and activation of the worksheet that I was working in. Below was the intructions given to me and the code I used to write for each instruction. Now I wasn't sure if you wanted me to remove the codes we used in the module so I kept them in. Scroll post the YearValueAnalysis sub to see the AllStocksAnalysisRefactored sub. disregard the worksheet, ticker index and skill drill. I first thought we needed a new worksheet for the refactor and the skill drill, I was practicing. 

    '1a) Create a ticker Index
        tickerIndex = 0

    '1b) Create three output arrays   
        Dim tickerVolumes(12) As Long
        Dim tickerStartingPrices(12) As Single
        Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero. 
    'based off the USA/GB analogy of the buildings, if tickers doesnt match, increase the tickerIndex
        For i = 0 to 11
            tickerVolumes(i) = 0
            tickerStartingPrices(i)=0
            tickerEndingPrices(i) = 0
        Next i
    ''2b) Loop over all the rows in the spreadsheet. 
        For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value

        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            If Cells(i,1).Value = tickers(tickerIndex) and Cells(i-1,1).Value < > tickers(tickerIndex) Then tickerStartingPrices(tickerIndex) = Cells(i,6).Value
         End If  
  
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            If Cells(i,1).Value = tickers(tickerIndex) And Cells(i +1, 1).Value < > tickers(tickerIndex) Then tickerEndingPrices(tickerIndex) = Cells(i,6).Value 
         End If   
            

        '3d Increase the tickerIndex. 
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then tickerIndex = tickerIndex + 1

        End If
            
        'End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i

Summary 
    I noticed the codes to use for the refactoring were more oragnanzed and rrun smoothly with less debugging to deal with. It will also help others who view this code be able ot read our analysis better with this refactor. Now if we did no have the refactor code, we could not have run the application smootly with and run the test properly. By having the refactor, we were able to run the macros faster and more efficiently. Macros ran around 0.2675 seconds. The screenshot in the .png will show the macros speed. Although I did notice a typo in the module challenge question asking for two pics of the macros for just 2018. I assumed it was a typo so Only run the year 2018 and not 2017. I still feel like I will need a few more repititions to get a better understanding of VBA and the refactors. 