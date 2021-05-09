##Purpose

The purpose of this project is to refactor a Microsoft Excel VBA code to collect total volume stock information in the year 2017 and 2018 and determine whether the stocks are worth investing. This process was initially executed in a similar format, but the goal is editing the previous code to improve the efficiency in the recent one.

###The Data

Data presented contains two charts with stock information on 12 different stocks. The stock information contains a ticker value, the date the stock was issued, the opening, closing and adjusted closing price, the highest and lowest price, and the volume of the stock. The aim is to retrieve the ticker, the total daily volume, and the return on each stock.

##Results

I started by copying the code needed to create the input box, chart headers and ticker array. Activate the appropriate worksheet and followed the steps listed to set the refactoring structure.
See below instruction and code as written in the file:

    '1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
    tickerVolumes(i) = 0
    tickerStartingPrices(i) = 0
    tickerEndingPrices(i) = 0
    Next i
       
    ''2b) Loop over all the rows in the spreadsheet.
    Worksheets("2017").Activate
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value      
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
         tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
         End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If  Then
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
         tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
         End If

            '3d Increase the tickerIndex.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
            End If
               
        'End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    
    Next i

According to findings, 2017 was a better year for most of the stock compared to 2018.
Despite poor stock performances for 2018, few stocks like ENPH and RUN still maintained profitability with RUN being the most appreciated and profitable stock for the year.
TERP stock continuous to experience further decline. Consequently, it is advised that the stock is taken off the books.

##Pros and Cons of Refactoring Code

Refactoring helps make our code readable, organized and easily understandable. A few advantages of an organized and clean code entails design and software improvement, debugging, and faster programming. It also benefits other users who view projects on a later date as it is easier to read, more concise and straightforward. 

The cons of refactoring code includes introduction of bugs that may complicated the code and make it hard to debug, and also it can be time consuming.

The Advantage of Refactoring Stock Analysis
The most significant advantage of the refactoring was a decrease in macro run time. The original analysis took approximately one second to run, whereas our new analysis only took about a four of the time (approximately 0.25 seconds) to run. 

