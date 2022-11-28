# Green Stock Analysis with Excel VBA

## Project Overview

Steve a friend has asked for an analysis on green stocks for his parents to see if they should invest. To accomplish this I used Visual Basic for Applications (VBA) wihtin excel to automate going through the data and get the results for each stocks total daily volume, and yearly return. After producing an accurate result I then made the automation more efficent by refactoring my code.

### Purpose
The purpose of this project was to analyze stocks efficently through automation using VBA. After the intial automation through VBA macro was done it was clear there was a more efficent way of automating. To accomplish this I had to refactor the code, this project looks to see if that refactoring accomplished its intended purpose of making the VBA macro more efficent.

## Results

### Analysis
To make the VBA macro more efficent I had to swtich from using a nested for loop and instead use three new arrays "tickerVolumes", "tickerStartingPrices", and "tickerEndingPrices" alongside the already existing "tickers" array. Additonally a variable "tickerIndex" was created to tie all the arrays together. 

#### Refactored Code

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
        For i = 0 To 11
            tickerVolumes(i) = 0
            
        Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
        For i = 2 To RowCount
        
        '3a) Increase volume for current ticker
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
                
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
                
            'End If
            End If
        '3c) check if the current row is the last row with the selected ticker
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
                
            End If
    
            
            
            'If the next rowâ€™s ticker doesnâ€™t match, increase the tickerIndex.
            '3d Increase the tickerIndex.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
            
        'End If
            End If
    Next i
    
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + i, 1).Value = tickers(i)
        
        Cells(4 + i, 2).Value = tickerVolumes(i)
        
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
#### Original Code
    '2) Initialize array of all tickers
    Dim tickers(11) As String
    
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
    
    '3a) Intialize variables for starting price and ending price
    
    Dim startingPrice As Double
    
    Dim endingPrice As Double
    
    '3b) Activate data worksheet
    
    Sheets(yearValue).Activate
    
    '3c) Get the number of rows to loop over
    
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '4) Loop through tickers
    
    For i = 0 To 11
    
        ticker = tickers(i)
        
        totalVolume = 0
        
        '5) loop through the rows in the data
        
        Worksheets(yearValue).Activate
        
        For j = 2 To RowCount
        
            '5a) Get total volume for current ticker
            
            If Cells(j, 1).Value = ticker Then
                
                totalVolume = totalVolume + Cells(j, 8).Value
                
            End If

            '5b) get starting price for current ticker
            
            If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
        
            
            startingPrice = Cells(j, 6).Value
        
        End If
            '5c) get ending price for current ticker
            
            If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
            
            'set ending price
            endingPrice = Cells(j, 6).Value
        
        End If
    Next j
   
    '6) Output data for current ticker
   
     Worksheets("All Stocks Analysis").Activate
   
     Cells(4 + i, 1).Value = ticker
   
     Cells(4 + i, 2).Value = totalVolume
   
     Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
   
    Next i
    
### Run-Time difference Original versus Refactored

The Run-Time using the original code

![2017 Original](https://github.com/Nuh-Khan/stock-analysis/blob/567e25a80122d64e17cde00df7e3e5577a1e22ca/Resources/VBA_Challenge_2017_original.png)
![2018 Original](https://github.com/Nuh-Khan/stock-analysis/blob/567e25a80122d64e17cde00df7e3e5577a1e22ca/Resources/VBA_Challenge_2018_original.png)

The Run-Time using the refactored code

![2017 Refactored](https://github.com/Nuh-Khan/stock-analysis/blob/567e25a80122d64e17cde00df7e3e5577a1e22ca/Resources/VBA_Challenge_2017.png)
![2018 Refactored](https://github.com/Nuh-Khan/stock-analysis/blob/567e25a80122d64e17cde00df7e3e5577a1e22ca/Resources/VBA_Challenge_2018.png)


## Summary

### Thoughts on Refactoring Code Generally
When you refactor code the general purpose is to make it more efficent. Additionally it can also allow the code to be read more easily and potentially make the code apply to additional data sets. However, refactoring can take a lot of time and may not always make the code more efficent.

### The effects of refactoring this code

By refactoring this code the VBA macro was able to run much faster. For 2018 our code went from running in 1.0 seconds to .078 seconds, a dramatic decrease in run-time. For 2017 our code went from running in .589 seconds to .093 seconds, again a dramatic decrease. Not only did our new macro become more efficent it also allows for use on additional data sets with simple changes. The only negative with our refactoring was the time it took to actually refactor the code, if there was a time crunch on our project it may not be viable to spend extra time refactoring our code. Both macros work and acomplish the same task. 


