# VBA Stock Analysis

## Overview of Project

### Purpose

The purpose of this project is to accurately analyze a dataset consisting of stock prices and trade volume of green energy stocks. Originally, the dataset only consisted of 12 individual stocks, and used nested For loops to analyze their properties. However, if a user wanted to analyze a larger amount of stocks, the original macro would exponentially increase in runtime. Therefore, the VBA macro had to be refactored into one For loop that pulled all the required data in one sweep.

## Results

### Original Code Performance

The original code consisted of the following nested For loop which loops through the entire dataset 12 times

```
  For i = 0 to 11
        
        ticker = tickers(i)
        totalVolume = 0
   ...
        '6) Output data for current ticker
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1) = ticker
        Cells(4 + i, 2) = totalVolume
        Cells(4 + i, 3) = (endingPrice / startingPrice) - 1
    Next i
 ```

 The first loop (above) only reset the Volume variable to 0 and increased the ticker count, then output the data

```
        '5) loop through rows in the data
        Worksheets("2018").Activate
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
        
                endingPrice = Cells(j, 6).Value
        
            End If
            
        Next j
```  

Whereas the second loop, nested inside the first, retrieved one ticker's data on each iteration of the loop, before moving back out to the first loop to do it all again for the next ticker. 
        
This worked well enough for only 12 stock tickers in the data, and had the following outcomes for the "2017" and "2018" stock datasets:
 
 ![2017_before_refactoring](https://user-images.githubusercontent.com/100869713/162584332-78503001-7729-433f-a8a6-54b7e85bfa0b.png)
![2018_before_refactoring](https://user-images.githubusercontent.com/100869713/162584333-55200e3a-8b29-4a0c-87bf-a13d14403fac.png)

As seen above, each macro took greater than half a second to run through the data. If a user scaled up their dataset to, for example, 120 stocks rather than 12, it may end up taking 10x as long to execute the macro, which starts to become ridiculous if a user wanted to analyze entire chunks of the stock market. This is more than likely not the result of the nested For loop, which swept through the data 12x more than necesssary.


### Refactored Code Performance

The refactored code reduced the data-retrieval For loop to one pass through the data to get the starting price, ending price, and trade volume of each ticker.

```
''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
  ```
  
  Using the variable tickerIndex, the loop was able to differentiate each individual ticker on one pass and increase only that ticker's volume
 
 ```       
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
```
     
Then it retrieved the starting and ending price of each ticker, without needing to loop back around to the beginning of the data.

```
                
            '3d Increase the tickerIndex.
            
                tickerIndex = tickerIndex + 1
        End If
        'End If
    
    Next i
```
And finally, it increased the ticker index and returned to the beginning of the single loop, allowing the whole macro to run in only one loop.

This allowed the macro to run ***88% faster*** for the 2018 dataset and ***86% faster*** for the 2017 dataset. 

![VBA_Challenge_2017](https://user-images.githubusercontent.com/100869713/162584862-a2925f9c-4b43-45ef-bf89-4267bd71c0b6.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/100869713/162584863-cfa5572e-8744-4fc3-bb0f-7694e3ad82cc.png)

Using a larger dataset, this speed increase and reduced memory usage will be a massive boon to any end user.

## Summary

### Advantages and disadvantages of refactoring code

The advantages of refactoring code are numerous, and include cleaner-looking and easier to understand code, faster runtimes on execution, and reduced memory usage when running the code. The main disadvantage is the possibility of breaking code entirely when trying to make it cleaner. It can also take an exponentially greater amount of time if the original code is written without comments or in a confusing or obfuscated way.

### Pros and cons of refactoring Stock Analysis script

The pros of refactoring the Stock Analysis script are, obviously, a faster runtime and a cleaner and more understandable script at a glance, as well as the ability to potentially scale it up to larger and larger datasets without significantly slowing it down. The major con, though, is the lack of a way to *actually* scale up the script without editing the arrays to retrieve a variable amount of stock tickers from the dataset. This means that a user would need a basic understanding of VBA to retrieve all the tickers in a dataset and add them into the script for it to function. At present, it only looks for the 12 tickers included in the green energy stocks dataset.
