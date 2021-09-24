# VBA Challenge

## Overview of Project
This project compiles yearly volume and annual return for 12 stocks over 2 different datasets broken out by year.  The purpose was to accept an input variable from an enduser and read through all rows of the datasets for that year, calculate rolling totals, 
and summarize the totals for each stock.



## Results
The original All Stocks Analysis macro performed slower than our refactored code.    
Original 2017 run time
(Resources/Original_2017_runtime.png)
Refactored 2017 run time
(Resources/VBA_Challenge_2017.png)
Original 2018 run time
(Resources/Original_2018_runtime.png)
Refactored 2018 run time
(Resources/VBA_Challenge_2018.png)

The difference in performance time is due to the difference in the way we handled the output in the 2 macros.  In the original code we wrote each output row within our nested 
for loops one row at a time. 

Original code:

For i = 0 To 11
        ticker = tickers(i)
        totalVolume = 0
        Worksheets(yearValue).Activate
        For j = 2 To RowCount
            If Cells(j, 1).Value = ticker Then
                totalVolume = totalVolume + Cells(j, 8).Value
            End If
            
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                startingPrice = Cells(j, 6).Value
            End If
            
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                endingPrice = Cells(j, 6).Value
            End If
            
            
        Next j
        Worksheets("All Stocks Analysis").Activate

        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = (endingPrice / startingPrice) - 1
             
    Next i
	
With the refactored code, we stored our calculated values in arrays and output the contents of the array all at once in its own for loop.

    tickerIndex = 0
    For i = 4 To 15
        
        Cells(i, 1) = tickers(tickerIndex)
        Cells(i, 2) = tickerVolumes(tickerIndex)
        Cells(i, 3) = ((tickerEndingPrices(tickerIndex)) / (tickerStartingPrices(tickerIndex)) - 1)
        tickerIndex = tickerIndex + 1
        
    Next i


## Summary
An advantage to refactoring code is that it allows you to take another look at alternate ways of doing the same task.  In this project we were
able to speed up the run time of the code by rethinking the way we stored values and rewriting the output statements.  This created greater efficiency in the overall
code's I/O.
Although, I didn't feel I ran up against any disadvantages in this particular project with refactoring code and I can think of a few things that a coder would need to 
pay attention to.  When you refactor code it is very important to make sure you know exactly where the output is pointing to or you could risk overwriting to a sheet 
you didn't intend to.  Also, it is very important to review all declared variables and datatypes and make sure they still apply to your new code.

