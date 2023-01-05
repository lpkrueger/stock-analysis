# VBA_Challenge - Module 2 Challenge

## Overview of Project

### Purpose
The project aims to produce a refactored version of the script from Module 2 solution code "AllStocksAnalysis" to run more efficiently and offer more logical clarity in the VBA code itself. This script analyses a dataset of thousands of stocks and outputs an organized list of the returns for 12 individual stocks of interest for either the year 2017 or 2018, depending on the input prompt. 

## Method
Modifications will be made to the VBA script to imrpove efficiency without changing functionality. Then, runtimes of the original and modified sripts will be compared to signify improvements in efficiency, and multiple subroutines will be combined to create one script, including the conditional formatting of output values.   

### Efficiency
Both the "tickerStartingPrices" and "tickerStartingPrices" variables were redefined as a single data type (stored as 4 bytes) instead of a double data type (stored as 8 bytes). This reduction in required memory over a large dataset results in significant runtime efficiency.  

Original code storing ticker prices as double data types: 
 ```
    Dim startingPrice As Double
    Dim endingPrice As Double
```
Refactored code storing ticker price arrays as single data types:
 ```
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
 ```
### Visual Clarity
Visual and logical clarity has been improved in the refactored script, including adding a "tickerIndex" integer variable, which makes it easier to track what is being incremented throughout the nested "for" loops and displayed in the output cells.

Refactored code defining "tickerIndex" and using it in the nested "for" loop:
 ```
    Dim tickerIndex As Integer
    tickerIndex = 0
    For tickerIndex = 0 To 11
        ticker = tickers(tickerIndex)
        tickerVolumes(tickerIndex) = 0
     Worksheets(yearValue).Activate
        For i = 2 To RowCount
        
            '3a) Increase volume for current ticker
            If Cells(i, 1).Value = ticker Then
                tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
            End If

            '3b) Check if the current row is the first row with the selected tickerIndex.
            If Cells(i - 1, 1).Value <> ticker And Cells(i, 1).Value = ticker Then
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            End If
                
            '3c) check if the current row is the last row with the selected ticker
            'If the next rowâ€™s ticker doesnâ€™t match, increase the tickerIndex.
            If Cells(i + 1, 1).Value <> ticker And Cells(i, 1).Value = ticker Then
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            End If
        
        Next i

    Next tickerIndex
 ```

## Results
The refactoring exorcise improved runtimes significantly:
- The original code from module produced runtimes of 1.755859 seconds for the year 2017 and 1.617188 seconds for the year 2018.
- The refactored code in this challenge produced runtimes of 1.208984 seconds for the year 2017 and 1.363281 seconds for the year 2018. 

Links to the visual timer outputs of the refactored code are as follows:
/blob/main/resources/VBA_Challenge_2017.png
/blob/main/resources/VBA_Challenge_2018.png

- The refactored code produced improved runtimes of 0.546875 seconds for the 2017 analysis and 0.253907 seconds for the 2018 analysis.


## Summary
The advantages of refactoring code influde improving runtimes, reducing the required memory, and cleaning up the logic or flow for easier viewing or editing by others. Disadvantages include that refactoring often takes more time to write or edit, so for simple one-off coding tasks this process may not be warranted.  
The refactored script in this exercise provided a much cleaner solution that ran faster than the original while providing conditional visual formatting to the ouptut data, and functionality to chose the desired year for analysis. 