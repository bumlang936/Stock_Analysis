# Stock-analysis


## Overview of Project

### Purpose of Analysis
The purpose of this challenge was to edit, (or "refactor") the original VBA script that was created to collect and analysis specific data from various stocks for the years 2017 and 2018 years. The original VBA script works great for the small dataset of only a dozen stocks, but the script might not work as well if it's used on a much larger data set with thousands of stocks. If the script indeed does work properly for a bigger data set, it still needs to be refactored so that it can reduce the time it takes to excecute the code. Refactoring the code can be done by reducing the amount of lines of code, using less memory, or by improving the logic of the code so any future user can easily read it.

## Results

Analysis of AllStocks (2017):

![2017results](https://user-images.githubusercontent.com/75760493/105090025-81bf6b80-5a63-11eb-9aa9-d0a05d090a5f.PNG)

From the results from the image above, it shows only 1 of the stocks (TERP) had a negative return of 7.2% throughout the whole year. While the other 11 stocks all had positive returns, ranging from 5.5% for RUN all the way to DQ which had a 199.4% return. It shows that 2017 was very beneficial for most of the stocks that were analyzed.

Analysis of AllStocks (2018):

![2018results](https://user-images.githubusercontent.com/75760493/105090281-d7941380-5a63-11eb-9523-ad6aad266829.PNG)

However, from the results from the image above, 2018 was not as benefical as 2017. Most of the stocks ended up having a negative return with only 2 stocks (ENPH and RUN) had a positive return. Even though there were only two stocks that had a positive return, their returns reached as high as 81.9%, so investing in these stocks would've been very beneficial to the investor. 10 out of the 12 stocks that were analyzed yielded negative yearly returns, ranging from -3.5% to -62.6%.


Code refactoring was a major part of this project. The initial analysis was written using a nested for loop - an iterative process within which multiple additional iterative processes are contained. An example of the code is shown below

Code:

    '1a) Create a ticker Index
    tickerIndex = 0
    
    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    '2a) Create a for loop to initialize the tickerVolumes to zero
    For i = 0 To 11
      tickerVolumes(i) = 0
      tickerStartingPrices(i) = 0
      tickerEndingPrices(i) = 0
    Next i
    
    '2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
      '3a) Increase volume for current ticker
      tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value

      '3b) Check if the current row is the first row with the selected tickerIndex.
      If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then     
      'If Then
      tickerStartingPrices(tickerIndex) = Cells(i, 6).Value     
      'End If
      End If
        
      '3c) check if the current row is the last row with the selected ticker
      'If the next row’s ticker doesn’t match, increase the tickerIndex. 
      'If Then
       If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

      '3d Increase the tickerIndex.
      If Cells(i + 1, 1).Value <> tickers(tickerIndex) AND Cells(i, 1).Value = tickers(tickerIndex) Then
        tickerIndex = tickerIndex + 1 
      'End If
      End If
             
    End If

    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
      Worksheets("All Stocks Analysis").Activate
      Cells(4 + i, 1).Value = tickers(i)
      Cells(4 + i, 2).Value = tickerVolumes(i)
      Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
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

Below are the run times for both 2017 and 2018 while using the original code and while using the refactored code.

Original Code:

![old_time_2017](https://user-images.githubusercontent.com/75760493/105074449-fd162280-5a4d-11eb-86eb-9d9b50b392a5.PNG)

![old_time_2018](https://user-images.githubusercontent.com/75760493/105074565-259e1c80-5a4e-11eb-9626-9fef607359de.PNG)


Refactored Code:

![VBA_Challenge_2017](https://user-images.githubusercontent.com/75760493/105075354-47e46a00-5a4f-11eb-9819-717f04a9676f.PNG)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/75760493/105075390-529eff00-5a4f-11eb-9e68-85a86a1fc22f.PNG)



## Summary



