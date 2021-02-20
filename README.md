# Stock-analysis


## Overview of Project

### Purpose of Analysis
The purpose of this challenge was to edit, (or "refactor") the original VBA script that was created to collect and analysis specific data from various stocks for the years 2017 and 2018 years. The original VBA script works great for the small dataset of only a dozen stocks, but the script might not work as well if it is used on a much larger data set with thousands of stocks. If the script indeed does work properly for a bigger data set, it still needs to be refactored so that it can reduce the time it takes to execute the code. Refactoring the code can be done by reducing the number of lines of code, using less memory, or by improving the logic of the code so any future user can easily read it.

## Results

Analysis of AllStocks (2017):

![2017results](https://user-images.githubusercontent.com/75760493/105090025-81bf6b80-5a63-11eb-9aa9-d0a05d090a5f.PNG)
t
From the results from the image above, it shows only 1 of the stocks (TERP) had a negative return of 7.2% throughout the whole year. While the other 11 stocks all had positive returns, ranging from 5.5% for RUN all the way to DQ which had a 199.4% return. It shows that 2017 was very beneficial for most of the stocks that were analyzed.

Analysis of AllStocks (2018):

![2018results](https://user-images.githubusercontent.com/75760493/105090281-d7941380-5a63-11eb-9523-ad6aad266829.PNG)

However, from the results from the image above, 2018 was not as beneficial as 2017. Most of the stocks ended up having a negative return with only 2 stocks (ENPH and RUN) had a positive return. Even though there were only two stocks that had a positive return, their returns reached as high as 81.9%, so investing in these stocks would've been very beneficial to the investor. 10 out of the 12 stocks that were analyzed yielded negative yearly returns, ranging from -3.5% to -62.6%. The main stock of interest was the "DQ" stock, which had an amazing yearly return of almost 200% in 2017. However, it took a big hit in 2018 by having a yearly return of -63%.


Code refactoring was a major part of this project. The initial analysis was written using a nested for loop - an iterative process within which multiple additional iterative processes are contained. An example of the code is shown below.

Refactored Code:

    '1a) Create a ticker Index

    tickerIndex = 0
    

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    
    '2a) Create a for loop to initialize the tickerVolumes to zero.
    
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
            
        'If  Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
        'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If Then
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        
        End If
        
            '3d Increase the tickerIndex.
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
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
        
 This refactored code above was altered from the original code by creating 4 new arrays: tickers, tickerVolumes, tickerStartingPrice, and tickerEndingPrice. Also, a new variable was created, tickerIndex, which is used to assign the 3 arrays (tickerVolumes, tickerStartingPrice, tickerEndingPrice) to each ticker value from ticker(0) to ticker(11). This will cause the code to run faster since now it only needs to access each row of data once, as opposed to the original code which had to access each 
piece of data for each possible ticker.

Below are the run times for both 2017 and 2018 while using the original code and while using the refactored code.

Original Code:

![old_time_2017](https://user-images.githubusercontent.com/75760493/105074449-fd162280-5a4d-11eb-86eb-9d9b50b392a5.PNG)

![old_time_2018](https://user-images.githubusercontent.com/75760493/105074565-259e1c80-5a4e-11eb-9626-9fef607359de.PNG)


Refactored Code:

![VBA_Challenge_2017](https://user-images.githubusercontent.com/75760493/105075354-47e46a00-5a4f-11eb-9819-717f04a9676f.PNG)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/75760493/105075390-529eff00-5a4f-11eb-9e68-85a86a1fc22f.PNG)


From the images above, the refactor code showed to be successful by reducing the time it took to run the code for 2017 from .8671875 seconds to .1640625 seconds, reducing the time by .703125 seconds. Likewise, with the year 2018, the refactored code reduced its execution time from .8496094 seconds to .1738281 seconds, reducing the time by .6757813 seconds. 


## Summary

### Disadvantages of refactoring the code:

One big disadvantage in trying to refactor this code is that you are altering and changing a code that already works and executes the way it needs to. Refactoring can easily lead to the code no longer to be able to execute or execute properly and give the desired results if it isn't done correctly.

### Advantages of refactoring the code:

The major advantage is the new refactored code is cleaner, the code is more organized, the code takes less memory to run and would then be able to run much faster which is needed to run a code over a large set of data. Plus having a more organized code would be beneficial in terms of trying to debug the code if there are errors. Not only that, but a cleaner and more organized code would help better explain what the code is trying to do in case someone who hasn't seen the code has to use it or view it. 

### Advantages and disadvantages of using the refactored VBA script

Like stated above, the biggest advantage of using the refactored VBA script is that it runs much faster and can handle very large sets of data. The disadvantage of trying to refactor the VBA script is that the code can get difficult with the numerous loops within one another, which could easily cause a syntax error that might be hard to find while debugging, which would make it difficult to ensure the code executes properly.

