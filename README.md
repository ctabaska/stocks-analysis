# VBA Stocks Analysis

A script that analyzes a set of stocks based on price and volume using Excel and VBA

### Overview

The purpose of this script is to go through the stock market and evaluate the stock return over a full year.

### Results

Our script was running relatively slow so we needed to refactor it.

Instead of looping through through the entirety of our transactions for each stock, I changed it to go through the spreedsheet once and check which stock it is to do calculations on it.
```
For i = 2 To RowCount
    
    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value

    If (Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex)) Then
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
    End If

    
    If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
      tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
    End If

    If (Cells(i + 1, 1).Value <> Cells(i, 1).Value) Then
      tickerIndex = tickerIndex + 1
    End If

Next i
```
The next thing we do is display all our information onto a seperate sheet.

With all the information we gather it was hard to read so we need to format everything properly. 

```
For i = dataRowStart To dataRowEnd
        
    If Cells(i, 3) > 0 Then
      Cells(i, 3).Interior.Color = vbGreen
    Else
      Cells(i, 3).Interior.Color = vbRed
    End If

Next i
```

This part was important because after the script runs, there is not an easy way to understand it without looking at each datapoint individually.

After refactoring our script we were able to get about a 60% decrease in scripting time.

Here is the runtime for the 2017 spreedsheet:

![2017 run time](https://github.com/ctabaska/stocks-analysis/blob/03efa611b7d60fa8790628fdf63285bdab3473f9/Resources/VBA_Challenge_2017.png)

Here is the runtime of the 2018 spreedsheet:

![2018 run time](https://github.com/ctabaska/stocks-analysis/blob/03efa611b7d60fa8790628fdf63285bdab3473f9/Resources/VBA_Challenge_2018.png)

### Summary

Refactoring our code was important for two reasons:

#### 1. Much of our code before we refactored it was messy and hard to read.

While I try to write code as concise and readable as possible, it gets messy when you're problem solving. Refactoring code can lead to less mess and easier readability because we only look at how best to represent the ideas we already thought of.

#### 2. We may overlook certain processes that would speed up our script time.

It's easy to keep working on the process that you have thought through. When you have finished that process, it's important for you to look at other options and compare that difference.

However refactoring takes time as you're working on it when you could be working on another project or another part of the same project.

Specifically in this project the process that we started was inefficient. The massive amount of loops we had to do was slowing down our process. If we took our existing code and scaled up our data, it could take way more time then our new refactored code.
