# stocks-analysis

## Overview of Project
### Purpose and Background

The purpose of this project is to analyze how certain stocks performed in 2017 and 2018. We will look at the total daily volume for each stock in order to determine how actively the stock is traded. We will also look at the yearly return for each stock, which will tell us the percentage difference in price from the beginning to end of the year.

Once we have this information, we will be able to better advise Steve and his parents on which stocks are performing best (and which ones they might want to invest in).

## Results
### Stock Performance and Comparison

After calculating the total daily volume and yearly return for each stock, we formatted the table to better visualize the data. In the yearly return column, we highlighted all the postive yearly return outcomes in green. Anything else was highlighted in red. This formatting makes it easy to see that our stocks had mostly postive yearly returns in 2017, and did less well in 2018.

![Output_Stocks_2017](https://user-images.githubusercontent.com/79174885/110248998-cd5da400-7f41-11eb-86b8-3433002ed255.png)

![Output_Stocks_2018](https://user-images.githubusercontent.com/79174885/110249007-d51d4880-7f41-11eb-9e22-fe5329adb9a9.png)

Based on the output tables above, we can also see how the DQ stock performed specifically in comparison to the other "green stocks" that Steve's parents could invest in. The total daily volume for DQ in 2017 was low compared to the other green stocks, but the yearly return was actually the highest of the group. In 2018, the total daily volume for DQ was on the low end, and it had the worst yearly return of the group. Steve can show this comparison to his parents and they can make a decision on the best green stock investment strategy moving forward.

### Execution Times

The original macro we used to get the output tables (called `AllStocksAnalysis()`) ran in 1.144531 seconds for 2017 and 1.148438 seconds for 2018. After refactoring the code, the new run times were significantly faster. The analysis for 2017 took 0.1132812 seconds and 0.1210938 seconds for the 2018 analysis. We can see the run time confirmation in the message boxes that pop up at the end of the code.

![VBA_Challenge_2017](https://user-images.githubusercontent.com/79174885/110249659-28dd6100-7f45-11eb-8972-3e867c7d34f6.png)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/79174885/110249660-2bd85180-7f45-11eb-8e41-aac0a33aa996.png)

## Summary
### Advantages of Refactoring Code

Refactoring code can be advantageous in many ways. Below are some reasons you may decide to refactor your code:
- Refactored code may run faster.
- Improve the "design" of your code-- make it more elegant and streamlined.
- Refactored code may be easier to understand.
- Going through the refactoring process has the potential to highlight bugs in your code.

### Disadvantages of Refactoring Code

In some cases, however, you may not want to refactor your code. These reasons could be...
- You are facing time constraints (refactoring can be a time intensive process!).
- Depending on the complexity of the code, you could make a mistake that costs you more time and energy to fix.

### Application: Refactoring in Module 2

At first, I didn't see where there was an opportunity to refactor in our original `AllStocksAnalysis()` code. After talking more with my classmates, I was able to identify the parts of our original code that were repetitive or excessive. One instance of code that could be cleaned up was
```
If Cells(j, 1).Value = ticker Then
  totalVolume = totalVolume + Cells(j, 8).Value
End If
                
'Find starting price for the current ticker
If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
  startingPrice = Cells(j, 6).Value
End If
                
'Find ending price for the current ticker
If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
  endingPrice = Cells(j, 6).Value
End If
 ```
We checked three different times to make sure that `Cells(j,1).Value = ticker` which is not efficient. As a solution, we added a line of code to make sure that we're always going to match the cell we're in with the ticker we're adding to. We did this by creating a variable `tickerIndex` and increasing it by one every time we had reached the end of the current ticker.
```
If Cells(j + 1, 1).Value <> tickers(tickerIndex) Then
  tickerIndex = tickerIndex + 1
End If
```
We also set the tickerVolumes for every ticker to zero at the beginning of the code so we didn't have to do it again every time we started a new ticker. We used a simple for loop to acheive this.
```
For i = 0 To 11
  tickerVolumes(i) = 0
Next i
```

Refactoring the code clearly improved the run time and efficiency, so it was well worth the additional time to refactor in this instance.
