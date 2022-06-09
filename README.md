# Excel VBA Green Stocks Analysis

## Project Overview
  
For this project we were tasked with writing a VBA script which looks at the trading volume and annual return of 12 different green energy stocks in 2017 and 2018.  The data supplied for the analysis is contained in an Excel Workbook.  The client, a financial advisor, is initially interested in the performance of Daqo (DQ) based on his client's preferences, but to better inform his investment strategy he would like to compare its performance against these other 11 green companies.  His goal is to help guide the investors toward selecting a green stock with a history of high return.  An initial analysis was completed which provided the metrics for the green stocks investment strategy.  However he would now like the script to have capability to evaluate more stocks over a larger time period.  To accomplish this, the code was refactored for efficiency.

  **GOALS**
    - Review the performance of (DQ) across 2017 to 2018
    - Refactor the analysis code for efficiency and larger data set capability
  
## Results

### Stock Performance

#### 2017 vs 2018

##### 2017

Overall all these green stocks performed well for the year 2017, with one exception (TERP).  DQ specifically had the highest rate of return for the year at 199%.  The trading volume stands out as relatively low for 2017. This metric should be interpreted in by the client with the return rate to infer trends in buyer's intrests in this share.

image 2017 returns

#### 2018

A clear downward trend with significant losses for DQ (-62.6%) occurred in 2018.  Overall this was true of most of the green stocks with the exception of ENPH (81.9%) and RUN (84%).  

image 2018 return

###### Summary

ENPH performed well over both years with high trading volumes.  The client may consider recommending this stock as a candidate for his investor's portfolio.

### Refactoring the Code

The goal of the refactoring is to increase efficiency, as measured by run time.

#### Initial Analysis

The initial code written the green stocks analysis stores the tickers in an array and then using nested loops, first initializes the first ticker, sets the trading volume to 0, activates the correct worksheet `yearValue`, and then loops through each row for that ticker to assess `tickerVolumes` `tickerStartingPrices` and `tickerEndingPrices` before outputting the data. It will then loop to the next ticker and continue to completion. 

Time to run the code for 2017: 0.3789063 seconds
Time to run the code for 2018: 0.3828125 seconds

The code can be viewed [here](link to sub in resources "Sub yearValueAnalysis)

#### Refactoring
For the refactoring, a `tickerIndex` was created and the `tickerVolumes` `tickerStartingPrices` and `tickerEndingPrices` were defined as arrays to store the values until they are output in the table.  This allows us to un-nest the remaining loops. The first loop reset all `tickerVolume` for all tickers.  The next loop uses the `tickerIndex` to pick up the data for that ticker and stores it in the array variable for `tickerVolumes` `tickerStartingPrices` and `tickerEndingPrices`.  When the next row's ticker doesn't match, it will increase the tickerIndex.  The next loop outputs the stored values in table and the final loop formats the table.

Time to run the code for 2017: 0.0625 seconds
Time to run the code for 2018: 0.0664 seconds

The code can be viewed [here](link to sub in resources "Sub AllStocksAnalysisRefactored")

## Summary

### Advantages and Disadvantages

##### Refactoring Code

Refactoring code is a re-structuring of exsisting code to improve its quality without changing it's behavior.  Quality improvements may include: reduction of run time, reduction of the code size, to make it more understandable, to improve the logic, to make it more dynamic, for ease of maintenance, adding featurs, or improve it's design elements.  

This activity requires TIME and can introduce BUGS.  If there is not adequate time to test it may not be a good idea to undertake a refactoring.

##### Original vs Refactored Stock Analysis 

The original and refactored scripts for the stock analysis gave the same outputs in the same formats.  Removing the nested loops and creating variable arrays to store the outputs reduced its run time by nearly 600%.  Using an index for the `tickerIndex` will make expanding the analysis to other stocks easier by just adding the new stocks to the `tickers` array and adjusting the second value in the 'tickers` loop and the output loop.  The table will have to be formatted to accomodate the added tickers also at `dataRowEnd`.

The primary disadvantage is the TIME it takes to get the refactoring correct.  The initial script delivered the immediate need of looking at the green stocks' performance to inform the investor's question in under one second.  The refactoring of the code took an additional 15 hours of work.  The value of the enhancement should be weighed in the context of the time it takes to rewrite.  


