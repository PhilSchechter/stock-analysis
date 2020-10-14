# stock-analysis
DataAnalytics class, module 2 exercises

# Analysis of Green Stocks
## Purpose of This Analysis
For those who want to do good while doing well, “Green” stocks can be an attractive proposition; invest in companies that are trying to save the environment while making a decent return.
However, not all Green stocks are the same. This analysis attempts to see which might be a reasonable investment, looking at
-	Trading volumes. Stocks that are heavily traded might be assumed to be subject to significant review by the market, and priced fairly relative to the market as a whole
-	Returns in 2017 and 2018. Past returns are not be a guarantee of future performance, but extreme returns in either direction would indicate that further investigation might be warranted
## Results
### Results of the analysis

 [VolumeAndReturns](/Resources/Returns17and18.PNG)

The stock that triggered this analysis was DQ. We can see in 2017, it value went up by almost 200% on fairly light trading. In 2018, it fell by about 62.6% on much heavier trading – this would certainly call for more investigation. If trading remains heavy, that would be a signal that the market views the 2018 correction favorably, and it may now be fairly priced.

ENPH may have an interesting story – that went up by a factor of 3.8 over the two year period (129% up in 2017, 82 up in 2018) with heavy trading – may be worth a further look.
A major takeaway – almost everything else did very well in 2017, poorly in 2018. It seems like this sector has a lot of potential, but a lot of volatility; may be attractive to day trading if you think you can beat the market, but otherwise, should be prepared to hold for the long haul, and not check your statements too often.

### Comments on Coding
Our first attempt at coding the full analysis was based on picking one ticker at a time, then reading through all the data (preliminary approach). This was later refined (refactored) to read through the full data set only once, and sending each entry’s data to the record for the appropriate ticker as we went.

In pseudo-code:
#### Preliminary approach
*in sub AllStocksAaalysis* 

For each **stock ticker** 
* For each line of data
	* Look at the ticker code in the data
		*	If this is the ticker from the out loop, accumulate the trading volume,
		*	If it is the first instance of the ticker, set the starting price
		*	If it is the last instance, set the ending price 
	* Next line of data
* Next Stock Ticker
#### Refactored approach
*in sub AllStocksAnalysisRefactored*

*only reading in each line of data once*



* For each **line of data**
	* See which ticker this data belongs to
	* Increment the trade volume for the appropriate ticker
	* If it is the first instance of the ticker, set the starting price
	* If it is the last instance, set the ending price
* Next line of data

The run time for the preliminary approach was about 1 second 
- [Prelim Method Timestamp](/Resources/FirstMethodTimeStamp.PNG)

The refactoring improved the run time to about 0.4 seconds 
-	[time stamp 2017](/Resources/VBA_Challenge_2017.png)
-	[time stamp 2018](/Resources/VBA_Challenge_2018.PNG)

## Summary
For this particular analysis as a one-off project, the refactoring was not necessary  - the stock analysis was completed with the initial code, and these results were used to validate the refactored code. So from that perspective, there was no need to rewrite the code.
However, if we were to expand this project to do repeated analysis on more intensive stock data (e.g. hourly prices, or every transaction within a fund, updated daily, etc) this refactored approach would make a good starting point, as the execution was 60% quicker.
This code could be further improved for reuse: e.g. the list of tickers could be generated through the macro rather than input manually (would require a ReDim of the array of macro names)
