# Stocks-VBA-Challenge
## Overview of Project 

  Steve wanted to be able to analyze and compare stocks from different years to see which options would be better investment oppurtunities for his parents. So I used VBA to write some code to be able to Analyze the stocks from the years 2017, and 2018. After the program finishes running, the run time is displayed in a pop up window. 

## Results 

  If we look at the performance of the stocks from the years 2017, and 2018, it is clear to see that 2017 was a better year overall when talking about the return on investment. We can also see that the trading volume in 2017 was significantly larger tha the trading volume for the year 2018, which helps us conclude that although 2017 was a good year overall, maybe its better to enter the market in 2018 while the market is bearish, to be able to capitalize on growth oppurtunities. If Steves parents entered late in 2017 they most liekly would've lost money, since the market dropped from 2017 going into 2018. 
  
### Code Run Times

    Since there are not that many stocks in the data set the run times are pretty fast and are able to run pretty consistantly. 
  
  The run time for the 2017 stocks seems to average around 5.52 seconds, 
 ![Run Time](https://github.com/lrngdtascinc/Stocks-VBA-Challenge/blob/27c3cfb40f1cf57f5bbe87d4d8c144aace5e20b0/VBA_Challenge_2017.png.png) 
  
  While the run time for the 2018 stocks seem to average out around 5.54 seconds.
  ![Run Time 2](https://github.com/lrngdtascinc/Stocks-VBA-Challenge/blob/813137fe652f859d030cd81afbe003404cf24254/VBA_Challenge_2018.png.png)

## Summary

  Overall the we can see that the year 2017 was a good year if you were entered in the market before hand, but we determined that entering during 2017 might have been a mistake due to the drops going into 2018. As for the code, it was a little challenging having to figure out a couple of errors, notably a subscribt out of range error, which I fixed by finding and correcting a typo. The other error I ran into was a overflow error, which I solved by changing the variables startingPrice, and endingPrice to Double instead of single.  
