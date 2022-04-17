# stock-analysis

## Overview of Project

### Purpose
In this project we are using VBA to help us analyze the stocks we are given. There are 12 different tickers and over 3,000 rows of data. Our goal is to look at each ticker and see the total volume and return that they are getting in each year since we have two sheets from two different years. In module 1 we completed this work, we compiled a code that made it easy for Steve to look into the spread sheet, click a button and See the total Volume and total return but now our goal is to refractor the code so that it is more efficent but can also hold more tickers if necessary.

## Analysis and Challenges

The initial VBA script was created and run with the file green_stocks. Our 2017 findings show that most of the tickers had a positive return and were safe investments. The Particular one that Steve's parents were looking into was DQ which had the greatest return of all. For the 2018 you can see that the returns were not nearly as high and the total volumes were down as well. 

![2017.png](/Resources/2017.png) ![2018.png](/Resources/2018.png)

The original code only initialized startingPrice and endingPrice and had us loop through all of the values, essentially making the process longer. You can see the running time at the pictures. 

![Green Stock_2017.png](/Resources/Green%20Stock%202017.png) ![Green Stock_2018.png](/Resources/Green%20Stock%202018.png)


#### Refactoring methodology
For this challenge we refactored the code so that the code would be more efficient and run faster. One of the main things we did in this refractored code that made it different was creating the four output arrays (tickerIndex, tickerVolumes, tickerStartingPrices and tickerEndingPrices) as arrays. Having the extra array allowed us to loop through the rows, increase our tickerVolume and add the volume for the current tickerIndex only. 

The following three changes made when refactoring the script influenced the run time of the script for efficiency.

The creation of a tickerIndex variable.

 Markup : * '1a)Create a ticker Index  
    * _tickerIndex = 0_

The creation of output arrays for tickerVolumes, tickerStartingPrices, and tickerEnding Prices.
 * '1b) Create three output arrays  
    * _Dim tickerVolumes(12) As Long_
    * _Dim tickerStartingPrices(12) As Single_
    * _Dim tickerEndingPrices(12) As Single_
Ensuring that the arrays’ starting value was set to zero at the beginning of each loop.
* '2a) Create a for loop to initialize the tickerVolumes to zero.
  * _For i = 0 To 11
  * tickerVolumes(i) = 0
  * tickerStartingPrices(i) = 0
  * tickerEndingPrices(i) = 0
  * Next i_



### Challenges and Difficulties Encountered
1. Overall task wasn't too challenging, however it caused me to give it a second thought when I needed to remove Quoters from the Rows in Pivot Table at the sheet Theater Outcomes by Launch Date. 
2. At the sheet Outcomes Based on Goals one should pay attention to the formula at the last row, because it is different from another and should use >= mathematical symbol.
3. At the first step I don't even noticed, that the chart Outcomes_vs_Goals should has a goals describing on the x-axis. Only when I myself couldn't remember what the points mean, I saw that in the sample assignment the goals directly listed. 


## Results

>- What are two conclusions you can draw about the Outcomes based on Launch Date?
1. Overall Theater campaigns are most of the time two times more successful than they are failed.
2. The best months for launching Theater campaign are May and June, and the worst month is evidently December.

>- What can you conclude about the Outcomes based on Goals?

Campaigns with the small goals are usually more successful, then campaigns with the big goals: plays with the goals less than 1K have the best percentage of success ever, and when goals of plays exceed 20K the probability of success becomes less than 50%.

>- What are some limitations of this dataset?

There could be other variables that affect the final outcome. For example, the plays of some authors may be more successful than the plays of others. Or advertising campaigns applied in some cases may be more successful than in others.

>- What are some other possible tables and/or graphs that we could create?

1. Calculating and visualizing by percentage of successful/failed/canceled Theater campaigns by the time of year in the sheet "Theater Outcomes by Launch Date", could be more comprehensive than actual numbers. 
2. Considering we have not the canceled campaigns at the "Outcomes_vs_Goals" chart, maybe another type of chart could be more demonstrative for percentage of successful campaigns instead of line chart, which basically duplicates information. 
