# stock-analysis

## Overview of Project

### Purpose
In this project we are using VBA to help us analyze the stocks we are given. There are 12 different tickers and over 3,000 rows of data. Our goal is to look at each ticker and see the total volume and return that they are getting in each year since we have two sheets from two different years. In module 1 we completed this work, we compiled a code that made it easy for Steve to look into the spread sheet, click a button and See the total Volume and total return but now our goal is to refractor the code so that it is more efficent but can also hold more tickers if necessary.

## Results

The initial VBA script was created and run with the file green_stocks. Our 2017 findings show that most of the tickers had a positive return and were safe investments. The Particular one that Steve's parents were looking into was DQ which had the greatest return of all. For the 2018 you can see that the returns were not nearly as high and the total volumes were down as well. 

![2017.png](/Resources/2017.png) ![2018.png](/Resources/2018.png)

The original code only initialized startingPrice and endingPrice and had us loop through all of the values, essentially making the process longer. You can see the running time at the pictures. 

![Green Stock_2017.png](/Resources/Green%20Stock%202017.png) ![Green Stock_2018.png](/Resources/Green%20Stock%202018.png)


#### Refactoring methodology
For this challenge we refactored the code so that the code would be more efficient and run faster. One of the main things we did in this refractored code that made it different was creating the four output arrays (tickerIndex, tickerVolumes, tickerStartingPrices and tickerEndingPrices) as arrays. Having the extra array allowed us to loop through the rows, increase our tickerVolume and add the volume for the current tickerIndex only. 

The following three changes made when refactoring the script influenced the run time of the script for efficiency.

The creation of a tickerIndex variable.

* '1a)Create a ticker Index  
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
There was appeared Overflow mistake at the tickerEndingPrices variable. The cause was on the step 3a wasn't direct instruction to set the tickerEndingPrices(i) = 0 before the starting the code. After changing that the code works. 

### Refactoring results

The file VBA_Challenge was used for to refator the VBA script. After refactoring the script, stock data for 2017 ran in 0.1796875 seconds, and stock data for 2018 ran in 0.1796875 seconds. Images VBA_Challenge_2017 and VBA_Challenge_2018 show the reduced run times. The result wasn't changed.

![VBA_Challenge_2017.png](/Resources/VBA_Challenge_2017.png) ![VBA_Challenge_2018.png](/Resources/VBA_Challenge_2018.png)

The refractored code had a much quicker run time because of the changes. This is a very positive outcome and means the code is much more efficient and most likely will run better if we have more tickers.

## Summary

### Advantages of refactoring code
Refactoring will start with a preexisting outline of the script, and the code can be used with the modules already in place. The refractored code had a much quicker run time because of the changes. This is a very positive outcome and means the code is much more efficient and most likely will run better if we have more tickers.

### Disadvantages of refactoring code
There is needed complit understanding of the VBA syntax and logic to make the script more efficient.

### Pros and cons apply to refactoring the original VBA script
There was a significant improvement in the run time of the script after refactoring. The syntax’s exact requirements meant that the order I edited the script could return an error before I finished the edit.

