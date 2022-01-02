# Stock Analysis

## Overview of Project

The purpose of this project is to analyze a set of stock market data to determine the total daily volume and return for twelve green energy stocks over the period of one year. The provided information for this analysis includes the daily open price, high price, low price, close price, adjusted close price, and total volume for each of the twelve stocks in 2017 and 2018.  Using Excel and Visual Basic Studios, I created a macros that automatically parses through the dataset to output performance metrics for each ticker symbol in an organized table. The table was then formatted in VBA to color code stocks that performed with positive return "green" and stocks that performed with negative return "red".

Once completing the macros, I was tasked with refactoring the code to be compatible with any arbitrary number of stocks as well as optimize the script to run faster.

## Results

The results of the analysis show that overall, the green energy stocks performed better in 2017 than 2018 with a few exceptions. In 2017, all of the stocks in this analysis performed with positive return except for 'TERP' that returned -7.2% for the year. In 2018, all the stocks performed with a negative return except for "ENPH' (+81.9%) and 'RUN' (+84.0%). The return and total daily volume are also indicators of how volatile each stock is and the associated risk level. The results are shown in the two figures below:

[insert results]

The original macros was designed to loop through the dataset twelve (12) times, once for each ticker symbol. With over 3,000 rows of data, the macros would have to loop through over 36,000 rows of data. In the refactored code, the macros only loops through the data set once, determining when a new ticker starts and ends using an if-then statement within the single loop. See the code examples below:

[insert loops]

Due to the design of the refactored code, the macros has a significantly shorter execution time. This is expected as the refactored code only loops through approxiamtely 3,000 rows of data compared to the 36,000 rows of data in the original. See the execution times below:

[insert old times and new times]


![VBA_Challenge_2017](https://user-images.githubusercontent.com/95327115/147885893-3d635078-7018-48e4-956b-e385e22ba49a.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/95327115/147885899-e3e82300-2e05-473e-8764-8e3de346835a.png)


## Summary

This project challenged me to find more efficient coding solutions to analyze the stock data. The advantage of refactoring code is greater efficiency, and shorter execution time, for complex codes. While in the process of refactoring you may also discover bug errors and/or "magic numbers" that would break the code if the dataset was slightly different. Refactoring is especially beneficial as the code becomes more complex and the processes begin to take longer. The disadvantage of refactoring code is that it can lead to the creation of unintended errors where the original code is sufficient. Additionally, refactoring takes time and may not be in ones best interest if the current code is functional and refactoring will only shave a short amount off of the execution time.

As applied to the VBA script for this project, the refactored macros helped to create a macros applicable to an arbitrary number of stocks using the 'tickerIndex' variable. It also helped to reduce runtime by a significant amount which would become even more apparent in a larger data set. During refactoring, I ran into challenges with indexing correctly into arrays and it took time to debug in order to produce the anticipated outcome. Overall, refactoring improved the quality and efficiency of the VBA script.

