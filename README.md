# Stock Analysis

## Overview of Project

The purpose of this project is to analyze a set of stock market data to determine the total daily volume and return for twelve green energy stocks over the period of one year. The provided information for this analysis includes the daily open price, high price, low price, close price, adjusted close price, and total volume for each of the twelve stocks in 2017 and 2018.  Using Excel and Visual Basic Studios, I created a macros that automatically parses through the dataset to output performance metrics for each ticker symbol in an organized table. The table was then formatted in VBA to color code stocks that performed with positive return "green" and stocks that performed with negative return "red".

Once completing the macros, I was tasked with refactoring the code to be compatible with any arbitrary number of stocks as well as optimize the script to run faster.

## Results

The results of the analysis show that overall, the green energy stocks performed better in 2017 than 2018 with a few exceptions. In 2017, all of the stocks in this analysis performed with positive return except for 'TERP' that returned -7.2% for the year. In 2018, all the stocks performed with a negative return except for "ENPH' (+81.9%) and 'RUN' (+84.0%). The return and total daily volume are also indicators of how volatile each stock is and the associated risk level. The results are shown in the two figures below:

|<img width="370" alt="Stock_Analysis_2017" src="https://user-images.githubusercontent.com/95327115/147886200-8afa0b50-c069-4855-b17f-0acb142e2940.png">|<img width="368" alt="Stock_Analysis_2018" src="https://user-images.githubusercontent.com/95327115/147886288-acbaa549-57a3-4939-ad19-62eca0d76943.png">|
|:--:|:--:|
| *Stock Analysis Results 2017* | *Stock Analysis Results 2018* |

The original macros was designed to loop through the dataset twelve (12) times, once for each ticker symbol. With over 3,000 rows of data, the macros would have to loop through over 36,000 rows of data. In the refactored code, the macros only loops through the data set once, determining when a new ticker starts and ends using an if-then statement within the single loop. See the code examples below:

|<img width="483" alt="Original_VBS_Code" src="https://user-images.githubusercontent.com/95327115/147886303-274432d1-7f65-47e0-8ef2-6e0d10c2af06.png">|<img width="587" alt="Refactored_VBS_Code" src="https://user-images.githubusercontent.com/95327115/147886307-1088e57f-8fb4-4213-8dc0-7699a508f82b.png">|
|:--:|:--:| 
| *Original VBS Code* | *Refactored VBS Code* |

Due to the design of the refactored code, the macros has a significantly shorter execution time. This is expected as the refactored code only loops through approxiamtely 3,000 rows of data compared to the 36,000 rows of data in the original. See the execution times below:

|<img width="273" alt="Original_Runtime_2017" src="https://user-images.githubusercontent.com/95327115/147886313-522b594f-bdf4-4083-93e4-00c97d9f4087.png">|<img width="271" alt="Original_Runtime_2018" src="https://user-images.githubusercontent.com/95327115/147886319-421a87ce-debd-459a-8e94-49298644b487.png">|
|:--:|:--:| 
| *Original Execution Time 2017* | *Original Execution Time 2018* |

|![VBA_Challenge_2017](https://user-images.githubusercontent.com/95327115/147885893-3d635078-7018-48e4-956b-e385e22ba49a.png)|![VBA_Challenge_2018](https://user-images.githubusercontent.com/95327115/147885899-e3e82300-2e05-473e-8764-8e3de346835a.png)|
|:--:|:--:| 
| *Refactored Execution Time 2017* | *Refactored Execution Time 2018* |

## Summary

This project challenged me to find more efficient coding solutions to analyze the stock data. The advantage of refactoring code is greater efficiency, and shorter execution time, for complex codes. While in the process of refactoring you may also discover bug errors and/or "magic numbers" that would break the code if the dataset was slightly different. Refactoring is especially beneficial as the code becomes more complex and the processes begin to take longer. The disadvantage of refactoring code is that it can lead to the creation of unintended errors where the original code is sufficient. Additionally, refactoring takes time and may not be in ones best interest if the current code is functional and refactoring will only shave a short amount off of the execution time.

As applied to the VBA script for this project, the refactored macros helped to create a macros applicable to an arbitrary number of stocks using the 'tickerIndex' variable. It also helped to reduce runtime by a significant amount which would become even more apparent in a larger data set. During refactoring, I ran into challenges with indexing correctly into arrays and it took time to debug in order to produce the anticipated outcome. Overall, refactoring improved the quality and efficiency of the VBA script.

