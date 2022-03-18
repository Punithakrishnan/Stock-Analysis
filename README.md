# Stock-Analysis
2.2.4
## Project Overview
This Project is to help which is the best investment for Steve's Parents. Investment they are interested in was DQ Analysis. Steve want to make sure that his parents are making the correct decision in stock market.The previous code was executed on Green stocks. For this particualr project we are refactoring the  code to analyse only 2018 and 2017 stocks. 
## Result
In 2018, only ENPH and RUN two stocks had positive yearly Return as well as large Total Daily Volume. Both of them was outperformance than others green stocks.
------
<img width="250" alt="Screen Shot 2022-03-16 at 1 19 11 AM" src="https://user-images.githubusercontent.com/98849217/159086231-f19744a3-19e1-429a-8a11-8da2bc291bfd.png">
-----
In 2017, all of stocks had positive Return except TERP (-7.2%). "DQ" made best yearly return with 199.4% but with lowest total Daily Volume (35,796,200) in 2017.
---

<img width="248" alt="Screen Shot 2022-03-16 at 1 23 04 AM" src="https://user-images.githubusercontent.com/98849217/159086167-b93d93d6-ccf9-4eb0-86a1-4756d72e4083.png"> 
-----
1 Main Loop for all data and assigned tickerIndex for 12 stocks respectively.
2 Nested loop in the main loop, executed stocks original data and retrieve ticker name, startingPrices and endingPrices, and save information to each related tickerIndex.
3 nested loop in 2nd loop, in order to get volume information for each Index.
4 new loop for putting all saved output information into an analysis sheet.
Refactoring the code is very efficent in terms of time.
