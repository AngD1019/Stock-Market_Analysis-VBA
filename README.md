# VBA Stock Market Analysis

## Overview of Project
This project exists to analyze a workbook of stocks to identify total daily volume traded to theoretically inform the quality of the stock's performance.

### Purpose
The purpose of this project is to assess the daily volume of stocks relative to stock performance/return with consideration of the VBA code used to analyze the data set over years 2017 & 2018. 
Additionally, the code will be refactored to identify any differences or changes in efficiency notable enough to implement into further stock assessments.

### Results

Using stock market data from 2017 and 2018, we refactored the VBA code in these documents by creating a tickerIndex variable with three output arrays of tickerVolumes, tickerStartingPrices and tickerEndingPrices. Then, we created a series of for loops to initialize the tickerVolumes to zero, to loop over all rows, and later to loop through all arrays. We also created if-then statements to assign current starting price and current closing price to tickerStartingPrices and tickerEndingPrices respectively. Then, wrote a script to increase the tickerIndex if the next row's ticker doesn't match the previou row's ticker. Then, when we ran the code, the following the pop-up messages appeared showing elapsed run time.

2017 Stocks Code run time:

<img width="493" alt="Year2017CodeTime" src="https://user-images.githubusercontent.com/114960958/226084076-3b6c1fce-4dfd-452b-bfe7-768926a8b086.png">

2018 Stocks Code run time:

<img width="463" alt="Year2018CodeTime" src="https://user-images.githubusercontent.com/114960958/226084087-211438ae-b21b-426f-9c1d-483944ae3591.png">

All stocks 2017

<img width="190" alt="AllStocks2017" src="https://user-images.githubusercontent.com/114960958/226084091-cc0d490b-c517-43b0-80f2-4f39aeb2a792.png">

All stocks 2018

<img width="191" alt="AllStocks2018" src="https://user-images.githubusercontent.com/114960958/226084095-82c43b3b-48a9-4584-a044-a8db7833e01a.png">

All stocks 2018 with refactor

<img width="476" alt="AllStocks2018WITHREFACTOR" src="https://user-images.githubusercontent.com/114960958/226084113-7e1f00a6-384e-4a6e-8cf1-5819016cdca6.png">

All stocks 2018 without refactor

<img width="485" alt="AllStocks2018NOREFACTOR" src="https://user-images.githubusercontent.com/114960958/226084110-05aec1dc-460d-409d-8d7b-46fb9f6be6af.png">


### Summary

Refactoring is an essential part of any project. Some specific code can be useful for first iterations, but over time refactoring code to a more efficient state 
can set up loops requiring fewer steps, use less memory, and generally improve the logic to make it easier to read and apply later.

Refactoring is relevant in this project because the standard code may be utilized for just a few stocks or perhaps in an ongoing fashion for extensive numbers of 
stocks. The original VBA structure was ideal for a smaller dataset since the code was easily referenced, checked, and amended for the data. 
However, with the published, refactored code we find faster executions of the code which implies more efficient logic.
