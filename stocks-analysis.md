# stocks-analysis
**Creating Analysis on Green Stocks and expand dataset to include entire stock market over the last few years**

## Overview of Project
The purpose of this challenge was to refactor my previously run code to loop through all the stock data one time in order to collect the same information as previously, and determine whether refactoring the code successfully made the VBA script run faster. As I refactor this code, the idea was not to add new functionality, but to make the current functionality more efficient and easier for future users to read, but reducing steps, using less memory, or improving the logic of the current code. 

Some background on this project:

The scenario involves the character of Steve who is helping his parents decide how to diversify their investments in green energy stocks by analyzing stock data. Steve has given me an Excel file containing the stock data he wants me to analyze, and therefore I used VBA to code and analyze this data. My task was to write VBA code to automate analyses to allow Steve to reuse the code with any stock, and to reduce the chance of accidents and errors in his manipulation and view of the data. After some initial setup and running a test Macro, I wrote code to find the Total Daily Volume of DQ stocks in 2018 by using loops, nested loops, conditionals, and patterns. Then, I got DQ's Yearly Return for 2018 loops and conditional statements again. Following that, I looped over all the stock tickers in this excel document. Then I found the starting price and ending price and coded to get the output for this data to show up in a worksheet in this document. After that, I did some formatting of the AllStocksAnalysis sheet and made it more interactive for the user by inserting buttons and creating a message box to allow the user to pull up data from one specific year at a time. 

## Results
**Refactor Code and Measure Performance**

Here is a sample of the original code I used:

As I worked to refactor the previous code, I used a series of arrays and loops instead of nested loops as previously done. The goal was to make the code more efficient and to therefore run faster. As you can see in the below screenshots, the screenshot showing the results of the code without or "with no" refactor ran slower than the code with refactor. The original code before refactoring, ran in 1.22 seconds, whereas the code ran in 0.85 seconds after the code was refactored. 
 
![2018 Stocks Without Refactor](AllStocks2018NOREFACTOR.png)

![2018 Stocks WITH Refactor](resources/AllStocks2018WITHREFACTOR.png)


## Summary

Advantages or disadvantages of refactoring code in general:

Advantages:

One clear example of an advantage of refactoring the code, was the reduced time it took to run the analysis, as seen in the images above. For the user, this would certainly count as an advantage when analyzing the stocks data in this sheet. Most of the time, errors in logic appeared in a clear manner, which allowed me to fix errors as I went without having to do much digging or guesswork. Making adjustments in small pieces, only one at a time, allowed me to see which pieces were being resolved as I went. 

Disadvantages:

There are some disadvantages of refactoring code as well. If the code was not properly annotated then it could be hard to decipher the purpose of a certain line or section of code. There is also the issue of the code not working once reworked. In this project, I found that the new code did not work right away for the 2017 page as there was some problems with the data itself. While it worked in the end, it did take some time to find the cause of the issue. The data corruption was not an issue with the original code, and for our small data set, the original code did the job just fine. 

Advantages and disadvantages of the original and refactored VBA script:

The original VBA script had the advantage of being clearer to me, where I understood what each part was doing. The refactored code, although more efficient, was more challenging to get it to be successful. The refactored code worked faster, which would certainly be an advantage to the user. 