# Stock Performance Analysis using VBA


## Overview of Project

### Purpose

The purpose of this project is to help our client, Scott, evaluate potential investment opportunities by automating the digestion of stock market data. Using Excel & VBA macros we can create a worksheet that allows the user to easily select the Year they would like to focus on and output a results table populated with the ticker, the total yearly volume, and the return for the year. 

Utilizing this framework we can provide our client with a way of evaluating stock data in mere seconds to help them make a more informed decision about their investment ideas. 

During the project we refactored our original code to make the program run even faster - theoretically, we can adjust our code to accept even more stock market data from more tickers without worrying about locking up or freezing our computer due to 'clunky' code. 

The VBA code & its resulting output can be found by downloading the ![file found here](https://github.com/bbinvt/stock-analysis/blob/28c1ffc991bd99d7d47253905d381d71e3076483/VBA_Challenge.xlsm)

## Results

### Output of code

The two images below show the results output of both the original and refactored code. Lucky for us they match!

2018 Stock Analysis (original code)

![image](https://github.com/bbinvt/stock-analysis/blob/28c1ffc991bd99d7d47253905d381d71e3076483/VBA_Challenge_2018_Results.png)


2018 Stock Analysis (re-factored code)

![image](https://github.com/bbinvt/stock-analysis/blob/28c1ffc991bd99d7d47253905d381d71e3076483/VBA_Challenge_2018_Results_Refactored.png)

### Runtimes of code

The difference between the run times of the two codes is where the real change lies. The screenshots below show the respective runtimes.

2018 Stock Analysis Runtime (original code)

![image](https://github.com/bbinvt/stock-analysis/blob/28c1ffc991bd99d7d47253905d381d71e3076483/VBA_Challenge_2018_RunTime.png)

2018 Stock Analysis Runtime (re-factored code)

![image](https://github.com/bbinvt/stock-analysis/blob/28c1ffc991bd99d7d47253905d381d71e3076483/VBA_Challenge_2018_RunTime_Refactored.png)

## Analysis of Results

We can start with the obvious - the two different versions of the code return the same outputs. This is exactly what we want as it means the calculations & order of operations are correct within each of our code bodies. 

The real difference is found in the performance of the two codes. The original code ran in 1.19 second, while the refactored code ran in 0.12 second - just about a tenth of the time! This is a huge time savings especially if our data set was much, much larger. 

This is one advantage of refactoring code - the body of code can be reworked to increase the computational effeciency while maintaining the desired output criteria. One can imagine if the data set were much, much larger it could be possible to crash the computer if the code is too 'clunky' or is taking too much compute / memory to process the data. Refactoring the code helps aleviate those concerns. 

One of the distinct disadvantages of refactoring the VBA code, at least for me at this time, is thinking through the order of operations. In the original code there is a specific flow to the If, Then statements and reading the code it is easy to follow. Each of the calculations is performed within the same For loop - it makes it easy to conceptually think through what is happening. The refactored code in my eyes is a bit more complicated because each of the 'chunk' is performed separately: first initialize, then run all of the calculations, then print all values, then format the output - each within its own for loop. For me I think it will take some effort to be able to 'see' how to break apart code & refactor to increase the performance. 
