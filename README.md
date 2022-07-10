# Assignment-2 Stock-analysis
Stock Analysis assignment
## Overview of Project: Explain the purpose of this analysis
### Purpose pf Project
- Refactor VBA codes of the Stock performance over 2017 and 2018, to compare results and make sure the code written is short, easy, and efficient.
### Background of Project
> Steve loves the workbook you prepared for him. At the click of a button, he can analyze an entire dataset. Now, to do a little more research for his parents, he wants to expand the dataset to include the entire stock market over the last few years. Although your code works well for a dozen stocks, it might not work as well for thousands of stocks. And if it does, it may take a long time to execute.
- The original VBA code written for Steve can only extract certain datasets, to expand the dataset to include the whole stock market data, we need to insert a button in the sheet that can easily filter a large data set.
## Results of Stock Analysis
### 2017 Performance
- Stocks performed well in 2017, 11 Tickers had positives returns, 4 Tickers showed positive return over 100%, only TERP had negative return of -7.2% shown in red. The execution only took 0.6823125 seconds. 
![2017 Stocks](/VBA_Challenge_2017.png)
### 2018 Performance
- The performance in 2018 is poor, only 2 stocks ENPH and RUN had positive returns over 80%, the 10 other stocks' returns are all negative. This indicates the same stocks are experiencing downward trends in the market. The execution time of the refactored script only took 0.6367188 seconds, which is also super fast.
![2018 Stocks](/VBA_Challenge_2018.png)
### Run time comparison
- In the original code, 'yearValue' was not used, the year was manually entered in the VBA code, but all other codes are very similar. The 2017 run time took 12 seconds, but the 2018 run time was only 0.61 seconds. By changing the 'yearValue' shortened the extraction time.
  - ![/Orig 2017.png](https://github.com/siqiou/assignment-2---stock-analysis/blob/5b10b95893b8c5f88564f1133c74e62ae7f07077/Orig%202017.png)
  - ![/orig 2018.png](https://github.com/siqiou/assignment-2---stock-analysis/blob/595f68817668b7804110f9ea4713bcf74e75b520/orig%202018.png)
## Summary
1. Advantages and Disadvantages of refactoring codes in general
-- Advantages: The run time can be shortened, take less memory, works well with older computers.
-- Disadvantages: It can be very time consuming, if there is a tight deadline of the project, programmers may not be able to finish the editing on time.
It is recommended to consider the available time for editing, if the run time of original code is not too slow, e.g. less than 10 seconds, it would be fine leave as is.
2. Advantages and Disadvantages of the original and refactored VBA script
-- Advantages: The coding file is more clear, it would help the programmer to make further changes. It is written in a way that end users can extract data by click of a button, while manual input was required in the original script.
-- Disadvantages: It did took at least half an hour to 1 hour to finalize the editing, and the run time was approximately the same. If this was a longer code, it might be worth the work. 
Clean and simple codes are the most ideal, if the project is top priority or needs focused attention to detail and contains large datasets, it is worth refactoring. So this really depends on the need of the job and time contraints in the real world situation.
