# VBA of Wall Street

## Overview of Project

Steve is a recent finance graduate and he wants to apply what he has learned to help his parents analyze different green energy stocks for investment to diversify their portfolio. We went through a list of green energy stocks and analyzed them with Excel's Visual Basic Analysis (VBA) - Microsoft's programming language for Office Applications.

### Purpose

Using VBA's functions, we want to build an application that automates the analyses performed on the green energy stocks. We want to figure out their historical performance in a quick and efficient manner. There are many ways to build the application, so we have refactored the orignal code of the application so that the analysis is more efficient and runs faster. 

## Analysis 
The data set used for analysis consisted of 12 different ticker stocks and their price data for each trading day in the years 2017 and 2018. We utilized VBA to calculate the Total Daily Volume of each stock and the Yearly Return for these two years. We utilized various tools in the code such as variables, arrays, for & nested loops, and conditional statements. To reduce the execution time of the script and make it smoother for larger data sets, we refactored the code to include four arrays and concise conditional statements. To make the worksheet interacative, we created buttons to allow the user to have a simple interface to perform the analysis. 

### Refactored Code Screenshots

![Refactored VBA Code - Part 1](https://github.com/anandohrid/stock-analysis/blob/main/Resources/Refactored_Code_1.png)
![Refactored VBA Code - Part 2](https://github.com/anandohrid/stock-analysis/blob/main/Resources/Refactored_Code_2.png)
![Refactored VBA Code - Part 3](https://github.com/anandohrid/stock-analysis/blob/main/Resources/Refactored_Code_3.png)
![Refactored VBA Code - Part 4](https://github.com/anandohrid/stock-analysis/blob/main/Resources/Refactored_Code_4.png)


## Results

As expected, the results from the analysis shows substantial yearly variance in performance from 2017 to 2018. 2017 was a good year for these green stocks as 11/12 or 92% of them had positive yearly returns, while 10/12 or 83% had negative returns in 2018.

The only negative performing stock in both years was TERP, with -7.2% in 2017 and -5% in 2018. DQ, ENPH, and SEDG had the highest returns. DQ had the highest yearly return out of all, but shows the least daily volume. In 2018, the only ones with positive returns are ENPH and RUN. ENPH even had the highest total daily volume out of all stocks. If I had to suggest a stock to Steve's parents, it would be to invest in ENPH, which showed the best progress for both years.

The execution times of the refactored VBA macros compared to the original scripts differed substantially. Without the code being refactored, the execution times for both 2017 and 2018 were 1.09 and 1.05 seconds, respectively. On the other hand, after being refactored, the times to the execute the code for 2017 and 2018 were 0.20 and 0.19 seconds, respectively. Clearly, the refactored macro produced much faster results.

### VBA Code Execution Time

##### 2017 All Stocks Analysis - Original Script
![Original Execution Time for 2017](https://github.com/anandohrid/stock-analysis/blob/main/Resources/VBA_Challenge_2017_Original.png)

##### 2018 All Stocks Analysis - Original Script
![Original Execution Time for 2018](https://github.com/anandohrid/stock-analysis/blob/main/Resources/VBA_Challenge_2018_Original.png)

##### 2017 All Stocks Analysis - Refactored Script
![Refactored Execution Time for 2017](https://github.com/anandohrid/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

##### 2018 All Stocks Analysis - Refactored Script
![Refactored Execution Time for 2018](https://github.com/anandohrid/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

## Summary

### Advantages and Disadvantages of Refacturing Code

An advantage to refactoring any general code is the script will be more concise, organized, and easier to follow. From what we saw in the analysis, refactored code tends to reduce the time to execute the whole code, which is ideal for huge datasets.

A disadvantage is that it can be risky to change any section of an already set code. It may introduce bugs and may be difficult to interpret for beginner coders.

### Advantages and Disadvantages of Original & Refactored VBA Scripts

I believe there are always advantages when refactoring than without. In this specific case, the refactored VBA script versus the original has clear positive results. The observable outputs were exactly the same, however, it made the code flexible and removed any needless complexity. At a small scale, the execution times were greatly reduced after refactoring, which is important when dealing with much bigger datasets.

Refactoring the original script can always be risky. In my case, I ended up debugging a few times because certain subroutines and statements were mismatched and easily gone unnoticed. With a complex code, one really has to decide if it's really worth risking breaking the code by trying to make it better. It worked for this challenge but may not be useful for data sets of other sizes and categories. The risk of complicating the code was a disadvantage on its own.
# stock-analysis
