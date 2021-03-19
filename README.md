# Stocks Analysis with VBA

## Overview of Project
This analysis examines the daily volume and return of different stock options in order to determine which is the best choice for investment.

## Results
### 2017 vs. 2018 Analysis of Stock Returns 
The return for a majority of the stocks decreased drastically between 2017 and 2018, going from a positive to a negative return. Although TERP was the only stock to have a negative return rate in 2017, it increased slightly in the year 2018. However, it remained negative for both years. ENPH, one of the stocks whose return decreased, still displays a positive overall return rate in 2018. However, as the return rate and total daily volume both decreased overall, I would caution against investing unless further information on the company Enphase Energy is provided. The only stock to display positive changes between the two years is RUN, which went from a 5.5% return rate in 2017 to a 84.0% return 2018. The total daily volume also increased. Investing in Sunrun seems like a great option, especially when investing for the first time.
### Script Refactoring
After refactoring the script, the execution time dropped dramatically, going from around 0.8 sec to around 0.1 sec. The execution times for both 2017 and 2018 after refactoring can be seen below.

![VBA_Challenge_2017](Resources/VBA_Challenge_2017.png)
![VBA_Challenge_2018](Resources/VBA_Challenge_2018.png)

## Summary
### Refactoring Code 
Although reusing and remodeling code which has already been written took more time to complete overall, it allowed for a much faster execution rate. It also allows for code that is easier to read and understand as each variable is associated to an attribute that is present in all options, rather than just focusing on one. 
### Original VBA Script vs. Refactored Script
Refactoring the original script allowed for a better comparison of the different stock options. The original script focused on one stock, while the refactored code executed the same method of analysis on all stocks. This removed any potential instance of bias towards one stock or another, allowing for numbers which accurately depict each stock as it compares to the others. This comparison can be seen in the difference in variables present in the images of code below. One displays code for only DQ while the other displays code for all stocks.

![DQ_Analysis](Resources/DQ_Analysis.png)

![All_Stocks_Analysis](Resources/All_Stocks_Analysis.png)
