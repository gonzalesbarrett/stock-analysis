# stock-analysis

Module 2: VBA

# Stock Analysis with VBA

## Overview of Project

Analysis of select stock data total volume and total return in 2017 and 2018.

### Purpose

The purpose of the analysis was to assist the user’s parents in making informed financial decisions by provided the total volume of trades and total return for twelve tickers. Two years’ worth of ticker data was provided and the goal was to utilize VBA to both analyze the data and make process it so that the relevant data is readable and effectively provides insight into the tickers.

## Results
![Screenshot](https://github.com/gonzalesbarrett/stock-analysis/blob/main/Original%20Code%202017.png)
![Screenshot](https://github.com/gonzalesbarrett/stock-analysis/blob/main/Original%20Code%202018.png)

Above you can see screenshots of the original code when ran for both years. Below are screenshots of the refactored code for both years.

![Screenshot](https://github.com/gonzalesbarrett/stock-analysis/blob/main/VBA_Challenge_2017.png)
![Screenshot](https://github.com/gonzalesbarrett/stock-analysis/blob/main/VBA_Challenge_2018.png)

As you can see, the refactored code ran 6 times faster for the 2017 ticker data and over 5 times faster for the 2018 ticker data. The stock data matches which indicates that our refactor code was not only more efficient but also accurate.
The data itself shows that ten out of the twelve tickers had a worse year in 2018 then 2017. The exceptions were RUN, and TERP. RUN had a significantly better year increasing with a return over 78 basis points higher in 2018 then 2017. TERP had a bad 2017 with a negative return of 7.2% but did slightly better in 2018 by only losing 5% in value. ENPH and RUN were the only two stocks to have positive returns both years, but ENPH had a significantly better 2017 then 2018. Even though 2018 was a tough year most stocks only partially erased their 2017 gains. AY, CSIQ, DQ, FSLR, HASI, SEDG, and VSLR all gained more in 2017 then they lost in 2018. With the exception of TERP, JKS, and SPWR all of the other stocks would be considered mid-term gainers and show the potential to still have long-term gains.

## Results
1.	What are the advantages or disadvantages of refactoring code?

The advantage of refactoring a code is making it more efficient and less time consuming. This not only saves time but also reduces computer processing requirements which can sometimes bog down your computer. Also, since you are editing it if you make a copy of the original code, it is easy to verify that your data is accurate since you can run both versions of the code and compare results. The primary disadvantage would be that is difficult to refactor a code especially if you were not the original person to create the code or it has been awhile since you wrote the code. Without good notes it can be challenging to figure out what’s going on let along figuring out a way to make it better.

2.	How do these pros and cons apply to refactoring the original VBA script?

The primary advantage of refactoring the code was decreased run time. If this program were expand and more tickers were added the reduction in time spent processing the data could easily add up to a significant amount of time. In addition, the refactored code offered more features. We originally had a separate program to edit the cells and add conditional formatting but we were able to wrap that into the refactored code and still reduce time. The main disadvantage of refactoring this code was the troubleshooting. I had an extremely difficult time refactoring the code and getting it to work. Once I started fresh and worked through the program again I was able to find an adequate solution which I then used to figure out what was wrong with my original code. I feel like simple mistakes are easier to miss when refactoring a code since the format may not be a memorable since I did not write it.
