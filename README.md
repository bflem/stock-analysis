# stock-analysis

## Overview of Project
The purpose of this analysis is to help Steve make an educated, data-based recommendation on stocks for his parents

## Results
Of the stocks we analyzed, 2018 was not a great performance year for many of them. Only ENPH and RUN had positivie returns of 81.9% and 84%, respectively. However, RUN was the only one that saw an increase in return from the previous year's 5.5%. Although ENPH had a positive return percentage for both years, it declined to from 129.5% in 2017 to only 81.9% in 2018.
![VBA_Challenge_Compare](C:/Users/JDCS5731/Desktop/Analysis Projects/Stock Analysis/Resources/VBA_Challenge_Compare)
Return percentages were calculated using
Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1

In addition, the refactored script for 2017 and 2018 ran .726563 seconds and .7578125 seconds faster, respectively, than the original scripts.
![green_stocks_2017](C:/Users/JDCS5731/Desktop/Analysis Projects/Stock Analysis/Resources/green_stocks_2017)
![green_stocks_2018](C:/Users/JDCS5731/Desktop/Analysis Projects/Stock Analysis/Resources/green_stocks_2018)
![VBA_Challenge_2017](C:/Users/JDCS5731/Desktop/Analysis Projects/Stock Analysis/Resources/VBA_Challenge_2017)
![VBA_Challenge_2018](C:/Users/JDCS5731/Desktop/Analysis Projects/Stock Analysis/Resources/VBA_Challenge_2018)

## Summary
Some advantages of refactoring are the code is easier to read, programs can run faster, and the design is improved. Some disadvantages are the project could take an unintended direction after being refactored and could possibly create unnecessary bugs. I think we saw all of the advantages named in this case of refactoring. The code is easier to read. The programs run faster and has a better design. As far as disadvantages, I did encounter a few bugs while attempting to refactor the original script.
