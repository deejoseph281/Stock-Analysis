# Stock-Analysis
## Overview of the Project
The purpose of this analysis was to analyze the performance of 12 green stocks using their market starting prices, market closing prices, and volumes traded in 2017 and 2018. In order to run this analysis, without manually reviewing each stock, we created a VBA code to analyze the 12 stocks, and the code allows the user to run the analysis on future stocks and years as well.

## Results
The results of this analysis was a code that can analyze the daily volume and return of each stock for the year, as well as the individual comparison of a single stock's daily volume and return. From the analysis, we can conclude that the stock with the strongest performance was ENPH and RUN and the other green stocks did not create positive returns in 2018. 

![image](https://user-images.githubusercontent.com/115019829/197040559-42d1f2f8-2deb-4003-9b37-23a9c4dc83ed.png)


### Comparable Stock Performance in 2017 and 2018
The green stocks performed significantly differently in 2017 and 2018. In 2017, only TERP yielded negative returns. In 2017, DQ and SEDG were among the strongest performers. 

![image](https://user-images.githubusercontent.com/115019829/197041361-bfda0f86-793a-42f6-9152-277388e0abfe.png)

However, in 2018, DQ was amongst the poorest performers within green stocks.

![image](https://user-images.githubusercontent.com/115019829/197041561-85bfd925-0c20-43d6-a3cf-9388180f0ecc.png)

The original run script time for 2017 and 2018 were as follows:

![image](https://user-images.githubusercontent.com/115019829/197042045-8c4c4249-5f6c-44f3-b130-c224f425fcbe.png)
![image](https://user-images.githubusercontent.com/115019829/197042059-97235be4-8b7b-4a8e-b4d1-01019762c2de.png)


The refactored run script time for 2017 and 2018 were as follows: 

![image](https://github.com/deejoseph281/Stock-Analysis/blob/main/Resources/VBA_Challenge_2017.png)
![image](https://github.com/deejoseph281/Stock-Analysis/blob/main/Resources/VBA_Challenge_2018.png)

The refactored code allows us to select the year in which we want to run the analysis, rather than a single year. The refactored code also removes any specific cell references allowing us to change the data sets for the years as we continue to run analysis year over year. We were also able to define variables that would be consistent with any future stocks analysis without having to manually review each macro code for references to specific cells, rows, or columns.

![image](https://user-images.githubusercontent.com/115019829/197043464-930769eb-3dd3-41a9-af1b-dfe82b8d3081.png)


## Summary
### The advantages and disadvantages to refactoring code
### How do the pros and cons apply to refactoring the original VBA script
