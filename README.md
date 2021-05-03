# Stock Performance Analysis

## Overview of Project

### Background of data
The [green_stocks.xlsx](https://github.com/cgurbatri/stock-analysis/files/6411515/green_stocks.xlsx) excel sheet, contains data relating to the price of 12 green energy stocks at the opening and closing of the day as well as the highest and lowest price for that day. The data spans 2017 and 2018, separated by two distinct tabs in the above excel sheet.

### Project goals
The goal of the code is to output two metrics "Total Daily Volume" and the "Yearly Return" for multiple stocks in a computationally efficient manner. Daily volume is defined as the total number of shares traded throughout the day and the yearly return is the percentage difference in price from the beginning of the year to the end of the year. As a proof of concept, this analysis was inititally performed for a single stock (ticker: DQ) and the corresponding analysis can be found on the "DQ Analysis tab" in the [VBA_Challenge.xlsm.zip](https://github.com/cgurbatri/stock-analysis/files/6411538/VBA_Challenge.xlsm.zip) file. 

To provide a more comprehensive analysis of all green energy stocks, we automated this analysis to output the "Total Daily Volume" and "Yearly Return" for each of the 12 stocks provided. This code was further refactored to be computationally efficient at analyzing any number of stocks. Buttons named "Run Analaysis" and "Run Refactored Analysis" are included  on the "All Stocks Analysis" tab in the [VBA_Challenge.xlsm.zip](https://github.com/cgurbatri/stock-analysis/files/6411538/VBA_Challenge.xlsm.zip) file. Once clicked, these buttons will prompt the excel user to input a year for which they would like to analyze stocks. The code will then populate the spreadsheet with the ticker labels, the daily volume per stock and the yearly return. Lastly, a message box will appear calculating the total time is took for the code to execute the analysis. 

## Results
Figures below show that both the original (**Figures 1 and 2**) and refactored code (**Figures 3 and 4**) return the same results, but the refactored code is about 10 times faster.

**Figure 1: 2017 original analysis**
<img width="683" alt="2018_with_results" src="https://user-images.githubusercontent.com/45336910/116817281-83550100-ab33-11eb-88ab-b159841b528d.png">

**Figure 2: 2018 original analysis**
<img width="702" alt="2017_with_results_refractored" src="https://user-images.githubusercontent.com/45336910/116817286-8a7c0f00-ab33-11eb-9571-3d114b7b4e0c.png">

**Figure 3: 2017 refactored analysis**
<img width="699" alt="2018_with_results_refractored" src="https://user-images.githubusercontent.com/45336910/116817287-8cde6900-ab33-11eb-8c9d-009c31bacc4c.png">

**Figure 4: 2017 refactored analysis**
<img width="692" alt="2017_with_results" src="https://user-images.githubusercontent.com/45336910/116817264-76381200-ab33-11eb-9377-e4187ca7c852.png">



## Summary
Refactoring code is the process of editing and reusing previously written code to make it more efficient without changing its functionality. 

Advantages to refactoring code include:
* increases code readibility and quality
* increases computational efficiency

Disadvantages to refactoring code:
* can be more time-consuming 
* can introduce new and unexpected errors

Refactoring the VBA code proved to be computationally efficient, resulting in a script run time that was 10 times faster than the original code. With the creation of a tickerIndex counter, the refactored code allowed for anaylysis beyond the original 12 stocks. Though it was more time-consuing, refactoring was advantageous because it ultimately made the code more efficient and generalizable. 
