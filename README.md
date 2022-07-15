# Stock Analysis - VBA Challenge

## Overview

### Purpose

The purpose of this project is to refactor the `AllStocksAnalysis()` macro to run more efficiently to enable the subroutine to process a larger volume of stocks in the same amount of time.

### Background

Steve, a recent finance graduate, has requested help with developing an investment portfolio for his parents. We have already analyzed a single stock and a small list of a dozen stocks, but now Steve is looking for a refactored version of the macro that can analyze many, even thousands of, rows much more efficiently than the original macro.

## Results

### Original Macro

After initializing all variables and arrays, recording the start time, and gathering input for the year of data that will be considered, the macro begins a loop through all of the tickers that were initialized in the tickers array. For each new ticker, the variable for total volume is reset to zero and a nested loop checks each row for the matching ticker ID. When a matching row is found, the total volume variable is incremented by the volume represented by that row. Starting and ending prices are found by checking if the previous and next rows match the current ticker -- if the previous row does not match, it is the first row, and the row's date can be recorded as the start date. Finally, the data is outputted.

### New Macro

The refactored version of the `AllStocksAnalysis()` macro removes the uneeded nested `for` loop which will improve the processing time greatly. Looping through all of the table rows for each ticker, as was done in the original macro, can be avoided since each row's value will only be added to a single ticker's total.

### 2017 vs 2018 Stock Performance

#### 2017 Baseline

Across the year 2017, all but one (TERP) of the tickers considered in the data saw a net **increase**. Furthermore, *DQ*, *ENPH*, *FSLR* and *SEDG* all saw a net return of over 100%. The highest total daily volume of 684,181,400 was recorded for *FSLR*.

![Table of 2017 results from refactored macro](Resources/VBA_Challenge_2017_Results.png)

#### 2018 Results

Across the year 2018, all but two (*ENPH* and *RUN*) of the tickers considered in the data saw a net **decrease**. The highest total daily volume of 607,473,500 was recorded for *ENPH*.

![Table of 2018 results from refactored macro](Resource/VBA_Challenge_2018_Results.png)

#### Code Efficiency

The 2017 data was pulled in *0.1054688* seconds, and the 2018 data was pulled in *0.109375* seconds. The 2017 data pull was completed about 3.6% faster than the 2018 data.

![Processing time to pull 2017 stocks data with refactored macro](Resource/VBA_Challenge_2017)

![Processing time to pull 2018 stocks data with refactored macro](Resources/VBA_Challenge_2018)

## Summary

### Advantages and Disadvantages of Refactoring Code

#### Advantages

Refactoring code provides an opportunity to improve a number of aspects of the code body. The main advantage to refactoring code is the ability to add, change or edit current functionality to meet new needs. When code functionality needs to change slightly, refactoring allows us to avoid re-coding things that have already been coded. Many of the remaining advantages stem from the ability to make changes to code quickly. These can include increased efficiency, better readability, more robust logging, clearer documentation, etc.

#### Disadvantages

While there are many advantages to refactoring code, there are also several friction points that must be considered. Choosing to make changes to code that has already been tested can result in additional bugs or errors being added into the code base. Additionally, the existing code might not always follow a structure that allows for simple refactoring.

### Application of Pros and Cons to Original VBA Script

For our application, the pros of refactoring outweighed the cons. The needs of our client changed, and the code was no longer able to meet those new needs. We were able to address existing inefficiencies in the macro to improve efficiency without altering the base functionality.
