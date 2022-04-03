# VBA Stock Analysis

## Overview

The purpose of this project was to analyze stocks for Steve's parents. The original code was built to analyze just 12 stocks. I then refactored the code to be able to analyze the entire stock market over the last few years.

## Results

After the repurposed code is complete and I ran the analysis for the stocks in both 2017 and 2018 it is clear that 2017 had overall better returns than in 2018. In the original code it would take a longer time to run all the vaiorus stock analysis, but in the refactored code it is built to handle analysisng many stocks. I have below the execution times for the refactored code for each year.

![2017 Execution Time]



![2018 Execution Time]


## Summary

### Advantages

An advantage of refactoring code is that it can improve the existing code and it can make the code cleaner and more condense. It can improve code because you are building off of what is already, making it better. For instance in this project I was able to improve our code so that it now has the capability to run an analysis for the entire stock market now. This is because it is able to run at a much faster speed. 

Another advantage of refactoring code in general is that you are able to clean up the original code and make it more condse. This can make the code easier to read and understand if someone else is viewing the code. In this project we were able to condese many sub functions into one or two sub functions. This made the code more optimized. With just one sub function way we were able to run the analysis in complete, including the conditional formatting, with the click of a button. In the original code, although it would run the analysis, because the formatting was in its own sub function if we cleared the workbook and then reran the analysis all of the conditional formatting was gone.

### Disadvantages

The disadvantage with refactoring code is that you can easily break the code. If you are not thorough in your modifications in structuring your code there could be something from the original code that does not work with the new code which can cause errors. In this code with combining the necessary code to one sub function for just the "All Stock Analysis" it would cause errors if we didn't also refactor the code for the "DQ Analysis" worksheet. Since we did not also refactor that code into this project that sheet currently returns an error when trying to run, therefore was not included in the final project submission.
