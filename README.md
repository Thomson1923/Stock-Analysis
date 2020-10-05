# VBA Stock Analysis

# NOTE TO FILE

# From: Bob
# Re: Development of VBA Code to Calculate Key Information for Publicly Traded Stocks

# Project Overview
## We recently developed Visual Basic for Application (VBA) code to calculate key stock information (annual trading volume and investment returns during calendar years) from daily information (). While calculation of these parameters can be easily performed using standard Excel formulas in typical worksheets, the VBA code provides the means for rapid calculation of numerous stocks during user-selected years and formatted output to simplify analysis.

# Results
## Stock Performance
### Comparison of the output for 12 selected stocks shows that most of the chosen stocks provide positive returns during 2017 (https://github.com/Thomson1923/Stock-Analysis/blob/master/Resources/VBA_Challenge_2017_Initial.png) while most showed disappointing performance the following year (https://github.com/Thomson1923/Stock-Analysis/blob/master/Resources/VBA_Challenge_2018_Initial.png). The use of conditional formatting (green cell interior shading for positive returns and red interior shading for negative returns) provides an extremely rapid means for assessing the performance of these stocks.

## Refactoring Code
### This project relied on code obtained from another source and included substantial refactoring work. The starting code, which relied on repeated review of the data sets to extract the relevant values, required well over one second to complete for information from 2017 and 2018 (see exhibits from Stock Performance). However, refactoring the code to allow for al data to be retrieved through a single review of the data set lowered the execution time by 82% for 2017 data (https://github.com/Thomson1923/Stock-Analysis/blob/master/Resources/VBA_Challenge_2017_Final.png) and 74% for 2018 data (https://github.com/Thomson1923/Stock-Analysis/blob/master/Resources/VBA_Challenge_2018_Final.png). The benefits offered by this improvement in the code will become clearer as we expended the number of stocks analyzed.

# Summary
### This project offers an excellent example of how refactoring code can lead to substantial improvement by finding better algorithms to complete the designated task, though one must have a clear understanding of the logic and structure of the initial code to ensure that refactoring work does not inadvertently lead to code that does not perform the given task or provide incorrect results. Using a single review of the data set to extract all relevant values presents the risk for information mismatch where the prices and volumes are not properly correlated to the correct ticker symbol.

