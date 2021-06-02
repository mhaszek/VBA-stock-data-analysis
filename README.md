# Stock Data Analysis with VBA
VBA project for Monash Data Analytics Boot Camp

The purpose of this project was to use VBA scripting to analyze real stock market data.

# Data

There were two key sources of data used:

* [Test Data](Resources/alphabetical_testing.xlsx) - short version of the main dataset used for script testing

* [Stock Data](Resources/Multiple_year_stock_data.xlsx) - main dataset: .xlsx file which contains stock data for 2014, 2015 and 2016
![data](Images/data.JPG)

# Analysis


* Loop through all the stocks for one year and output the following information:

  * The ticker symbol.

  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The total stock volume of the stock.

* Apply conditional formatting that will highlight positive change in green and negative change in red.

* Return the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume".

### 2014:
![2014_start](Images/HW02-VBA_MilenaH_2014.JPG)
![2014_end](Images/HW02-VBA_MilenaH_2014_last.JPG)

### 2015:
![2015_start](Images/HW02-VBA_MilenaH_2015.JPG)
![2015_end](Images/HW02-VBA_MilenaH_2015_last.JPG)

### 2016:
![2016_start](Images/HW02-VBA_MilenaH_2016.JPG)
![2016_end](Images/HW02-VBA_MilenaH_2016_last.JPG)


# Demo

To run the examples locally, save the chosen data file as Excel Macro-Enabled Workbook, go to `Developer` tab, then open `Visual Basic`, import and run the `Homework_02-VBA_MilenaH` VBScript Script File.


# Used Tools
 * VBA

#

#### Contact: mil.haszek@gmail.com