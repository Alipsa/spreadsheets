# Spreadsheets - Handling spreadsheets in Renjin R

This package will give you the ability to work with (read, write) spreadsheets.

It currently supports reading of excel files but will be able to handle
Open Office spreadsheets as well soon.

## Usage
* All indexes start with 1 (as is common practice in R), e.g. sheetNumber 1 refers to the 
first sheet in the spreadsheet and column number 1 is the first (A) column etc.

### Find a row in a column


## Background / motivation
Sometimes I have problems with loading the `xlsx` package and I miss some 
search functionality to make imports more dynamic in my R code. This is a "Renjin native"
package that attempts to address some of those issues.