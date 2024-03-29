# Spreadsheets - Handling spreadsheets in Renjin R

This package will give you the ability to work with (read, write) spreadsheets.

It supports reading of Excel and Open Office/LibreOffice spreadsheets files.

To use it add the following dependency to your pom.xml:
```xml
<dependency>
  <groupId>se.alipsa</groupId>
  <artifactId>spreadsheets</artifactId>
  <version>1.3.4</version>
</dependency>
```
(Note that version 1.3.4 and later requires java 11). The module name is se.alipsa.spreadsheets.

...and use it your Renjin R code after loading it with:
```r
library("se.alipsa:spreadsheets")
```

## Usage
* All indexes start with 1 (as is common practice in R), e.g. sheetNumber 1 refers to the 
first sheet in the spreadsheet and column number 1 is the first (A) column etc.

The file extension is used to determine whether it is an Excel (xls/xlsx) or Calc (ods) file. 

### findRowNumber: Find a row in a column
To find the first row where the cell value matches the cellContent parameter:  

```r
rowNum <- findRowNumber(filePath = "df.xlsx", sheet = 1, column = 1, "Iris")
```

You can also reference the sheet by name:

```r
rowNum <- findRowNumber(filePath = "df.ods", sheet = "theSheetName", column = 1, "Iris")
```

or only use names

```r
rowNum <- findRowNumber(filePath = "df.xlsx", sheet = "theSheetName", column = "A", "Iris")
```

### findColumnNumber: Find a column in a row
To find the first column where the cell value matches the cellContent parameter:  

```r
colNum <- findColumnNumber(filePath = "df.xlsx", sheet = 1, row = 2, "carb")`
```

You can also reference the sheet by name:

```r
colNum <- findColumnNumber("df.xlsx", "project-dashboard", 2, "carb")
```

The return value of findColumnNumber is an Integer with the matching row index
or -1 if no such cell was found.

### columnIndex and columnName: Get the index number for the corresponding column name and vice versa
Sometimes it is more convenient to refer to the column by the name e.g. A for the first column, B for the second.
To convert an index to a name you can do:
```r
print(as.columnName(14))
[1] "N"
```

But sometimes you want the other way around:

```r
print(as.columnIndex("AF"))
[1] 32
```

### importSpreadsheet: import an Excel or Open Office spreadsheet
Reads the content of the spreadsheet and return a data.frame
```r
excelDf <- importSpreadsheet(
    filePath = "df.xlsx",
    sheet = 1,
    startRow = 2,
    endRow = 34,
    startColumn = 1,
    endColumn = 11,
    firstRowAsColumnNames = TRUE
  )
```
The parameters are as follows:
* filePath: The filePath to the excel file to import. It must be a path to file that is physically accessible. A remote url will not work.
* sheet: The sheet index (index starting with 1) for the sheet to import. Can alternatively be the name of the sheet. Default: 1 
* startRow: The row to start reading from. Default: 1
* endRow: The last row to read from
* startColumn: The column index (or name e.g. "A") to start reading from. default: 1
* endColumn: The last column index (or name) to read from.
* firstRowAsColumnNames: If true then use the values of the first column as column names for the data.frame

_Return value_ A data.frame of Character vectors (strings).

Since the resulting dataframe will return all values as character strings (except missing values which will be NA), 
so you will likely need to massage the data after the import to get what you want. e.g.
```r
excelDf$mpg <- as.numeric(sub(",", ".", excelDf$mpg))
```
In the example above, the regional setting of the excel sheet used comma as the decimal separator so we replace them with 
dots to we can then convert them to numerics.

Dates are converted to strings in the format yyyy-MM-dd HH:mm:ss.SSS which is the default format for POSIXct and POSIXlt so you can do:
```r
library("se.alipsa:spreadsheets")
timeMeasuresDf <- importSpreadsheet(
  filePath = "E:\\some\\path\\data\\timeMeasures.ods",
  sheet = 1,
  startRow = 1,
  endRow = 7,
  startColumn = "A",
  endColumn = "F",
  firstRowAsColumnNames = TRUE
)
# change the startDate column to Dates: 
timeMeasuresDf$startDate <- as.Date(as.POSIXlt(timeMeasuresDf$startDate))
```

### importSpreadsheets: import several Excel or Open Office spreadsheets at once
Reads the content of the spreadsheets and return a named list of data.frame's
```r
sheets <- importSpreadsheets(
  filePath=paste0(getwd(), "/mySpreadseet.ods"),
  sheets = c('mtcars', 'iris', 'PlantGrowth'),
  importAreas = list(
    'mtcars' = c(1, 33, 1, 11),
    'iris' = c(2, 152, 1, 5),
    'PlantGrowth' = c(3, 32, 2, 3)
  ),
  firstRowAsColumnNames = list(
    'mtcars' = TRUE,
    'iris' = TRUE,
    'PlantGrowth' = FALSE
  )
)
  
irisDf <- sheets$iris 
```
The parameters are as follows:
* _filePath_ the full path or relative path to the Excel file
* _sheetNames_ a vector of sheet names e.g. `c('sheet1', 'sheet2')`
* _importAreas_ a named list of numeric vectors containing start row, end row, start column, end column e.g.
  `list('sheet1' = c(1, 33, 1, 11), 'sheet2' = c(2, 152, 1, 5))`
* _firstRowAsColumnNames_ a named vector of logical values for whether the first row should be used as 
  column names for the dataframe in the sheet or not.
  E.g. `list('sheet1' = TRUE, 'sheet2' = FALSE)`

_Return value_ a named vector of data.frame's (ListVectors) corresponding to the imported sheets

See import importSpreadsheet for notes about values conversion.

### exportSpreadsheet: export an excel or Open Office spreadsheet

To export to a new spreadsheet use
```r
exportSpreadsheet(filePath, df)
```
Where filePath the path to the new sheet and df is the data-frame to export. If the file already exist, no action
will be taken.


The "upsert" (create new if not exists, update if exist) version is:

```r
exportSpreadsheet(filePath, df, sheet)
```
Where df is the data-frame to export and filePath the path to the new or existing spreadsheet, 
and sheet is the sheet name to create or update. 

The function returns TRUE if successful or FALSE if not. 

### exportSpreadsheets: export multiple data.frames to an excel or Open Office spreadsheet
Just like above, when you have several dataframes that you want to export in one go you can
do it like this:
```r
exportSpreadsheets(
  filePath = paste0(getwd(), "/dfExport.ods"), 
  dfList = list(mtcars, iris, PlantGrowth), 
  sheetNames = c("cars", "flowers", "plants")
)
```
The number of sheet names must match the number of data frames in the list.


There are more functions in the api than what is described above, see [SpreadsheetTests.R](https://github.com/Alipsa/spreadsheets/blob/master/src/test/R/SpreadsheetTests.R) for more examples.

## Background / motivation
Why not just use one of the existing packages such as xlsx, XLConnect, or gdata? 
Sometimes I had problems with loading these packages, or some functions did not work (none of them fully passes 
the tests on renjin cran).
Also, I missed some search functionality to make imports more dynamic in my R code as well as the ability to handle 
the OpenOffice format (readOds is not available in Renjin yet).
As the gcc-bridge (which compiles C code to jvm byte code) gets better, the first kind of problem will disappear,
but I needed something "now". This is a "Renjin native" package which attempts to address some of those issues.

## Dependencies / 3:rd party libraries used

1. Renjin (https://www.renjin.org/, https://github.com/bedatadriven/renjin).
This is a Renjin package (extension) so obviously it requires Renjin to use. 
I have tested with version 3.5-beta76 but there is no particular Renjin version required, 
anything from version 0.9 and later should work.

2. POI (https://poi.apache.org/)
Used to read and write Excel files. Built and tested with poi version 5.

3. SODS (https://github.com/miachm/SODS)
Used to read and write Open Document Spreadsheets (Open Office / Libre Office Calc files).
Built and tested with SODS version 1.4.


# Version history

### 1.3.5
- Add Automatic-Module-Name

### 1.3.4, Apr 12, 2022
- Add support for import of multiple sheets at once
- Upgrade to java 11
- Upgrade apache poi dependencies

### 1.3.3, Feb 6, 2022
- make ods import behave similar to excel when importing percentages (i.e import it as a decimal e.g. 0.54 instead of 54%)
- improve test: check that column headers are imported correctly
- upgrade poi and slf4j

### 1.3.2, Jan 30, 2022
- Change data.frame creation of row.names to be future-proof by replacing the RowNamesVector with a ConvertingStringVector
- update poi, logging and some plugin versions

### 1.3.1, Aug 22, 2021
- upgrade dependencies (notably SODS which in version 1.4 has a greatly reduced footprint)

### 1.3, Feb 21, 2021
- close workbook properly when calling getSheetNames()
- upgrade SODS and poi versions

### 1.2, Aug 02, 2020
- Changed from primitives to Object wrappers (int -> Integer etc.) so that we can correctly return
NULL for missing values (which will be NA in the data.frame).
- Allow export to update existing file.

### 1.1, May 31, 2020
- Api change: modified the api so that we always start with filePath to make it more consistent.
              renamed columnIndex function to as.columnIndex and similar for columnName.
- Add support for exporting multiple data.frames   
- Enhanced documentation

### 1.0 Initial release, May 27, 2020           