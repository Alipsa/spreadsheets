# Spreadsheets - Handling spreadsheets in Renjin R

This package will give you the ability to work with (read, write) spreadsheets.

It supports reading of excel and Open Office/Libre Office spreadsheets files.

To use it add the following dependency to your pom.xml:
```xml
<dependency>
  <groupId>se.alipsa</groupId>
  <artifactId>spreadsheets</artifactId>
  <version>1.2</version>
</dependency>
```
and use it your Renjin R code after loading it with:
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

### exportSpreadsheet: export an excel or Open Office spreadsheet

To export to a new spread sheet use
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
Also, I missed some search functionality to make imports more dynamic in my R code. 
As the gcc-bridge (which compiles C code to jvm byte code) gets better, the first kind of problem will disappear,
but I needed something "now". This is a "Renjin native" package which attempts to address some of those issues.

## Dependencies / 3:rd party libraries used

1. Renjin (https://www.renjin.org/, https://github.com/bedatadriven/renjin).
This is a Renjin package (extension) so obviously it requires Renjin to use. 
I have tested with version 3.5-beta76 but there is no particular Renjin version required, 
anything from version 0.9 and later should work.

2. POI (https://poi.apache.org/)
Used to read and write Excel files. Built and tested with poi version 4.1.2.

3. SODS (https://github.com/miachm/SODS)
Used to read and write Open Document Spreadsheets (Open Office / Libre Office Calc files).
Built and tested with SODS version 1.2.2.


# Version history

1.3
- close workbook properly when calling getSheetNames()

1.2
- Changed from primitives to Object wrappers (int -> Integer etc.) so that we can correctly return
NULL for missing values (which will be NA in the data.frame).
- Allow export to update existing file.

1.1
- Api change: modified the api so that we always start with filePath to make it more consistent.
              renamed columnIndex function to as.columnIndex and similar for columnName.
- Add support for exporting multiple data.frames   
- Enhanced documentation

1.0 Initial release           