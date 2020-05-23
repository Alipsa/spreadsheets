# Spreadsheets - Handling spreadsheets in Renjin R

This package will give you the ability to work with (read, write) spreadsheets.

It currently supports reading of excel files but will be able to handle
Open Office spreadsheets as well soon.

## Usage
* All indexes start with 1 (as is common practice in R), e.g. sheetNumber 1 refers to the 
first sheet in the spreadsheet and column number 1 is the first (A) column etc.

### findRowNumber: Find a row in a column
To find the first row where the cell value matches the cellContent parameter:  

```r
rowNum <- findRowNumber(fileName = "df.xlsx", sheet = 1, column = 1, "Iris")
```

You can also reference the sheet by name:

```r
rowNum <- findRowNumber(fileName = "df.xlsx", sheet = "theSheetName", column = 1, "Iris")
```

or only use names

```r
rowNum <- findRowNumber(fileName = "df.xlsx", sheet = "theSheetName", column = "A", "Iris")
```

### findColumnNumber: Find a column in a row
To find the first column where the cell value matches the cellContent parameter:  

```r
colNum <- findColumnNumber(fileName = "df.xlsx", sheet = 1, row = 2, "carb")`
```

You can also reference the sheet by name:

```r
colNum <- findColumnNumber("df.xlsx", "project-dashboard", 2, "carb")
```

The return value of findColumnNumber is an Integer with the matching row index
or 0 if no such cell was found.

### columnIndex and columnName: Get the index number for the corresponding column name and vice versa
Sometimes it is more convenient to refer to the column by the name e.g. A for the first column, B for the second.
To convert an index to a name you can do:
```r
print(columnName(14))
[1] "N"
```

But sometimes you want the other way around:

```r
print(columnIndex("AF"))
[1] 32
```

### importExcel: import a spreadsheet
Reads the content of the spreadsheet and return a data.frame
```r
excelDf <- importExcel(
    filePath = "df.xlsx",
    sheet = 1,
    startRow = 2,
    endRow = 34,
    startColumn = 1,
    endColumn = 11,
    firstRowAsColumnNames = TRUE
  )
```
The resulting dataframe will read all values as character strings so you will likely need to
massage the data efter the import to get what you want. e.g.

```r
excelDf$mpg <- as.numeric(sub(",", ".", excelDf$mpg)
```

In the example above, the regional setting of the excel sheet used comma as the decimal separator so we replace them with 
dots to we can then convert them to numerics.

The parameters are as follows:
* filePath: The filePath to the excel file to import. It must be a path to file that is physically accessible. A remote url will not work.
* sheet: The sheet index (index starting with 1) for the sheet to import. Can alternatively be the name of the sheet. Default: 1 
* startRow: The row to start reading from. Default: 1
* endRow: The last row to read from
* startColumn: The column index (or name e.g. "A") to start reading from. default: 1
* endColumn: The last column index (or name) to read from.
* firstRowAsColumnNames: If true then use the values of the first column as column names for the data.frame

### exportExcel: export a spreadsheet

To export to a new excel sheet use
```r
exportExcel(df, filePath)
```
Where df is the data-frame to export and filePath the path to the new sheet. If the excel file already exist, no action
will be taken.


The "upsert" (create new if not exists, update if exist) version is:

```r
exportExcel(df, filePath, sheet)
```
Where df is the data-frame to export and filePath the path to the new or existing spreadsheet, 
and sheet is the sheet name to create or update. 

The function returns TRUE if successful or FALSE if not. 

## Background / motivation
Why not just use one of the existing packages such as xlsx, XLConnect, or gdata? 
Sometimes I had problems with loading these packages, or some functions did not work (none of them fully passes 
the tests on renjin cran).
Also I missed some search functionality to make imports more dynamic in my R code. 
As the gcc-bridge (which compiles C code to jvm byte code) gets better, the first kind of problem will disappear,
but I needed something "now". This is a "Renjin native" package that attempts to address some of those issues.
