library('hamcrest')
library('se.alipsa:spreadsheets')

test.findRowNumSunny <- function() {
  rowNum <- findRowNumber(fileName = "df.xlsx", sheet = 1, column = 1, "Iris")
  assertThat(rowNum, equalTo(36))

  rowNum <- findRowNumber(fileName = "df.xlsx", sheet = "project-dashboard", column = 1, "Iris")
  assertThat(rowNum, equalTo(36))
}

test.finRowNumRainy <- function() {
  tryCatch(
    findRowNumber(fileName = "doesnotexist.xlsx", sheet =1, column = 1, "Iris"),
    
    error = function(err) {
      #print(paste("Expected error was: ", err))
      assertTrue(endsWith(err$message, " not exist"))
    }
  )

  rowNum <- findRowNumber("df.xlsx", 1, 1, "Nothing that exist")
  assertThat(rowNum, equalTo(0))
}

test.findColumnsSunny <- function() {
  colNum <- findColumnNumber("df.xlsx", 1, 2, "carb")
  assertThat(colNum, equalTo(11L))

  colNum <- findColumnNumber("df.xlsx", "project-dashboard", 2, "carb")
  assertThat(colNum, equalTo(11L))
}

test.importExcelWithHeaderRow <- function() {
  excelDf <- importExcel(
    filePath = "df.xlsx",
    sheet = 1,
    startRow = 2,
    endRow = 34,
    startColumn = 1,
    endColumn = 12,
    firstRowAsColumnNames = TRUE
  )
  #print(head(excelDf,1))
  #print(tail(excelDf,1))
  assertThat(nrow(excelDf), equalTo(32))
  assertThat(ncol(excelDf), equalTo(11))
  assertThat(as.numeric(sub(",", ".", excelDf$mpg)), equalTo(mtcars$mpg))
  G26 <- excelDf$qsec[24]
  assertThat(G26, equalTo("15.41"))
}

test.importExcelNoHeaderRow <- function() {
  excelDf <- importExcel(
    filePath = "df.xlsx",
    sheet = 1,
    startRow = 3,
    endRow = 34,
    startColumn = 1,
    endColumn = 12,
    firstRowAsColumnNames = FALSE
  )
  assertThat(nrow(excelDf), equalTo(32))
  assertThat(ncol(excelDf), equalTo(11))
  assertThat(as.numeric(excelDf[,2]), equalTo(mtcars$cyl))
}

test.importExcelWithHeaderNames <- function() {
  excelDf <- importExcel(
    filePath = "df.xlsx",
    sheet = 1,
    startRow = 3,
    endRow = 34,
    startColumn = 1,
    endColumn = 12,
    columnNames = c("1","2","3","4","5","6","7","8","9","10","11")
  )
  assertThat(nrow(excelDf), equalTo(32))
  assertThat(ncol(excelDf), equalTo(11))
  assertThat(as.numeric(excelDf[,4]), equalTo(mtcars$hp))
}

test.columnNameConversions <- function() {
  assertThat(columnIndex("N"), equalTo(14))
  assertThat(columnIndex("AF"), equalTo(32))
  assertThat(columnIndex("AAB"), equalTo(704))

  assertThat(columnName(14), equalTo("N"))
  assertThat(columnName(32), equalTo("AF"))
  assertThat(columnName(704), equalTo("AAB"))
}
