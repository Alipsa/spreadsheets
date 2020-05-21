library('hamcrest')
library('se.alipsa:spreadsheets')

test.findRowNumSunny <- function() {
  rowNum <- findRowNumber("df.xlsx", 1, 1, "Iris")
  assertThat(rowNum, equalTo(36))
}

test.finRowNumRainy <- function() {
  tryCatch(
    findRowNumber("doesnotexist.xlsx", 1, 1, "Iris"),
    
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
}
