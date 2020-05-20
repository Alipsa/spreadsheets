library('hamcrest')
library('se.alipsa:excel-utils')

test.findRowNumSunny <- function() {
  rowNum <- findRowNumber("df.xlsx", 0, 0, "Iris")
  assertThat(rowNum, equalTo(35))
}

test.barf <- function() {
  tryCatch( 
    {
      barf("Some problem")
    }, 
    
    error = function(err) {
      #print(err$message)
      assertTrue(endsWith(err$message, " problem"))
    }
  )
}

test.finRowNumRainy <- function() {
  tryCatch(
    findRowNumber("doesnotexist.xlsx", 0, 0, "Iris"),
    
    error = function(err) {
      #print(paste("Expected error was: ", err))
      assertTrue(endsWith(err$message, " not exist"))
    }
  )

  rowNum <- findRowNumber("df.xlsx", 0, 0, "Nothing that exist")
  assertThat(rowNum, equalTo(-1))
}

test.findColumnsSunny <- function() {
  colNum <- findColumnNumber("df.xlsx", 0, 1, "carb")
  assertThat(colNum, equalTo(10L))
}

test.importExcelWithHeaderRow <- function() {
  excelDf <- importExcel(
    filePath = "df.xlsx",
    sheetNumber = 0,
    startRowNum = 1,
    endRowNum = 33,
    startColNum = 0,
    endColNum = 11,
    firstRowAsColNames = TRUE
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
    sheetNumber = 0,
    startRowNum = 2,
    endRowNum = 33,
    startColNum = 0,
    endColNum = 11,
    firstRowAsColNames = FALSE
  )
  assertThat(nrow(excelDf), equalTo(32))
  assertThat(ncol(excelDf), equalTo(11))
}

test.importExcelWithHeaderNames <- function() {
  excelDf <- importExcel(
    filePath = "df.xlsx",
    sheetNumber = 0,
    startRowNum = 2,
    endRowNum = 33,
    startColNum = 0,
    endColNum = 11,
    columnNames = c("1","2","3","4","5","6","7","8","9","10","11")
  )
  assertThat(nrow(excelDf), equalTo(32))
  assertThat(ncol(excelDf), equalTo(11))
}
