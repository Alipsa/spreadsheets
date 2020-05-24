library('hamcrest')
library('se.alipsa:spreadsheets')

test.findOdsRowNumSunny <- function() {
  rowNum <- findRowNumber(fileName = "df.ods", sheet = 1, column = 1, "Iris")
  assertThat(rowNum, equalTo(36))

  rowNum <- findRowNumber(fileName = "df.ods", sheet = "project-dashboard", column = 1, "Iris")
  assertThat(rowNum, equalTo(36))
}

test.findOdsRowNumRainy <- function() {
  tryCatch(
    findRowNumber(fileName = "doesnotexist.ods", sheet = 1, column = 1, "Iris"),
    
    error = function(err) {
      #print(paste("Expected error was: ", err))
      assertTrue(endsWith(err$message, " not exist"))
    }
  )

  rowNum <- findRowNumber("df.ods", 1, 1, "Nothing that exist")
  assertThat(rowNum, equalTo(-1))
}

test.findOdsColumnsSunny <- function() {
  colNum <- findColumnNumber("df.ods", 1, 2, "carb")
  assertThat(colNum, equalTo(11L))

  colNum <- findColumnNumber("df.ods", "project-dashboard", 2, "carb")
  assertThat(colNum, equalTo(11L))
}

test.importOdsWithHeaderRow <- function() {
  excelDf <- importSpreadsheet(
    filePath = "df.ods",
    sheet = 1,
    startRow = 2,
    endRow = 34,
    startColumn = 1,
    endColumn = 11,
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

test.importOdsNoHeaderRow <- function() {
  df <- importSpreadsheet(
    filePath = "df.ods",
    sheet = 1,
    startRow = 3,
    endRow = 34,
    startColumn = 1,
    endColumn = 11,
    firstRowAsColumnNames = FALSE
  )
  assertThat(nrow(df), equalTo(32))
  assertThat(ncol(df), equalTo(11))
  assertThat(as.numeric(df[,2]), equalTo(mtcars$cyl))
}

test.importOdsWithHeaderNames <- function() {
  df <- importSpreadsheet(
    filePath = "df.ods",
    sheet = 1,
    startRow = 3,
    endRow = 34,
    startColumn = 1,
    endColumn = 11,
    columnNames = c("1","2","3","4","5","6","7","8","9","10","11")
  )
  assertThat(nrow(df), equalTo(32))
  assertThat(ncol(df), equalTo(11))
  assertThat(as.numeric(df[,4]), equalTo(mtcars$hp))
}

test.exportNewOds <- function() {
  exportSpreadsheet(mtcars, "test.ods")
  gearCol <- findColumnNumber("test.xlsx", 1, 1, "gear")
  expected <- columnIndex("J")
  assertThat(gearCol, equalTo(expected))
}

test.updateOds <- function() {
  exportSpreadsheet(mtcars, "test2.ods")
  exportSpreadsheet(iris, "test2.ods", "iris")
  gearCol <- findColumnNumber("test2.ods", 1, 1, "gear")
  assertThat(gearCol, equalTo(columnIndex("J")))
  versicolorRow <- findRowNumber("test2.ods", "iris", columnIndex("E") , "versicolor")
  assertThat(versicolorRow, equalTo(52))
}
