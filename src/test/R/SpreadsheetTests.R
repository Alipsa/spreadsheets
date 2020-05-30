library('hamcrest')
library('se.alipsa:spreadsheets')

wdfile <- function(fileName) {
  paste0(getwd(), "/", fileName)
}

findRowNumSunnyTest <- function(fileName) {
  rowNum <- findRowNumber(fileName = fileName, sheet = 1, column = 1, "Iris")
  assertThat(rowNum, equalTo(36))

  rowNum <- findRowNumber(fileName = fileName, sheet = "project-dashboard", column = 1, "Iris")
  assertThat(rowNum, equalTo(36))
}

test.findRowNumSunnyExcel <- function() {
  findRowNumSunnyTest(wdfile("df.xlsx"))
}

test.findRowNumSunnyOds <- function() {
  findRowNumSunnyTest(wdfile("df.ods"))
}

finRowNumRainy <- function(notExistingFileName, fileName) {
  tryCatch(
    findRowNumber(fileName = notExistingFileName, sheet = 1, column = 1, "Iris"),

    error = function(err) {
      #print(paste("Expected error was: ", err))
      assertTrue(endsWith(err$message, " not exist"))
    }
  )

  rowNum <- findRowNumber(fileName, 1, 1, "Nothing that exist")
  assertThat(rowNum, equalTo(-1))
}

test.finRowNumRainyExcel <- function() {
  finRowNumRainy("doesnotexist.xlsx", wdfile("df.xlsx"))
}
test.finRowNumRainyOds <- function() {
  finRowNumRainy("doesnotexist.ods", wdfile("df.ods"))
}

findColumnsSunny <- function(fileName) {
  colNum <- findColumnNumber(fileName, 1, 2, "carb")
  assertThat(colNum, equalTo(11L))

  colNum <- findColumnNumber(fileName, "project-dashboard", 2, "carb")
  assertThat(colNum, equalTo(11L))
}

test.findColumnsSunnyExcel <- function() {
  findColumnsSunny(wdfile("df.xlsx"))
}
test.findColumnsSunnyOds <- function() {
  findColumnsSunny(wdfile("df.ods"))
}

importWithHeaderRow <- function(fileName) {
  df <- importSpreadsheet(
    filePath = fileName,
    sheet = 1,
    startRow = 2,
    endRow = 34,
    startColumn = 1,
    endColumn = 11,
    firstRowAsColumnNames = TRUE
  )
  assertThat(nrow(df), equalTo(32))
  assertThat(ncol(df), equalTo(11))
  assertThat(as.numeric(sub(",", ".", df$mpg)), equalTo(mtcars$mpg))
  G26 <- df$qsec[24]
  assertThat(G26, equalTo("15.41"))
}

test.importWithHeaderRowExcel <- function() {
  importWithHeaderRow(wdfile("df.xlsx"))
}

test.importWithHeaderRowOds <- function() {
  importWithHeaderRow(wdfile("df.ods"))
}

importNoHeaderRow <- function(fileName) {
  df <- importSpreadsheet(
    filePath = fileName,
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

test.importNoHeaderRowExcel <- function() {
  importNoHeaderRow(wdfile("df.xlsx"))
}
test.importNoHeaderRowOds <- function() {
  importNoHeaderRow(wdfile("df.ods"))
}

importWithHeaderNames <- function(fileName) {
  df <- importSpreadsheet(
    filePath = fileName,
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

test.importWithHeaderNamesExcel <- function() {
  importWithHeaderNames(wdfile("df.xlsx"))
}
test.importWithHeaderNamesOds <- function() {
  importWithHeaderNames(wdfile("df.ods"))
}

importComplex <- function(fileName) {
  df <- importSpreadsheet(
    filePath = fileName,
    sheet = 1,
    startRow = 1,
    endRow = 7,
    startColumn = "A",
    endColumn = "F",
    firstRowAsColumnNames = TRUE
  )
  date <- as.Date(as.POSIXlt(df[3,1]))
  assertThat(date, equalTo(as.Date("2020-05-03")))
  assertThat(as.character(as.POSIXlt(df[3,2])), equalTo("2020-05-03 15:43:12"))
  assertThat(as.integer(df[3,3]), equalTo(102L))
  assertThat(df[3,4], equalTo("5.222"))
  assertThat(df[3,5], equalTo("three"))
  assertThat(df[3,6], equalTo("96.778"))
}

test.importComplexExcel <- function() {
  importComplex(wdfile("complex.xlsx"))
}

test.importComplexOds <- function() {
  importComplex(wdfile("complex.ods"))
}

test.columnNameConversions <- function() {
  assertThat(columnIndex("N"), equalTo(14))
  assertThat(columnIndex("AF"), equalTo(32))
  assertThat(columnIndex("AAB"), equalTo(704))

  assertThat(columnName(14), equalTo("N"))
  assertThat(columnName(32), equalTo("AF"))
  assertThat(columnName(704), equalTo("AAB"))
}

exportNew <- function(fileName) {
  if (file.exists(fileName)) file.remove(fileName)
  result <- exportSpreadsheet(mtcars, fileName)
  assertThat(result, equalTo(TRUE))
  gearCol <- findColumnNumber(fileName, 1, 1, "gear")
  expected <- columnIndex("J")
  assertThat(gearCol, equalTo(expected))
}

test.exportNewExcel <- function() {
  exportNew(wdfile("test.xlsx"))
}

test.exportNewOds <- function() {
  exportNew(wdfile("test.ods"))
}

update <- function(fileName) {
  if (file.exists(fileName)) file.remove(fileName)
  result <- exportSpreadsheet(mtcars, fileName)
  assertThat(result, equalTo(TRUE))
  result <- exportSpreadsheet(iris, fileName, "iris")
  assertThat(result, equalTo(TRUE))
  gearCol <- findColumnNumber(fileName, 1, 1, "gear")
  assertThat(gearCol, equalTo(columnIndex("J")))
  versicolorRow <- findRowNumber(fileName, "iris", columnIndex("E") , "versicolor")
  assertThat(versicolorRow, equalTo(52))
}

test.updateExcel <- function() {
  update(wdfile("test2.xlsx"))
}

test.updateOds <- function() {
  update(wdfile("test2.ods"))
}

exportSpreadsheetsTest <- function(fileName) {
  if (file.exists(fileName)) file.remove(fileName)
  sheetNames <- c("mtcars", "iris", "PlantGrowth")
  result <- exportSpreadsheets(list(mtcars, iris, PlantGrowth), sheetNames, fileName)
  assertThat(result, equalTo(TRUE))
  sheets <- getSheetNames(fileName)
  assertThat(sheetNames, equalTo(sheets))
}

test.exportSpreadsheetsExcel <- function() {
  exportSpreadsheetsTest("testSheets.xlsx")
}

test.exportSpreadsheetsOds <- function() {
  exportSpreadsheetsTest("testSheets.ods")
}

test.exportSpreadsheetRainy <- function() {
  assertThat(exportSpreadsheets(NA, sheetNames, fileName), throwsError())

  assertThat(exportSpreadsheets(list(mtcars), NULL, fileName), throwsError())
}
