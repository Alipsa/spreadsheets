library('hamcrest')
library('se.alipsa:spreadsheets')

wdfile <- function(filePath) {
  paste0(getwd(), "/", filePath)
}

findRowNumSunnyTest <- function(filePath) {
  rowNum <- findRowNumber(filePath = filePath, sheet = 1, column = 1, "Iris")
  assertThat(rowNum, equalTo(36))

  rowNum <- findRowNumber(filePath = filePath, sheet = "project-dashboard", column = 1, "Iris")
  assertThat(rowNum, equalTo(36))
}

test.findRowNumSunnyExcel <- function() {
  findRowNumSunnyTest(wdfile("df.xlsx"))
}

test.findRowNumSunnyOds <- function() {
  findRowNumSunnyTest(wdfile("df.ods"))
}

finRowNumRainy <- function(notExistingFileName, filePathThatExist) {
  if (!file.exists(filePathThatExist)) {
    stop("Wrong test data, filePathThatExistmust exist")
  }
  tryCatch(
    findRowNumber(filePath = notExistingFileName, sheet = 1, column = 1, "Iris"),

    error = function(err) {
      #print(paste("Expected error was: ", err))
      assertTrue(endsWith(err$message, " not exist"))
    }
  )

  rowNum <- findRowNumber(filePathThatExist, 1, 1, "Nothing that exist")
  assertThat(rowNum, equalTo(-1))
}

test.finRowNumRainyExcel <- function() {
  finRowNumRainy("doesnotexist.xlsx", wdfile("df.xlsx"))
}
test.finRowNumRainyOds <- function() {
  finRowNumRainy("doesnotexist.ods", wdfile("df.ods"))
}

findColumnsSunny <- function(filePath) {
  colNum <- findColumnNumber(filePath, 1, 2, "carb")
  assertThat(colNum, equalTo(11L))

  colNum <- findColumnNumber(filePath, "project-dashboard", 2, "carb")
  assertThat(colNum, equalTo(11L))
}

test.findColumnsSunnyExcel <- function() {
  findColumnsSunny(wdfile("df.xlsx"))

  colNum <- findColumnNumber("complex.xlsx", 1, 2, "96.748")
  assertThat(colNum, equalTo(6L))

  # This does not work: Named values does not work
  #colNum <- findColumnNumber("complex.xlsx", 2, 3, "BEG")
  #assertThat(colNum, equalTo(2L))
}


test.findColumnsSunnyOds <- function() {
  findColumnsSunny(wdfile("df.ods"))

  colNum <- findColumnNumber("complex.ods", 1, 5, "96.019")
  assertThat(colNum, equalTo(6L))
}

importWithHeaderRow <- function(filePath) {
  df <- importSpreadsheet(
    filePath = filePath,
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
  K34 <- as.integer(df$carb[32])
  assertThat(K34, equalTo("2"))
}

test.importWithHeaderRowExcel <- function() {
  importWithHeaderRow(wdfile("df.xlsx"))
}

test.importWithHeaderRowOds <- function() {
  importWithHeaderRow(wdfile("df.ods"))
}

importNoHeaderRow <- function(filePath) {
  df <- importSpreadsheet(
    filePath = filePath,
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
  assertThat(as.integer(df[,11]), equalTo(mtcars$carb))
}

test.importNoHeaderRowExcel <- function() {
  importNoHeaderRow(wdfile("df.xlsx"))
}
test.importNoHeaderRowOds <- function() {
  importNoHeaderRow(wdfile("df.ods"))
}

importWithHeaderNames <- function(filePath) {
  df <- importSpreadsheet(
    filePath = filePath,
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
  assertThat(as.integer(df[,11]), equalTo(mtcars$carb))
}

test.importWithHeaderNamesExcel <- function() {
  importWithHeaderNames(wdfile("df.xlsx"))
}
test.importWithHeaderNamesOds <- function() {
  importWithHeaderNames(wdfile("df.ods"))
}

importComplex <- function(filePath) {
  df <- importSpreadsheet(
    filePath = filePath,
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
  assertThat(as.columnIndex("N"), equalTo(14))
  assertThat(as.columnIndex("AF"), equalTo(32))
  assertThat(as.columnIndex("AAB"), equalTo(704))

  assertThat(as.columnName(14), equalTo("N"))
  assertThat(as.columnName(32), equalTo("AF"))
  assertThat(as.columnName(704), equalTo("AAB"))
}

exportNew <- function(filePath) {
  if (file.exists(filePath)) file.remove(filePath)
  result <- exportSpreadsheet(filePath, mtcars)
  assertThat(result, equalTo(TRUE))
  gearCol <- findColumnNumber(filePath, 1, 1, "gear")
  expected <- as.columnIndex("J")
  assertThat(gearCol, equalTo(expected))
}

test.exportNewExcel <- function() {
  exportNew(wdfile("test.xlsx"))
}

test.exportNewOds <- function() {
  exportNew(wdfile("test.ods"))
}

update <- function(filePath) {
  if (file.exists(filePath)) file.remove(filePath)
  result <- exportSpreadsheet(filePath, mtcars)
  assertThat(result, equalTo(TRUE))
  result <- exportSpreadsheet(filePath, iris, "iris")
  assertThat(result, equalTo(TRUE))
  gearCol <- findColumnNumber(filePath, 1, 1, "gear")
  assertThat(gearCol, equalTo(as.columnIndex("J")))
  versicolorRow <- findRowNumber(filePath, "iris", as.columnIndex("E") , "versicolor")
  assertThat(versicolorRow, equalTo(52))
}

test.updateExcel <- function() {
  update(wdfile("test2.xlsx"))
}

test.updateOds <- function() {
  update(wdfile("test2.ods"))
}

exportSpreadsheetsTest <- function(filePath) {
  if (file.exists(filePath)) file.remove(filePath)
  sheetNames <- c("mtcars", "iris", "PlantGrowth")
  result <- exportSpreadsheets(filePath, list(mtcars, iris, PlantGrowth), sheetNames)
  assertThat(result, equalTo(TRUE))
  assertThat(file.exists(filePath), equalTo(TRUE))
  sheets <- getSheetNames(filePath)
  assertThat(sheetNames, equalTo(sheets))
}

test.exportSpreadsheetsExcel <- function() {
  exportSpreadsheetsTest(wdfile("testSheets.xlsx"))
}

test.exportSpreadsheetsOds <- function() {
  exportSpreadsheetsTest(wdfile("testSheets.ods"))
}

test.exportSpreadsheetRainy <- function() {
  assertThat(exportSpreadsheets("filePath.xlsx", NA, sheetNames), throwsError())

  assertThat(exportSpreadsheets("filePath.ods", list(mtcars), NULL), throwsError())
}
