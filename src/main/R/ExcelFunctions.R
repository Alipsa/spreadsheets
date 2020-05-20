# remember to add export(function name) to NAMESPACE to make them available

findRowNumber <- function(fileName, sheetNumber, column, cellContent) {
  tryCatch({
    #import(se.alipsa.excelutils.ExcelReader)
    util <- ExcelReader$new()
    rowNum <- util$findRowNum(fileName, as.integer(sheetNumber), as.integer(column), cellContent)
    return(rowNum)
  },
  error = function(cond) {
    stop(cond);
  })
}

findColumnNumber <- function(fileName, sheetNumber, row, cellContent) {
  tryCatch({
    util <- ExcelReader$new()
    colNum <- util$findColNum(fileName, as.integer(sheetNumber), as.integer(row), cellContent)
    return(colNum)
  },
    error = function(cond) {
    stop(cond);
  })
}

importExcel <- function(filePath, sheetNumber = 0, startRowNum = 0, endRowNum, startColNum = 0, endColNum,
                        firstRowAsColNames = FALSE, columnNames = NA) {
  if (firstRowAsColNames == TRUE & !is.na(columnNames)) {
    stop("Column names are defined but firstRowAsColNames is set to TRUE")
  }
  if (!is.na(columnNames) & !(is.list(columnNames) | is.vector(columnNames))) {
    stop("columnNames must be a vector or a list")
  }
  if (is.vector(columnNames)) {
    columnNames <- as.list(columnNames)
  }
  if (startRowNum > endRowNum) {
    stop("wrong arguments: startRowNum > endRowNum")
  }
  if (startColNum > endColNum) {
    stop("wrong arguments: startColNum > endColNum")
  }
  if (is.na(columnNames)) {
    excelDf <- ExcelImporter$importExcel(
      filePath,
      as.integer(sheetNumber),
      as.integer(startRowNum),
      as.integer(endRowNum),
      as.integer(startColNum),
      as.integer(endColNum),
      firstRowAsColNames
    )
    return(excelDf)
  } else {
    excelDf <- ExcelImporter$importExcel(
      filePath,
      as.integer(sheetNumber),
      as.integer(startRowNum),
      as.integer(endRowNum),
      as.integer(startColNum),
      as.integer(endColNum),
      columnNames
    )
  }
}

# Not sure why this is working here but in findRowNum i have to catch and rethrow
barf <- function(msg) {
  import(se.alipsa.excelutils.Barfer)
  barfer <- Barfer$new()
  barfer$barf(msg)
}