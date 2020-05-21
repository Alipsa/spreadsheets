# remember to add export(function name) to NAMESPACE to make them available

findRowNumber <- function(fileName, sheet = 1, column, cellContent) {
  tryCatch({
    sheetNum <- as.integer(sheet) - 1L
    colNum <- as.integer(column) - 1L
    rowNum <- ExcelReader$findRowNum(fileName, sheetNum, colNum, cellContent)
    return(rowNum + 1)
  },
  error = function(cond) {
    stop(cond);
  })
}

findColumnNumber <- function(fileName, sheet = 1, row, cellContent) {
  tryCatch({
    sheetNum <- as.integer(sheet) - 1L
    rowNum <- as.integer(row) - 1L
    colNum <- ExcelReader$findColNum(fileName, sheetNum, rowNum, cellContent)
    return(colNum + 1)
  },
    error = function(cond) {
    stop(cond);
  })
}

importExcel <- function(filePath, sheet = 1, startRow = 1, endRow, startColumn = 1, endColumn,
                        firstRowAsColumnNames = FALSE, columnNames = NA) {
  if (firstRowAsColumnNames == TRUE & !is.na(columnNames)) {
    stop("Column names are defined but firstRowAsColumnNames is set to TRUE")
  }
  if (!is.na(columnNames) & !(is.list(columnNames) | is.vector(columnNames))) {
    stop("columnNames must be a vector or a list")
  }
  if (is.vector(columnNames)) {
    columnNames <- as.list(columnNames)
  }
  if (startRow > endRow) {
    stop("wrong arguments: startRow > endRow")
  }
  if (startColumn > endColumn) {
    stop("wrong arguments: startColumn > endColumn")
  }
  sheetNum <- as.integer(sheet) - 1L
  startRowNum <- as.integer(startRow) - 1L
  endRowNum <- as.integer(endRow) -1L
  startColNum <- as.integer(startColumn) - 1L
  endColNum <- as.integer(endColumn) - 1L
  if (is.na(columnNames)) {
    excelDf <- ExcelImporter$importExcel(
      filePath,
      sheetNum,
      startRowNum,
      endRowNum,
      startColNum,
      endColNum,
      firstRowAsColumnNames
    )
    return(excelDf)
  } else {
    excelDf <- ExcelImporter$importExcel(
      filePath,
      sheetNum,
      startRowNum,
      endRowNum,
      startColNum,
      endColNum,
      columnNames
    )
  }
}