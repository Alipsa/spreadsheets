
#' @param fileName The path to the excel file
#' @param sheet Either the sheet index OR the sheet name
#' @param column The column index or name (eg. A) for the column to search in
#' @param cellContent The content to search for
findRowNumber <- function(fileName, sheet = 1, column, cellContent) {
  if (!(is.numeric(sheet) | is.character(sheet))) {
    stop("sheet parameter must either be an index or a string corresponding to the sheet")
  }

  tryCatch({
    if (is.numeric(sheet)) {
     sheet <- as.integer(sheet) - 1L
    }
    if (is.numeric(column)) {
     column <- as.integer(column) - 1L
    }
    rowNum <- ExcelReader$findRowNum(fileName, sheet, column, cellContent)
    return(rowNum + 1)
  },
  error = function(cond) {
    stop(cond);
  })
}

#' @param fileName The path to the excel file
#' @param sheet Either the sheet index OR the sheet name
#' @param row The row index for the row to search in
#' @param cellContent The content to search for
findColumnNumber <- function(fileName, sheet = 1, row, cellContent) {
  tryCatch({
    if (!(is.numeric(sheet) | is.character(sheet))) {
      stop("sheet parameter must either be an index or a string corresponding to the sheet")
    }
    if (is.numeric(sheet)) {
      sheet <- as.integer(sheet) - 1L
    }
    rowNum <- as.integer(row) - 1L
    colNum <- ExcelReader$findColNum(fileName, sheet, rowNum, cellContent)
    return(colNum + 1)
  },
    error = function(cond) {
    stop(cond);
  })
}

columnIndex <- function(columnName) {
  if (!is.character(columnName)) {
    stop("columnName parameter must be a character string")
  }
  ExcelUtil$toColumnNumber(columnName)
}

columnName <- function(columnIndex) {
  if (!is.numeric(columnIndex)) {
    stop("columnIndex parameter must be a number")
  }
  ExcelUtil$toColumnName(as.integer(columnIndex))
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
  if (is.numeric(sheet)) {
    sheet <- as.integer(sheet) - 1L
  }
  startRowNum <- as.integer(startRow) - 1L
  endRowNum <- as.integer(endRow) -1L
  if (is.numeric(startColumn)) {
    startColumn <- as.integer(startColumn) - 1L
  }
  if (is.numeric(endColumn)) {
    endColumn <- as.integer(endColumn) - 1L
  }
  if (is.na(columnNames)) {
    excelDf <- ExcelImporter$importExcel(
      filePath,
      sheet,
      startRowNum,
      endRowNum,
      startColumn,
      endColumn,
      firstRowAsColumnNames
    )
    return(excelDf)
  } else {
    excelDf <- ExcelImporter$importExcel(
      filePath,
      sheet,
      startRowNum,
      endRowNum,
      startColumn,
      endColumn,
      columnNames
    )
  }
}