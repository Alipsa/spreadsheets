
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
     sheet <- as.integer(sheet)
    }
    if (is.numeric(column)) {
     column <- as.integer(column)
    }
    if (endsWith(tolower(fileName), ".ods") | endsWith(tolower(fileName), ".ods")) {
      rowNum <- OdsReader$findRowNum(fileName, sheet, column, cellContent)
    } else {
      rowNum <- ExcelReader$findRowNum(fileName, sheet, column, cellContent)
    }
    return(rowNum)
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
      sheet <- as.integer(sheet)
    }
    rowNum <- as.integer(row)
    if (endsWith(tolower(fileName), ".ods") | endsWith(tolower(fileName), ".ods")) {
      colNum <- OdsReader$findColNum(fileName, sheet, rowNum, cellContent)
    } else {
      colNum <- ExcelReader$findColNum(fileName, sheet, rowNum, cellContent)
    }
    return(colNum)
  },
    error = function(cond) {
    stop(cond);
  })
}

columnIndex <- function(columnName) {
  if (!is.character(columnName)) {
    stop("columnName parameter must be a character string")
  }
  SpreadsheetUtil$toColumnNumber(columnName)
}

columnName <- function(columnIndex) {
  if (!is.numeric(columnIndex)) {
    stop("columnIndex parameter must be a number")
  }
  SpreadsheetUtil$toColumnName(as.integer(columnIndex))
}

importSpreadsheet <- function(filePath, sheet = 1, startRow = 1, endRow, startColumn = 1, endColumn,
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
    sheet <- as.integer(sheet)
  }
  startRow <- as.integer(startRow)
  endRow <- as.integer(endRow)
  if (is.numeric(startColumn)) {
    startColumn <- as.integer(startColumn)
  }
  if (is.numeric(endColumn)) {
    endColumn <- as.integer(endColumn)
  }

  if (endsWith(tolower(filePath), ".ods") | endsWith(tolower(filePath), ".ods")) {
    return(importOds(filePath = filePath, sheet = sheet, startRow = startRow, endRow = endRow, startColumn = startColumn,
                     endColumn = endColumn, firstRowAsColumnNames = firstRowAsColumnNames, columnNames = columnNames))
  } else {
    return(importExcel(filePath = filePath, sheet = sheet, startRow = startRow, endRow = endRow, startColumn = startColumn,
                       endColumn = endColumn, firstRowAsColumnNames = firstRowAsColumnNames, columnNames = columnNames))
  }
}

importExcel <- function(filePath, sheet = 1, startRow = 1, endRow, startColumn = 1, endColumn,
                        firstRowAsColumnNames = FALSE, columnNames = NA) {

  if (is.na(columnNames)) {
    excelDf <- ExcelImporter$importExcel(
      filePath,
      sheet,
      startRow,
      endRow,
      startColumn,
      endColumn,
      firstRowAsColumnNames
    )
    return(excelDf)
  } else {
    excelDf <- ExcelImporter$importExcel(
      filePath,
      sheet,
      startRow,
      endRow,
      startColumn,
      endColumn,
      columnNames
    )
    return(excelDf)
  }
}

importOds <- function(filePath, sheet = 1, startRow = 1, endRow, startColumn = 1, endColumn,
                      firstRowAsColumnNames = FALSE, columnNames = NA) {
  if (is.na(columnNames)) {
    odsDf <- OdsImporter$importOds(
      filePath,
      sheet,
      startRow,
      endRow,
      startColumn,
      endColumn,
      firstRowAsColumnNames
    )
    return(odsDf)
  } else {
    odsDf <- OdsImporter$importOds(
      filePath,
      sheet,
      startRow,
      endRow,
      startColumn,
      endColumn,
      columnNames
    )
    return(odsDf)
  }
}

exportSpreadsheet <- function(df, filePath, sheet = NA) {
  if (endsWith(tolower(filePath), ".ods") | endsWith(tolower(filePath), ".ods")) {
    exportOds(df, filePath, sheet)
  } else {
    exportExcel(df, filePath, sheet)
  }
}

exportExcel <- function(df, filePath, sheet = NA) {
  if (is.na(sheet)) {
    return(ExcelExporter$exportExcel(df, filePath))
  } else {
    return(ExcelExporter$exportExcel(df, filePath, sheet))
  }
}

exportOds <- function(df, filePath, sheet = NA) {
  if (is.na(sheet)) {
    return(OdsExporter$exportOds(df, filePath))
  } else {
    return(OdsExporter$exportOds(df, filePath, sheet))
  }
}
