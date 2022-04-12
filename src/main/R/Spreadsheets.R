
#' @param filePath The path to the excel file
#' @param sheet Either the sheet index OR the sheet name
#' @param column The column index or name (eg. A) for the column to search in
#' @param cellContent The content to search for
findRowNumber <- function(filePath, sheet = 1, column, cellContent) {
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
    if (endsWith(tolower(filePath), ".ods")) {
      rowNum <- OdsReader$findRowNum(filePath, sheet, column, cellContent)
    } else {
      rowNum <- ExcelReader$findRowNum(filePath, sheet, column, cellContent)
    }
    return(rowNum)
  },
  error = function(cond) {
    stop(cond);
  })
}

#' @param filePath The path to the excel file
#' @param sheet Either the sheet index OR the sheet name
#' @param row The row index for the row to search in
#' @param cellContent The content to search for
findColumnNumber <- function(filePath, sheet = 1, row, cellContent) {
  tryCatch({
    if (!(is.numeric(sheet) | is.character(sheet))) {
      stop("sheet parameter must either be an index or a string corresponding to the sheet")
    }
    if (is.numeric(sheet)) {
      sheet <- as.integer(sheet)
    }
    rowNum <- as.integer(row)
    if (endsWith(tolower(filePath), ".ods")) {
      colNum <- OdsReader$findColNum(filePath, sheet, rowNum, cellContent)
    } else {
      colNum <- ExcelReader$findColNum(filePath, sheet, rowNum, cellContent)
    }
    return(colNum)
  },
    error = function(cond) {
    stop(cond);
  })
}

as.columnIndex <- function(columnName) {
  if (!is.character(columnName)) {
    stop("columnName parameter must be a character string")
  }
  SpreadsheetUtil$asColumnNumber(columnName)
}

as.columnName <- function(columnIndex) {
  if (!is.numeric(columnIndex)) {
    stop("columnIndex parameter must be a number")
  }
  SpreadsheetUtil$asColumnName(as.integer(columnIndex))
}

getSheetNames <- function(filePath) {
  if (endsWith(tolower(filePath), ".ods")) {
    return(OdsReader$getSheetNames(filePath))
  } else {
    return(ExcelReader$getSheetNames(filePath))
  }
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

  if (endsWith(tolower(filePath), ".ods")) {
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

# Example usage
# sheets <- importSpreadsheets(
#     filePath=paste0(getwd(), "/results.xslx"),
#     sheets = c("Sheet1","Sheet2","Sheet3"),
#     importAreas = list(
#         "Sheet1"=c(1, 1212, 1, 17),
#         "Sheet2"=c(5,123,2,10),
#         "Sheet3"=c(1,345, 2, 34)
#     ), firstRowAsColumnNames = list(
#         "Sheet1"=FALSE,
#         "Sheet2"=TRUE,
#         "Sheet3"=TRUE
#     )
# )
# sheet2 <- sheets[["sheet2"]]
importSpreadsheets <- function(filePath, sheets, importAreas, firstRowAsColumnNames) {
  printUsage <- function() {
    cat(paste("Example usage",
      "sheets <- importSpreadsheets(",
                "    filePath=paste0(getwd(), '/results.xslx'),",
                "    sheets = c('Sheet1', 'Sheet2', 'Sheet3'),",
                "    importAreas = list(",
                "        'Sheet1' = c(1, 1212, 1, 17),",
                "        'Sheet2' = c(5, 123, 2, 10),",
                "        'Sheet3' = c(1, 345, 2, 34)",
                "    ),",
                "    firstRowAsColumnNames = list(",
                "        'Sheet1' = FALSE,",
                "        'Sheet2' = TRUE,",
                "        'Sheet3' = TRUE",
                "    )",
                ")", sep="\n"))
  }
  if (!file.exists(filePath)) {
    printUsage()
    stop(paste0("filePath (", filePath, ") does not exist"))
  }
  if (!(is.list(sheets) | is.vector(sheets))) {
    printUsage()
    stop("sheets must be a vector or a list")
  }
  if (!is.character(sheets)) {
    printUsage()
    stop("sheets must be character vector or list")
  }
  if (!is.list(importAreas)) {
    printUsage()
    stop("importAreas must be a list of named vectors")
  }
  for (i in 1:length(importAreas) ) {
    if (!is.vector(importAreas[[i]], mode = "numeric")) {
      printUsage()
      stop("non numeric vector detected in importAreas")
    }
    if (length(importAreas[[i]]) != 4) {
      printUsage()
      stop(paste0("numeric vector for index ", i, " (", names(importAreas)[i], ") is not of length 4"))
    }
  }
  if (!is.list(firstRowAsColumnNames)) {
    printUsage()
    stop("firstRowAsColumnNames must be a list of name / logical values pairs")
  }

  if (!all(sapply(firstRowAsColumnNames, is.logical))) {
    printUsage()
    stop("firstRowAsColumnNames non logical value detected in the named list")
  }

  if(!all.equal(sheets, names(importAreas), names(firstRowAsColumnNames))) {
    printUsage()
    stop("sheet names, names(importAreas) and names(firstRowAsColumnNames) does not match")
  }

  if (endsWith(tolower(filePath),".ods")) {
    return(OdsImporter$importOdsSheets(filePath, sheets, importAreas, firstRowAsColumnNames))
  } else {
    return(ExcelImporter$importExcelSheets(filePath, sheets, importAreas, firstRowAsColumnNames))
  }
}


exportSpreadsheet <- function(filePath, df, sheet = NA) {
  if (!dir.exists(dirname(filePath))) {
    stop(paste(dirname(filePath), "does not exists, create it first before exporting a file there!"))
  }
  if (endsWith(tolower(filePath), ".ods")) {
    result <- exportOds(filePath, df, sheet)
  } else {
    result <- exportExcel(filePath, df, sheet)
  }
  if (result == FALSE) {
    warning(paste("Failed to export spreadsheet to", filePath))
  }
  result
}

exportExcel <- function(filePath, df, sheet = NA) {
  if (is.na(sheet)) {
    return(ExcelExporter$exportExcel(filePath, df))
  } else {
    return(ExcelExporter$exportExcel(filePath, df, sheet))
  }
}

exportOds <- function(filePath, df, sheet = NA) {
  if (is.na(sheet)) {
    return(OdsExporter$exportOds(filePath, df))
  } else {
    return(OdsExporter$exportOds(filePath, df, sheet))
  }
}

exportSpreadsheets <- function(filePath, dfList, sheetNames ) {

  if (class(filePath) == "NULL" | class(filePath) == "NA") {
    stop("filePath must be specified")
  }

  if (class(dfList) == "NULL" | class(dfList) == "NA") {
    stop("dfList must be specified")
  }

  if (class(sheetNames) == "NULL" | class(sheetNames) == "NA") {
    stop("sheetNames not specified")
  }

  if (!is.list(dfList)) {
    stop(paste("dfList must be a list of data.frames but is", class(dfList)))
  }

  if (!is.character(filePath)) {
    stop(paste("filePath must be a character string but is ", class(filePath)))
  }

  if (!is.character(sheetNames)) {
    stop(paste("sheetNames must be a vector of character strings but is ", class(sheetNames)))
  }

  if (length(dfList) != length(sheetNames)) {
    stop(paste("You need to supply a name for each data.frame. Number of data.frames =", length(dfList),
               ", number of sheet names =", length(sheetNames)))
  }

  if (!dir.exists(dirname(filePath))) {
    stop(paste(dirname(filePath), "does not exists, create it first before exporting a file there!"))
  }

  if (endsWith(tolower(filePath), ".ods")) {
    result <- OdsExporter$exportOdsSheets(filePath, dfList, sheetNames)
  } else {
    result <- ExcelExporter$exportExcelSheets(filePath, dfList, sheetNames)
  }
  if (result == FALSE) {
    warning(paste("Failed to export spreadsheet to", filePath))
  }
  result
}

