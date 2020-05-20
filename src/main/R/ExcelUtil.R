library(R6)

ExcelUtil <- R6Class(
  "ExcelUtil",

  private = list(
    m_filePath = NULL,
    m_util = NULL
  ),

  public = list(

    initialize = function(path) {
      if (is.na(path) | is.null(path)) {
        stop(paste("Argument must be a string path but was", path))
      }
      if (!file.exists(path)) {
        stop(paste(path, "does not exist"))
      }
      private$m_filePath <- path
      private$m_util <- ExcelReader$new()
      private$m_util$setExcel(private$m_filePath)
    },

    findRowNum = function(sheetNumber, column, cellContent) {
      #import(se.alipsa.excelutils.ExcelReader)
      #jUtil <- ExcelReader$new()$setExcel(m_filePath)
      private$m_util$findRowNum(as.integer(sheetNumber), as.integer(column), cellContent)
    },

    close = function() {
      private$m_util$close()
    },

    finalize = function() {
      private$m_util$close()
    }
  )
)
