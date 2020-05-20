library('hamcrest')
library('se.alipsa:excel-utils')

test.findRowNumSunny <- function() {
  util <- ExcelUtil$new("df.xlsx")
  rowNum <- util$findRowNum(0, 0, "Iris")
  util$close()
  assertThat(rowNum, equalTo(35))
}

test.finRowNumRainy <- function() {
  tryCatch({
    util <- ExcelUtil$new("doesnotexist.xlsx")
    util$findRowNum(0, 0, "Iris")
  },
  error = function(err) {
    #print(paste("Expected error was: ", err))
    assertTrue(endsWith(trimws(err$message), " not exist"))
  })
}

test.finRowNumCloudy <- function() {
  util <- ExcelUtil$new("df.xlsx")
  rowNum <- util$findRowNum(0, 0, "Nothing that exist")
  assertThat(rowNum, equalTo(-1))
  util$close()
}
