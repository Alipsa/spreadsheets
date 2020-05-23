package se.alipsa.excelutils;

import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.IOException;
import java.util.Iterator;

import static se.alipsa.excelutils.FileUtil.checkFilePath;

public class ExcelReader {

   static FormulaEvaluator evaluator;
   static Workbook workbook;

   private ExcelReader() {
      // Prevent instantiation
   }

   public static void setExcel(String filePath) throws Exception {
      File excelFile = checkFilePath(filePath);
      workbook = WorkbookFactory.create(excelFile);
      evaluator = workbook.getCreationHelper().createFormulaEvaluator();
   }

   public static void close() throws IOException {
      workbook.close();
      workbook = null;
      evaluator = null;
   }

   /**
    *
    * @param filePath the excel file
    * @param sheetNumber the sheet index (1 indexed)
    * @param colNumber the column number (1 indexed)
    * @param content the string to search for
    * @return the Row as seen in Excel (1 is first row)
    * @throws Exception if something goes wrong
    */
   public static int findRowNum(String filePath, int sheetNumber, int colNumber, String content) throws Exception {
      try {
         setExcel(filePath);
         Sheet sheet = workbook.getSheetAt(sheetNumber -1);
         int rowNum = findRowNum(sheet, colNumber, content);
         close();
         return rowNum;
      } catch (Exception e) {
         if (workbook != null) close();
         throw e;
      }
   }

   public static int findRowNum(String filePath, int sheetNumber, String colName, String content) throws Exception {
      return findRowNum(filePath, sheetNumber, ExcelUtil.toColumnNumber(colName), content);
   }

   public static int findRowNum(String filePath, String sheetName, String colName, String content) throws Exception {
      return findRowNum(filePath, sheetName, ExcelUtil.toColumnNumber(colName), content);
   }

   /**
    *
    * @param filePath the excel file
    * @param sheetName the name of the sheet
    * @param colNumber the column number (1 indexed)
    * @param content the string to search for
    * @return the Row as seen in Excel (1 is first row)
    * @throws Exception if something goes wrong
    */
   public static int findRowNum(String filePath, String sheetName, int colNumber, String content) throws Exception {
      try {
         setExcel(filePath);
         Sheet sheet = workbook.getSheet(sheetName);
         int rowNum = findRowNum(sheet, colNumber, content);
         close();
         return rowNum;
      } catch (Exception e) {
         if (workbook != null) close();
         throw e;
      }
   }

   /**
    *
    * @param sheet the sheet to search in
    * @param colNumber the column number (1 indexed)
    * @param content the string to search for
    * @return the Row as seen in Excel (1 is first row)
    */
   private static int findRowNum(Sheet sheet, int colNumber, String content) {
      ValueExtractor ext = new ValueExtractor(sheet);
      int poiColNum = colNumber -1;
      for (int rowCount = 0; rowCount < sheet.getLastRowNum(); rowCount ++) {
         Row row = sheet.getRow(rowCount);
         if (row == null) continue;
         Cell cell = row.getCell(poiColNum);
         //System.out.println(rowCount + ": " + ext.getString(cell));
         if (content.equals(ext.getString(cell))) {
            return rowCount + 1;
         }
      }
      return -1;
   }

   /**
    *
    * @param filePath the excel file
    * @param sheetNumber the sheet index (1 indexed)
    * @param rowNumber the row number (1 indexed)
    * @param content the string to search for
    * @return the row number that matched or -1 if not found
    * @throws Exception
    */
   public static int findColNum(String filePath, int sheetNumber, int rowNumber, String content) throws Exception {
      try {
         setExcel(filePath);
         Sheet sheet = workbook.getSheetAt(sheetNumber - 1);
         int colNum = findColNum(sheet, rowNumber, content);
         close();
         return colNum;
      } catch (Exception e) {
         if (workbook != null) workbook.close();
         throw e;
      }
   }

   /** return the column as seen in excel (e.g. using column(), 1 is the first column etc
    * @param filePath the excel file
    * @param sheetName the name of the sheet
    * @param rowNumber the row number (1 indexed)
    * @param content the string to search for
    */
   public static int findColNum(String filePath, String sheetName, int rowNumber, String content) throws Exception {
      try {
         setExcel(filePath);
         Sheet sheet = workbook.getSheet(sheetName);
         int colNum = findColNum(sheet, rowNumber, content);
         close();
         return colNum;
      } catch (Exception e) {
         if (workbook != null) workbook.close();
         throw e;
      }
   }

   /**
    * @param sheet the Sheet to search
    * @param rowNumber the row number (1 indexed)
    * @param content the string to search for
    * @return return the column as seen in excel (e.g. using column(), 1 is the first column etc
    */
   private static int findColNum(Sheet sheet, int rowNumber, String content) {
      if (content==null) return -1;
      ValueExtractor ext = new ValueExtractor(sheet);
      int poiRowNum = rowNumber - 1;
      Row row = sheet.getRow(poiRowNum);
      for (int colNum = 0; colNum < row.getLastCellNum(); colNum++) {
         Cell cell = row.getCell(colNum);
         if (content.equals(ext.getString(cell))) {
            return colNum + 1;
         }
      }
      return -1;
   }
}
