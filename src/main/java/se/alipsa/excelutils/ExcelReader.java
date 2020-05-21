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

   public static int findRowNum(String filePath, int sheetNumber, int colNumber, String content) throws Exception {
      try {
         setExcel(filePath);
         Sheet sheet = workbook.getSheetAt(sheetNumber);
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

   private static int findRowNum(Sheet sheet, int colNumber, String content) {
      Iterator<Row> it = sheet.rowIterator();
      ValueExtractor ext = new ValueExtractor(sheet);
      int rowCount = 0;
      while (it.hasNext()) {
         rowCount++;
         Row row = it.next();
         Cell cell = row.getCell(colNumber);
         if (content.equals(ext.getString(cell))) {
            return rowCount;
         }
      }
      return -1;
   }


   public static int findColNum(String filePath, int sheetNumber, int rowNumber, String content) throws Exception {
      try {
         setExcel(filePath);
         Sheet sheet = workbook.getSheetAt(sheetNumber);
         int colNum = findColNum(sheet, rowNumber, content);
         close();
         return colNum;
      } catch (Exception e) {
         if (workbook != null) workbook.close();
         throw e;
      }
   }

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

   private static int findColNum(Sheet sheet, int rowNumber, String content) {
      if (content==null) return -1;
      ValueExtractor ext = new ValueExtractor(sheet);
      Row row = sheet.getRow(rowNumber);
      int colNum = 0;
      for (Iterator<Cell> iter = row.cellIterator(); iter.hasNext(); ) {
         Cell cell = iter.next();
         if (content.equals(ext.getString(cell))) {
            return colNum;
         }
         colNum++;
      }
      return -1;
   }
}
