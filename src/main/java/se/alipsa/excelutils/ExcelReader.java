package se.alipsa.excelutils;

import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.IOException;
import java.util.Iterator;

import static se.alipsa.excelutils.FileUtil.checkFilePath;

public class ExcelReader {

   static FormulaEvaluator evaluator;
   static DataFormatter formatter = new DataFormatter();
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
         int rowNum = findRowNum(sheetNumber, colNumber, content);
         close();
         return rowNum;
      } catch (Exception e) {
         if (workbook != null) close();
         throw e;
      }
   }

   private static int findRowNum(int sheetNumber, int colNumber, String content) {
      Sheet sheet = workbook.getSheetAt(sheetNumber);
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
         if (workbook == null) {
            setExcel(filePath);
         } else {
            System.err.println("findColNum: workbook is not null, please do not mix OO and functional methods");
         }
         int colNum = findColNum(sheetNumber, rowNumber, content);
         close();
         return colNum;
      } catch (Exception e) {
         if (workbook != null) workbook.close();
         throw e;
      }
   }

   private static int findColNum(int sheetNumber, int rowNumber, String content) {
      if (content==null) return -1;
      Sheet sheet = workbook.getSheetAt(sheetNumber);
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
