package se.alipsa.excelutils;

import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.IOException;
import java.util.Iterator;

import static se.alipsa.excelutils.FileUtil.checkFilePath;

public class ExcelReader {

   FormulaEvaluator evaluator;
   static DataFormatter formatter = new DataFormatter();
   Workbook workbook;

   public ExcelReader() {
   }

   public ExcelReader setExcel(String filePath) throws Exception {
      File excelFile = checkFilePath(filePath);
      workbook = WorkbookFactory.create(excelFile);
      evaluator = workbook.getCreationHelper().createFormulaEvaluator();
      return this;
   }

   public void close() throws IOException {
      workbook.close();
   }

   public int findRowNum(String filePath, int sheetNumber, int colNumber, String content) throws Exception {
      try {
         if (workbook == null) {
            setExcel(filePath);
         } else {
            System.err.println("findRowNum: workbook is not null, please do not mix OO and functional methods");
         }
         int rowNum = findRowNum(sheetNumber, colNumber, content);
         close();
         return rowNum;
      } catch (Exception e) {
         if (workbook != null) close();
         throw e;
      }
   }

   public int findRowNum(int sheetNumber, int colNumber, String content) {
      Sheet sheet = workbook.getSheetAt(sheetNumber);
      Iterator<Row> it = sheet.rowIterator();
      int rowCount = 0;
      while (it.hasNext()) {
         rowCount++;
         Row row = it.next();
         Cell cell = row.getCell(colNumber);
         if (content.equals(cellVal(cell))) {
            return rowCount;
         }
      }
      return -1;
   }


   public int findColNum(String filePath, int sheetNumber, int rowNumber, String content) throws Exception {
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

   private int findColNum(int sheetNumber, int rowNumber, String content) {
      if (content==null) return -1;
      Sheet sheet = workbook.getSheetAt(sheetNumber);
      Row row = sheet.getRow(rowNumber);
      int colNum = 0;
      for (Iterator<Cell> iter = row.cellIterator(); iter.hasNext(); ) {
         Cell cell = iter.next();
         if (content.equals(cellVal(cell))) {
            return colNum;
         }
         colNum++;
      }
      return -1;
   }

   public String cellVal(Cell cell) {
      return ExcelUtil.stringCellVal(cell, evaluator, formatter);
   }
  
}
