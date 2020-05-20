package se.alipsa.excelutils;

import org.apache.poi.ss.usermodel.*;
import org.renjin.primitives.vector.RowNamesVector;
import org.renjin.sexp.*;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

import static org.apache.poi.ss.usermodel.CellType.BOOLEAN;
import static org.apache.poi.ss.usermodel.CellType.NUMERIC;
import static se.alipsa.excelutils.FileUtil.checkFilePath;

public class ExcelImporter {

   public static ListVector importExcel(String filePath, int sheetNumber, int startRowNum, int endRowNum, int startColNum, int endColNum, boolean firstRowAsColNames) throws Exception {
      File excelFile = checkFilePath(filePath);
      DataFormatter formatter = new DataFormatter();
      List<String> header = new ArrayList<>();
      if (firstRowAsColNames) {
         try (Workbook workbook = WorkbookFactory.create(excelFile)) {
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            Sheet sheet = workbook.getSheetAt(sheetNumber);
            Row row = sheet.getRow(startRowNum);
            for (int i = 0; i < endColNum - startColNum; i++) {
               Cell cell = row.getCell(startColNum + i);
               header.add(ExcelUtil.stringCellVal(cell, evaluator, formatter));
            }
         }
         startRowNum = startRowNum + 1;
      } else {
         for (int i = 1; i <= endColNum - startColNum; i++) {
            header.add(String.valueOf(i));
         }
      }
      return importExcel(excelFile, sheetNumber, startRowNum, endRowNum, startColNum, endColNum, header);
   }

   public static ListVector importExcel(String filePath, int sheetNumber, int startRowNum, int endRowNum,int startColNum, int endColNum, Vector colNames) throws Exception {
      File excelFile = checkFilePath(filePath);
      List<String> header = new ArrayList<>();
      for (int i = 0; i < colNames.length(); i++) {
         header.add(colNames.getElementAsString(i));
      }
      return importExcel(excelFile, sheetNumber, startRowNum, endRowNum, startColNum, endColNum, header);
   }

   private static ListVector importExcel(File excelFile, int sheetNumber, int startRowNum, int endRowNum, int startColNum, int endColNum, List<String> colNames) throws Exception {
      try (Workbook workbook = WorkbookFactory.create(excelFile)) {
         Sheet sheet = workbook.getSheetAt(sheetNumber);
         ValueExtractor ext = new ValueExtractor(sheet);
         List<StringVector.Builder> builders = stringBuilders(startColNum, endColNum);
         int numRows = 0;
         for (int rowIdx = startRowNum; rowIdx <= endRowNum; rowIdx++) {
            numRows++;
            Row row = sheet.getRow(rowIdx);
            int i = 0;
            for (int colIdx = startColNum; colIdx < endColNum; colIdx++) {
               builders.get(i++).add(ext.getString(row, colIdx));
            }
         }
         ListVector columnVector = columnInfo(colNames);
         /* call build() on each column and add them as named cols to df */
         ListVector.NamedBuilder dfBuilder = new ListVector.NamedBuilder();
         for (int i = 0; i < columnVector.length(); i++) {
            ListVector ci = (ListVector) columnVector.get(i);
            dfBuilder.add(ci.get("name").asString(), builders.get(i).build());
         }
         dfBuilder.setAttribute("row.names", new RowNamesVector(numRows));
         dfBuilder.setAttribute("class", StringVector.valueOf("data.frame"));
         return dfBuilder.build();
      }
   }

   private static List<StringVector.Builder> stringBuilders(int startColNum, int endColNum) {
      List<StringVector.Builder> builder = new ArrayList<>();
      for (int i = 0; i < endColNum - startColNum; i++) {
         builder.add(new StringVector.Builder());
      }
      return builder;
   }

   private static List<Vector.Builder<?>> guessBuilders(int startColNum, int endColNum, Row row) {
      List<Vector.Builder<?>> builder = new ArrayList<>();
      // Guess the column type based on the first value row
      for (int colIdx = startColNum; colIdx < endColNum; colIdx++) {
         Cell cell = row.getCell(colIdx);
         CellType cellType = cell.getCellType();
         if (cellType == BOOLEAN) {
            builder.add(new LogicalArrayVector.Builder());
         } else if (cellType == NUMERIC) {
            if (DateUtil.isCellDateFormatted(cell)) {
               // we have a date, store as string
               builder.add(new StringVector.Builder());
            } else if (cell.getCellStyle().getDataFormatString().contains(".")) {
               // it is a double
               builder.add(new DoubleArrayVector.Builder());
            } else {
               // it is an int
               builder.add(new IntArrayVector.Builder());
            }
         } else {
            // String and everything else
            builder.add(new StringVector.Builder());
         }
      }
      return builder;
   }

   public static ListVector columnInfo(List<String> colNames) {
      ListVector.Builder tv = new ListVector.Builder();
      for (String name : colNames) {
         ListVector.NamedBuilder cv = new ListVector.NamedBuilder();
         cv.add("name", name);
         tv.add(cv.build());
      }

      return tv.build();
   }
}
