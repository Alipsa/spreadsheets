package se.alipsa.excelutils;

import org.apache.poi.ss.usermodel.*;
import org.renjin.primitives.Types;
import org.renjin.sexp.*;

import java.io.*;
import java.util.Iterator;
import java.util.Map;

/**
 * Export data.frame(s) (ListVector) to an Excel file.
 */
public class ExcelExporter {

   private ExcelExporter() {
      // prevent instantiation
   }

   /**
    * Create a new excel file.
    *
    * @param dataFrame the data.frame to export
    * @param filePath the file path + file name of the file to export to. Should end with one of .xls, .xlsx, .ods
    * @return true if successful, false if not written (file exists or cannot be written to)
    */
   public static boolean exportExcel(String filePath, ListVector dataFrame) {
      File file = new File(filePath);
      if (file.exists()) {
         System.err.println("Overwrite is false and file already exists");
         return false;
      }
      String lcFilePath = filePath.toLowerCase();

      if (!(lcFilePath.endsWith("xls") || lcFilePath.endsWith("xlsx"))) {
         System.err.println("Non typical extension detected, will save as xlsx format");
      }

      boolean asXssf = isXssf(lcFilePath);

      try(Workbook workbook = WorkbookFactory.create(asXssf)) {
         Sheet sheet = workbook.createSheet();
         buildSheet(dataFrame, sheet);
         try(FileOutputStream fos = new FileOutputStream(file)) {
            workbook.write(fos);
         }
         return true;
      } catch (IOException e) {
         System.err.println("Failed to create excel file: " + e.toString());
         e.printStackTrace();
         return false;
      }
   }

   private static boolean isXssf(String filePath) {
      String lcFilePath = filePath.toLowerCase();

      return !lcFilePath.endsWith(".xls");

   }

   /**
    * upsert: Create new or update existing excel, adding or updating a sheet with the name specified
    *
    * @param dataFrame the data.frame to export
    * @param filePath the file path + file name of the file to export to. Should end with one of .xls, .xlsx, .ods
    * @param sheetName the name of the sheet to write to
    * @return true if successful, false if not written (file exists or cannot be written to)
    */
   public static boolean exportExcel(String filePath, ListVector dataFrame, String sheetName) {
      return exportExcelSheets(filePath, new ListVector(dataFrame), new StringArrayVector(sheetName));
   }

   private static void upsertSheet(ListVector dataFrame, String sheetName, Workbook workbook) {
      Sheet sheet = workbook.getSheet(sheetName);
      if (sheet == null) {
         sheet = workbook.createSheet(sheetName);
      }
      buildSheet(dataFrame, sheet);
   }

   public static boolean exportExcelSheets(String filePath, ListVector dataFrames, StringArrayVector sheetNames) {
      File file = new File(filePath);

      try {
         Workbook workbook;
         FileInputStream fis = null;
         if (file.exists()) {
            fis = new FileInputStream(file);
            workbook = WorkbookFactory.create(fis);
         } else {
            workbook = WorkbookFactory.create(isXssf(filePath));
         }

         for (int i = 0; i < dataFrames.length(); i++) {
            ListVector dataFrame = (ListVector)dataFrames.get(i);
            String sheetName = sheetNames.toArray()[i];
            //System.out.println("Writing sheet " + sheetName + " with a dataframe with "
            //   + dataFrame.maxElementLength() + " rows and " + dataFrame.length() + " columns");
            upsertSheet(dataFrame, sheetName, workbook);
         }

         if (fis != null) {
            fis.close();
         }
         try(FileOutputStream fos = new FileOutputStream(file)) {
            workbook.write(fos);
         }
         workbook.close();
         //System.out.println(file.getAbsolutePath() + " created");
         return true;
      } catch (IOException e) {
         System.err.println("Failed to create excel file: " + e.toString());
         e.printStackTrace();
         return false;
      }
   }

   private static void buildSheet(ListVector dataFrame, Sheet sheet) {

      AtomicVector names = dataFrame.getNames();
      Row headerRow = sheet.createRow(0);
      for (int i = 0; i < names.length(); i++) {
         headerRow.createCell(i).setCellValue(names.getElementAsString(i));
      }

      Iterator<SEXP> it = dataFrame.iterator();
      int colIdx = 0;
      while (it.hasNext()) {
         Vector colVec = (Vector)it.next();
         String typeName = colVec.getTypeName();
         for (int i = 0; i < colVec.length(); i++) {
            int excelRow = i + 1;
            if (Types.isFactor(colVec)) {
               AttributeMap attributes = colVec.getAttributes();
               Map<Symbol, SEXP> attrMap = attributes.toMap();
               Symbol s = attrMap.keySet().stream().filter(p -> "levels".equals(p.getPrintName())).findAny().orElse(null);
               Vector vec = (Vector) attrMap.get(s);
               Row row = sheet.getRow(excelRow);
               if (row == null) row = sheet.createRow(excelRow);
               row.createCell(colIdx, CellType.STRING).setCellValue(
                  vec.getElementAsString(colVec.getElementAsInt(i) - 1));

            } else {
               Row row = sheet.getRow(excelRow);
               if (row == null) row = sheet.createRow(excelRow);
               Cell cell = row.createCell(colIdx);
               if ("double".equals(typeName)) {
                  cell.setCellValue(colVec.getElementAsDouble(i));
               } else if ("integer".equals(typeName)) {
                  cell.setCellValue(colVec.getElementAsInt(i));
               } else if ("logical".equals(typeName)) {
                  cell.setCellValue(colVec.getElementAsLogical(i).toBooleanStrict());
               } else {
                  cell.setCellValue(colVec.getElementAsString(i));
               }
            }
         }
         colIdx++;
      }
   }
}
