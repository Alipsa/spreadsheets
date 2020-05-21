package se.alipsa.excelutils;

import org.apache.poi.ss.usermodel.*;
import org.renjin.primitives.Types;
import org.renjin.sexp.*;

import java.io.*;
import java.util.Iterator;
import java.util.Map;

public class ExcelExporter {

   // Create a new excel file
   public static boolean exportExcel(ListVector dataFrame, String filePath, boolean overWrite) {
      File file = new File(filePath);
      if (!overWrite && file.exists()) {
         System.err.println("Overwrite is false and file already exists");
         return false;
      }
      String lcFilePath = filePath.toLowerCase();

      if (!(lcFilePath.endsWith("xls") || lcFilePath.endsWith("xlsx"))) {
         System.err.println("Non typical extension detected, will save as xlsx format");
      }

      boolean asXssf = true;

      if (lcFilePath.endsWith(".xls")) asXssf = false;

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


   // Update and existing excel
   public static boolean exportExcel(ListVector dataFrame, String filePath, int sheetIndex) {
      File file = new File(filePath);
      if (!file.exists()) {
         System.err.println("No excel file name " + filePath + " exists");
         return false;
      }

      try {
         FileInputStream fis = new FileInputStream(file);
         Workbook workbook = WorkbookFactory.create(fis);
         Sheet sheet = workbook.getSheetAt(sheetIndex);
         if (sheet == null) {
            sheet = workbook.createSheet();
         }
         buildSheet(dataFrame, sheet);
         fis.close();
         try(FileOutputStream fos = new FileOutputStream(file)) {
            workbook.write(fos);
         }
         workbook.close();
         return true;
      } catch (IOException e) {
         System.err.println("Failed to create excel file: " + e.toString());
         e.printStackTrace();
         return false;
      }
   }

   // Create new or update existing excel, adding or updating a sheet with the name specified
   public static boolean exportExcel(ListVector dataFrame, String filePath, String sheetName, boolean overWrite) {
      File file = new File(filePath);

      if (!overWrite && file.exists()) {
         System.err.println("Overwrite is false and file already exists");
         return false;
      }
      try {
         Workbook workbook;
         FileInputStream fis = null;
         if (file.exists()) {
            fis = new FileInputStream(file);
            workbook = WorkbookFactory.create(fis);
         } else {
            workbook = WorkbookFactory.create(file);
         }

         Sheet sheet = workbook.getSheet(sheetName);
         if (sheet == null) {
            sheet = workbook.createSheet(sheetName);
         }
         buildSheet(dataFrame, sheet);
         if (fis != null) {
            fis.close();
         }
         try(FileOutputStream fos = new FileOutputStream(file)) {
            workbook.write(fos);
         }
         workbook.close();
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
