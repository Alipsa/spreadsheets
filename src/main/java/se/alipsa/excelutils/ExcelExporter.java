package se.alipsa.excelutils;

import org.apache.poi.ss.usermodel.*;
import org.renjin.primitives.Types;
import org.renjin.sexp.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.util.Iterator;
import java.util.Map;

/**
 * Export data.frame(s) (ListVector) to an Excel file.
 */
public class ExcelExporter {

   private static final Logger logger = LoggerFactory.getLogger(ExcelExporter.class);

   private ExcelExporter() {
      // prevent instantiation
   }

   /**
    * Create a new excel file.
    *
    * @param dataFrame the data.frame to export
    * @param filePath the file path + file name of the file to export to. Should end with one of .xls, .xlsx
    * @return true if successful, false if not written (e.g. file cannot be written to)
    */
   public static boolean exportExcel(String filePath, ListVector dataFrame) {
      File file = new File(filePath);
      if (file.exists()) {
         logger.info("File {} already exists, file length is {} kb", file.getAbsolutePath(), file.length()/1024 );
      }
      String lcFilePath = filePath.toLowerCase();

      if (!(lcFilePath.endsWith("xls") || lcFilePath.endsWith("xlsx"))) {
         logger.warn("Non typical extension detected (file neither ends with .xls or .xlsx), so will save as xlsx format");
      }

      Workbook workbook;
      try {
         FileInputStream fis = null;
         if (file.exists()) {
            fis = new FileInputStream(file);
            workbook = WorkbookFactory.create(fis);
         } else {
            workbook = WorkbookFactory.create(isXssf(lcFilePath));
         }

         Sheet sheet = workbook.createSheet();
         buildSheet(dataFrame, sheet);
         if (fis != null) {
            fis.close();
         }
         return writeFile(file, workbook);
      } catch (IOException e) {
         logger.error("Failed to create excel file {}" + file.getAbsolutePath(), e);
         return false;
      }
   }

   private static boolean writeFile(File file, Workbook workbook) throws IOException {
      if (workbook == null) {
         logger.warn("Workbook is null, cannot write to file");
         return false;
      }
      logger.info("Writing spreadsheet to {}", file.getAbsolutePath());
      try(FileOutputStream fos = new FileOutputStream(file)) {
         workbook.write(fos);
      } finally {
         try {
            workbook.close();
         } catch (IOException e) {
            e.printStackTrace();
         }
      }
      if (!file.exists()) {
         logger.warn("Failed to write to file");
         return false;
      }
      return true;
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
    * @return true if successful, false if not written (e.g. file cannot be written to)
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
         return writeFile(file, workbook);
      } catch (IOException e) {
         logger.error("Failed to create excel file {}" + file.getAbsolutePath(), e);
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
