package se.alipsa.excelutils;

import com.github.miachm.sods.Sheet;
import com.github.miachm.sods.SpreadSheet;
import org.renjin.primitives.Types;
import org.renjin.sexp.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.Map;

public class OdsExporter {

   /**
    * Create a new Open Document Spreadsheet file.
    *
    * @param dataFrame the data.frame to export
    * @param filePath the file path + file name of the file to export to. Should end with one of .xls, .xlsx, .ods
    * @return true if successful, false if not written (file exists or cannotbe written to)
    */
   public static boolean exportOds(String filePath, ListVector dataFrame) {
      File file = new File(filePath);
      if (file.exists()) {
         System.err.println("Overwrite is false and file already exists");
         return false;
      }

      try {
         SpreadSheet spreadSheet = new SpreadSheet();
         int nextIdx = spreadSheet.getNumSheets() + 1;
         Sheet sheet = new Sheet(String.valueOf(nextIdx), dataFrame.maxElementLength() + 1, dataFrame.length() + 1);
         spreadSheet.appendSheet(sheet);
         buildSheet(dataFrame, sheet);
         try (FileOutputStream fos = new FileOutputStream(file)) {
            spreadSheet.save(fos);
         }
         return true;
      } catch (IOException e) {
         e.printStackTrace();
         return false;
      }
   }

   /**
    * upsert: Create new or update existing Open Document Spreadsheet, adding or updating a sheet with the name specified
    *
    * @param dataFrame the data.frame to export
    * @param filePath the file path + file name of the file to export to. Should end with one of .xls, .xlsx, .ods
    * @param sheetName the name of the sheet to write to
    * @return true if successful, false if not written (file exists or cannotbe written to)
    */
   public static boolean exportOds(String filePath, ListVector dataFrame, String sheetName) {
      return exportOdsSheets(filePath, new ListVector(dataFrame), new StringArrayVector(sheetName));
   }

   public static boolean exportOdsSheets(String filePath, ListVector dataFrames, StringArrayVector sheetNames) {
      File file = new File(filePath);

      try {
         SpreadSheet spreadSheet;
         if (file.exists()) {
            spreadSheet = new SpreadSheet(file);
         } else {
            spreadSheet = new SpreadSheet();
         }

         for (int i = 0; i < dataFrames.length(); i++) {
            ListVector dataFrame = (ListVector)dataFrames.get(i);
            String sheetName = sheetNames.toArray()[i];
            //System.out.println("Writing sheet " + sheetName + " with a dataframe with "
            //   + dataFrame.maxElementLength() + " rows and " + dataFrame.length() + " columns");
            upsertSheet(dataFrame, sheetName, spreadSheet);
         }

         try(FileOutputStream fos = new FileOutputStream(file)) {
            spreadSheet.save(fos);
         }
         return true;
      } catch (IOException e) {
         System.err.println("Failed to create ods file: " + e.toString());
         e.printStackTrace();
         return false;
      }
   }

   private static void upsertSheet(ListVector dataFrame, String sheetName, SpreadSheet spreadSheet) {
      Sheet sheet = spreadSheet.getSheet(sheetName);
      if (sheet == null) {
         sheet = new Sheet(sheetName, dataFrame.maxElementLength() + 1, dataFrame.length() +1);
         spreadSheet.appendSheet(sheet);
      }
      buildSheet(dataFrame, sheet);
   }

   private static void buildSheet(ListVector dataFrame, Sheet sheet) {

      AtomicVector names = dataFrame.getNames();

      //Ensure there is enough space
      if (sheet.getLastColumn() < names.length()) {
         sheet.appendColumns(names.length() - sheet.getLastColumn());
      }
      if (sheet.getLastRow() < dataFrame.maxElementLength() + 1) {
         sheet.appendRows(dataFrame.maxElementLength() + 1 - sheet.getLastRow());
      }

      for (int i = 0; i < names.length(); i++) {
         sheet.getRange(0, i).setValue(names.getElementAsString(i));
      }

      Iterator<SEXP> it = dataFrame.iterator();
      int colIdx = 0;
      while (it.hasNext()) {
         Vector colVec = (Vector)it.next();
         String typeName = colVec.getTypeName();
         for (int i = 0; i < colVec.length(); i++) {
            int row = i + 1;
            if (Types.isFactor(colVec)) {
               AttributeMap attributes = colVec.getAttributes();
               Map<Symbol, SEXP> attrMap = attributes.toMap();
               Symbol s = attrMap.keySet().stream().filter(p -> "levels".equals(p.getPrintName())).findAny().orElse(null);
               Vector vec = (Vector) attrMap.get(s);
               sheet.getRange(row, colIdx).setValue(vec.getElementAsString(colVec.getElementAsInt(i) - 1));
            } else {
               if ("double".equals(typeName)) {
                  sheet.getRange(row, colIdx).setValue(colVec.getElementAsDouble(i));
               } else if ("integer".equals(typeName)) {
                  sheet.getRange(row, colIdx).setValue(colVec.getElementAsInt(i));
               } else if ("logical".equals(typeName)) {
                  sheet.getRange(row, colIdx).setValue(colVec.getElementAsLogical(i).toBooleanStrict());
               } else {
                  sheet.getRange(row, colIdx).setValue(colVec.getElementAsString(i));
               }
            }
         }
         colIdx++;
      }
   }
}
