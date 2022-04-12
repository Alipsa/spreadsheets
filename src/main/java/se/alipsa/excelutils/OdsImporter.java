package se.alipsa.excelutils;

import com.github.miachm.sods.Sheet;
import com.github.miachm.sods.SpreadSheet;
import org.renjin.primitives.sequence.IntSequence;
import org.renjin.primitives.vector.ConvertingStringVector;
import org.renjin.sexp.*;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

import static se.alipsa.excelutils.FileUtil.checkFilePath;

/**
 * Import Calc (ods file) into Renjin R in the form of a data.frame (ListVector).
 */
public class OdsImporter {

   private OdsImporter() {
      // prevent instantiation
   }

   /**
    *
    * @param filePath the full path or relative path to the Calc file
    * @param sheetName the name of the sheet to import
    * @param startRowNum the starting row number (1 indexed)
    * @param endRowNum the ending row number (1 indexed)
    * @param startColName the starting column name (e.g. A)
    * @param endColNum the ending column number (1 indexed)
    * @param firstRowAsColNames whether the first row should be used as column names for the dataframe or not.
    * @return a data.frame of string characters with the values in the range specified
    * @throws Exception of the file cannot be read or som other issue occurs.
    */
   public static ListVector importOds(String filePath, String sheetName, int startRowNum, int endRowNum, String startColName, int endColNum, boolean firstRowAsColNames) throws Exception {
      return importOds(filePath, sheetName, startRowNum, endRowNum, SpreadsheetUtil.asColumnNumber(startColName), endColNum, firstRowAsColNames);
   }

   /**
    *
    * @param filePath the full path or relative path to the Calc file
    * @param sheetName the name of the sheet to import
    * @param startRowNum the starting row number (1 indexed)
    * @param endRowNum the ending row number (1 indexed)
    * @param startColNum the start column number (1 indexed)
    * @param endColName the name of the ending column (e.g. Z)
    * @param firstRowAsColNames whether the first row should be used as column names for the dataframe or not.
    * @return a data.frame of string characters with the values in the range specified
    * @throws Exception of the file cannot be read or som other issue occurs.
    */
   public static ListVector importOds(String filePath, String sheetName, int startRowNum, int endRowNum, int startColNum, String endColName, boolean firstRowAsColNames) throws Exception {
      return importOds(filePath, sheetName, startRowNum, endRowNum, startColNum, SpreadsheetUtil.asColumnNumber(endColName), firstRowAsColNames);
   }

   /**
    *
    * @param filePath the full path or relative path to the Calc file
    * @param sheetName the name of the sheet to import
    * @param startRowNum the starting row number (1 indexed)
    * @param endRowNum the ending row number (1 indexed)
    * @param startColName the starting column name (e.g. A)
    * @param endColName the name of the ending column (e.g. Z)
    * @param firstRowAsColNames whether the first row should be used as column names for the dataframe or not.
    * @return a data.frame of string characters with the values in the range specified
    * @throws Exception of the file cannot be read or som other issue occurs.
    */
   public static ListVector importOds(String filePath, String sheetName, int startRowNum, int endRowNum, String startColName, String endColName, boolean firstRowAsColNames) throws Exception {
      return importOds(filePath, sheetName, startRowNum, endRowNum, SpreadsheetUtil.asColumnNumber(startColName), SpreadsheetUtil.asColumnNumber(endColName), firstRowAsColNames);
   }

   /**
    * @param filePath the full path or relative path to the Calc file
    * @param sheetNum the number of the sheet (1 indexed) to read
    * @param startRowNum the starting row number (1 indexed)
    * @param endRowNum the ending row number (1 indexed)
    * @param startColName the starting column name (e.g. A)
    * @param endColName the name of the ending column (e.g. Z)
    * @param firstRowAsColNames whether the first row should be used as column names for the dataframe or not.
    * @return a data.frame of string characters with the values in the range specified
    * @throws Exception of the file cannot be read or som other issue occurs.
    */
   public static ListVector importOds(String filePath, int sheetNum, int startRowNum, int endRowNum, String startColName, String endColName, boolean firstRowAsColNames) throws Exception {
      return importOds(filePath, sheetNum, startRowNum, endRowNum, SpreadsheetUtil.asColumnNumber(startColName), SpreadsheetUtil.asColumnNumber(endColName), firstRowAsColNames);
   }

   /**
    *
    * @param filePath the full path or relative path to the Calc file
    * @param sheetName the name of the sheet to import
    * @param startRowNum the starting row number (1 indexed)
    * @param endRowNum the ending row number (1 indexed)
    * @param startColNum the start column number (1 indexed)
    * @param endColNum the ending column number (1 indexed)
    * @param firstRowAsColNames whether the first row should be used as column names for the dataframe or not.
    * @return a data.frame of string characters with the values in the range specified
    * @throws Exception of the file cannot be read or som other issue occurs.
    */
   public static ListVector importOds(String filePath, String sheetName, int startRowNum, int endRowNum, int startColNum, int endColNum, boolean firstRowAsColNames) throws Exception {
      File excelFile = checkFilePath(filePath);
      List<String> header = new ArrayList<>();
      SpreadSheet spreadSheet = new SpreadSheet(excelFile);
      Sheet sheet = spreadSheet.getSheet(sheetName);
      if (firstRowAsColNames) {
         buildHeaderRow(startRowNum, startColNum, endColNum, header, sheet);
         startRowNum = startRowNum + 1;
      } else {
         for (int i = 1; i <= endColNum - startColNum; i++) {
            header.add(String.valueOf(i));
         }
      }
      return importOds(sheet, startRowNum, endRowNum, startColNum, endColNum, header);

   }

   /**
    *
    * @param filePath the full path or relative path to the Calc file
    * @param sheetNum the number of the sheet (1 indexed) to read
    * @param startRowNum the starting row number (1 indexed)
    * @param endRowNum the ending row number (1 indexed)
    * @param startColNum the start column number (1 indexed)
    * @param endColNum the ending column number (1 indexed)
    * @param firstRowAsColNames whether the first row should be used as column names for the dataframe or not.
    * @return a data.frame of string characters with the values in the range specified
    * @throws Exception of the file cannot be read or som other issue occurs.
    */
   public static ListVector importOds(String filePath, int sheetNum, int startRowNum, int endRowNum, int startColNum, int endColNum, boolean firstRowAsColNames) throws Exception {
      File excelFile = checkFilePath(filePath);
      List<String> header = new ArrayList<>();

      SpreadSheet spreadSheet = new SpreadSheet(excelFile);
      Sheet sheet = spreadSheet.getSheet(sheetNum - 1);
      if (firstRowAsColNames) {
         buildHeaderRow(startRowNum, startColNum, endColNum, header, sheet);
         startRowNum = startRowNum + 1;
      } else {
         for (int i = 0; i <= endColNum - startColNum; i++) {
            header.add(String.valueOf(i+1));
         }
      }
      //System.out.println("Header size is " + header.size() + "; " + header);
      return importOds(sheet, startRowNum, endRowNum, startColNum, endColNum, header);
   }

   /**
    *
    * @param filePath the full path or relative path to the Calc file
    * @param sheetNum the number of the sheet (1 indexed) to read
    * @param startRowNum the starting row number (1 indexed)
    * @param endRowNum the ending row number (1 indexed)
    * @param startColNum the start column number (1 indexed)
    * @param endColNum the ending column number (1 indexed)
    * @param colNames a Vector of column names to use as header
    * @return a data.frame of string characters with the values in the range specified
    * @throws Exception of the file cannot be read or som other issue occurs.
    */
   public static ListVector importOds(String filePath, int sheetNum, int startRowNum, int endRowNum, int startColNum, int endColNum, Vector colNames) throws Exception {
      File odsFile = checkFilePath(filePath);
      List<String> header = new ArrayList<>();
      for (int i = 0; i < colNames.length(); i++) {
         header.add(colNames.getElementAsString(i));
      }
      SpreadSheet spreadSheet = new SpreadSheet(odsFile);
      Sheet sheet = spreadSheet.getSheet(sheetNum - 1);
      return importOds(sheet, startRowNum, endRowNum, startColNum, endColNum, header);

   }

   /**
    *
    * @param filePath the full path or relative path to the Calc file
    * @param sheetNames a vector of sheet names e.g. c('sheet1', 'sheet2')
    * @param importAreas a named list of numeric vectors containing start row, end row, start column, end column e.g.
    *                    list('sheet1' = c(1, 33, 1, 11), 'sheet2' = c(2, 152, 1, 5))
    * @param firstRowAsColumnNames a named vector of logical values for whether the first row should be used as
    *                              column names for the dataframe in the sheet or not.
    *                              E.g. list('sheet1' = TRUE, 'sheet2' = FALSE)
    * @return a named vector of data.frame (ListVectors) corresponding to the imported sheets
    * @throws Exception if the file cannot be read or som other issue occurs.
    */
   public static Vector importOdsSheets(String filePath, Vector sheetNames, ListVector importAreas, Vector firstRowAsColumnNames) throws Exception {
      File odsFile = checkFilePath(filePath);
      ListVector.Builder result = new ListVector.Builder();
      SpreadSheet spreadSheet = new SpreadSheet(odsFile);

      for (int sheetNameIdx = 0; sheetNameIdx < sheetNames.length(); sheetNameIdx++) {
         String sheetName = sheetNames.getElementAsString(sheetNameIdx);
         //System.out.println("Reading sheet " + sheetName);
         Sheet sheet = spreadSheet.getSheet(sheetName);
         List<String> header = new ArrayList<>();

         //System.out.println("Assigning ranges");
         DoubleArrayVector ranges = (DoubleArrayVector) importAreas.getElementAsVector(sheetName);
         int startRowNum = ranges.getElementAsInt(0);
         int endRowNum = ranges.getElementAsInt(1);
         int startColNum = ranges.getElementAsInt(2);
         int endColNum = ranges.getElementAsInt(3);

         //System.out.println("Creating header");
         if (firstRowAsColumnNames.getElementAsLogical(sheetNameIdx).toBooleanStrict()) {
            buildHeaderRow(startRowNum, startColNum, endColNum, header, sheet);
            startRowNum = startRowNum + 1;
         } else {
            for (int i = 0; i <= endColNum - startColNum; i++) {
               header.add(String.valueOf(i + 1));
            }
         }
         //System.out.println("Header size is " + header.size() + "; " + header);
         result.add(importOds(sheet, startRowNum, endRowNum, startColNum, endColNum, header));
      }

      result.setAttribute("names", sheetNames);
      return result.build();
   }

   private static void buildHeaderRow(int startRowNum, int startColNum, int endColNum, List<String> header, Sheet sheet) {
      startRowNum--;
      startColNum--;
      endColNum--;
      OdsValueExtractor ext = new OdsValueExtractor(sheet);
      for (int i = 0; i <= endColNum - startColNum; i++) {
         header.add(ext.getString(startRowNum, startColNum + i));
      }
   }

   private static ListVector importOds(Sheet sheet, int startRowNum, int endRowNum, int startColNum, int endColNum, List<String> colNames) {
      startRowNum--;
      endRowNum--;
      startColNum--;
      endColNum--;

      OdsValueExtractor ext = new OdsValueExtractor(sheet);
      List<StringVector.Builder> builders = stringBuilders(startColNum, endColNum);
      int numRows = 0;
      for (int rowIdx = startRowNum; rowIdx <= endRowNum; rowIdx++) {
         numRows++;
         int i = 0;
         for (int colIdx = startColNum; colIdx <= endColNum; colIdx++) {
            //System.out.println("Adding ext.getString(" + rowIdx + ", " + colIdx+ ") = " + ext.getString(row, colIdx));
            String val = ext.getString(rowIdx, colIdx);
            if (val == null) {
               builders.get(i++).addNA();
            } else {
               if (val.endsWith("%")) {
                  // In excel there is no % in the end of percentage cell, so we make it the same
                  // This has the unfortunate consequence that intentional columns ending with % will be changed
                  try {
                     double dblVal = Double.parseDouble(val.replace("%", "").replace(",", ".")) / 100;
                     val = String.valueOf(dblVal);
                  } catch (NumberFormatException ignored) {
                     // it is not a percentage number, leave the value as it was
                  }
               }
               builders.get(i++).add(val);
            }
         }
      }
      ListVector columnVector = columnInfo(colNames);
      /* call build() on each column and add them as named cols to df */
      ListVector.NamedBuilder dfBuilder = new ListVector.NamedBuilder();
      for (int i = 0; i < columnVector.length(); i++) {
         ListVector ci = (ListVector) columnVector.get(i);
         dfBuilder.add(ci.get("name").asString(), builders.get(i).build());
      }
      dfBuilder.setAttribute("row.names", new ConvertingStringVector(IntSequence.fromTo(1, numRows)));
      dfBuilder.setAttribute("class", StringVector.valueOf("data.frame"));
      return dfBuilder.build();

   }

   private static List<StringVector.Builder> stringBuilders(int startColNum, int endColNum) {
      List<StringVector.Builder> builder = new ArrayList<>();
      for (int i = 0; i <= endColNum - startColNum; i++) {
         builder.add(new StringVector.Builder());
      }
      //System.out.println("created " + builder.size() + " stringBuilders");
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
