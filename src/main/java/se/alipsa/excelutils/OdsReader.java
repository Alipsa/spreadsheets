package se.alipsa.excelutils;

import com.github.miachm.sods.Sheet;
import com.github.miachm.sods.SpreadSheet;
import org.renjin.sexp.StringArrayVector;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

import static se.alipsa.excelutils.FileUtil.checkFilePath;

public class OdsReader {

   private static SpreadSheet spreadSheet;

   private OdsReader() {
      // Prevent instantiation
   }

   private static void setOds(String filePath) throws Exception {
      File odsFile = checkFilePath(filePath);
      spreadSheet = new SpreadSheet(odsFile);
   }

   /**
    * Find the first row index matching the content.
    * @param filePath the ods file
    * @param sheetNumber the sheet index (1 indexed)
    * @param colNumber the column number (1 indexed)
    * @param content the string to search for
    * @return the Row as seen in Calc (1 is first row)
    * @throws Exception if something goes wrong
    */
   public static int findRowNum(String filePath, int sheetNumber, int colNumber, String content) throws Exception {
      setOds(filePath);
      Sheet sheet = spreadSheet.getSheet(sheetNumber -1);
      return findRowNum(sheet, colNumber, content);
   }

   /**
    * Find the first row index matching the content.
    * @param filePath the ods file
    * @param sheetNumber the sheet index (1 indexed)
    * @param colName the column name (A for first column etc.)
    * @param content the string to search for
    * @return the Row as seen in Calc (1 is first row)
    * @throws Exception if something goes wrong
    */
   public static int findRowNum(String filePath, int sheetNumber, String colName, String content) throws Exception {
      return findRowNum(filePath, sheetNumber, SpreadsheetUtil.asColumnNumber(colName), content);
   }

   /**
    * Find the first row index matching the content.
    * @param filePath the ods file
    * @param sheetName the name of the sheet to search in
    * @param colName the column name (A for first column etc.)
    * @param content the string to search for
    * @return the Row as seen in Calc (1 is first row)
    * @throws Exception if something goes wrong
    */
   public static int findRowNum(String filePath, String sheetName, String colName, String content) throws Exception {
      return findRowNum(filePath, sheetName, SpreadsheetUtil.asColumnNumber(colName), content);
   }

   /**
    *
    * @param filePath the ods file
    * @param sheetName the name of the sheet
    * @param colNumber the column number (1 indexed)
    * @param content the string to search for
    * @return the Row as seen in Excel (1 is first row)
    * @throws Exception if something goes wrong
    */
   public static int findRowNum(String filePath, String sheetName, int colNumber, String content) throws Exception {
      setOds(filePath);
      Sheet sheet = spreadSheet.getSheet(sheetName);
      return findRowNum(sheet, colNumber, content);
   }

   private static int findRowNum(Sheet sheet, int colNumber, String content) {
      OdsValueExtractor ext = new OdsValueExtractor(sheet);
      int poiColNum = colNumber -1;

      for (int rowCount = 0; rowCount < sheet.getDataRange().getLastRow(); rowCount ++) {
         //System.out.println(rowCount + ": " + ext.getString(rowCount, poiColNum));
         if (content.equals(ext.getString(rowCount, poiColNum))) {
            return rowCount + 1;
         }
      }
      return -1;
   }

   /**
    * Find the first column index matching the content criteria
    * @param filePath the ods file
    * @param sheetNumber the sheet index (1 indexed)
    * @param rowNumber the row number (1 indexed)
    * @param content the string to search for
    * @return the row number that matched or -1 if not found
    * @throws Exception if some read problem occurs
    */
   public static int findColNum(String filePath, int sheetNumber, int rowNumber, String content) throws Exception {
      setOds(filePath);
      Sheet sheet = spreadSheet.getSheet(sheetNumber - 1);
      return findColNum(sheet, rowNumber, content);
   }

   /** return the column as seen in the Open Document Spreadsheet (e.g. using column(), 1 is the first column etc
    * @param filePath the ods file
    * @param sheetName the name of the sheet
    * @param rowNumber the row number (1 indexed)
    * @param content the string to search for
    * @return the row number that matched or -1 if not found
    * @throws Exception if the file cannot be read
    */
   public static int findColNum(String filePath, String sheetName, int rowNumber, String content) throws Exception {
      setOds(filePath);
      Sheet sheet = spreadSheet.getSheet(sheetName);
      return findColNum(sheet, rowNumber, content);
   }

   private static int findColNum(Sheet sheet, int rowNumber, String content) {
      if (content==null) return -1;
      OdsValueExtractor ext = new OdsValueExtractor(sheet);
      int poiRowNum = rowNumber - 1;
      for (int colNum = 0; colNum < sheet.getDataRange().getLastColumn(); colNum++) {
         if (content.equals(ext.getString(poiRowNum, colNum))) {
            return colNum + 1;
         }
      }
      return -1;
   }

   public static StringArrayVector getSheetNames(String filePath) throws Exception {
      setOds(filePath);
      List<String> names = new ArrayList<>();
      spreadSheet.getSheets().forEach(s -> names.add(s.getName()));
      return new StringArrayVector(names);
   }
}
