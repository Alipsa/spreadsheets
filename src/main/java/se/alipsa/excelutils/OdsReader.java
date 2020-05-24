package se.alipsa.excelutils;

import com.github.miachm.sods.Range;
import com.github.miachm.sods.Sheet;
import com.github.miachm.sods.SpreadSheet;

import java.io.File;

import static se.alipsa.excelutils.FileUtil.checkFilePath;

public class OdsReader {

   static SpreadSheet spreadSheet;

   private OdsReader() {
      // Prevent instantiation
   }

   public static void setOds(String filePath) throws Exception {
      File odsFile = checkFilePath(filePath);
      spreadSheet = new SpreadSheet(odsFile);
   }

   /**
    *
    * @param filePath the excel file
    * @param sheetNumber the sheet index (1 indexed)
    * @param colNumber the column number (1 indexed)
    * @param content the string to search for
    * @return the Row as seen in Excel (1 is first row)
    * @throws Exception if something goes wrong
    */
   public static int findRowNum(String filePath, int sheetNumber, int colNumber, String content) throws Exception {
      setOds(filePath);
      Sheet sheet = spreadSheet.getSheet(sheetNumber -1);
      return findRowNum(sheet, colNumber, content);
   }

   public static int findRowNum(String filePath, int sheetNumber, String colName, String content) throws Exception {
      return findRowNum(filePath, sheetNumber, SpreadsheetUtil.toColumnNumber(colName), content);
   }

   public static int findRowNum(String filePath, String sheetName, String colName, String content) throws Exception {
      return findRowNum(filePath, sheetName, SpreadsheetUtil.toColumnNumber(colName), content);
   }

   /**
    *
    * @param filePath the excel file
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

   /**
    *
    * @param sheet the sheet to search in
    * @param colNumber the column number (1 indexed)
    * @param content the string to search for
    * @return the Row as seen in Excel (1 is first row)
    */
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
    *
    * @param filePath the excel file
    * @param sheetNumber the sheet index (1 indexed)
    * @param rowNumber the row number (1 indexed)
    * @param content the string to search for
    * @return the row number that matched or -1 if not found
    * @throws Exception if something goes wrong
    */
   public static int findColNum(String filePath, int sheetNumber, int rowNumber, String content) throws Exception {
      setOds(filePath);
      Sheet sheet = spreadSheet.getSheet(sheetNumber - 1);
      return findColNum(sheet, rowNumber, content);
   }

   /** return the column as seen in excel (e.g. using column(), 1 is the first column etc
    * @param filePath the excel file
    * @param sheetName the name of the sheet
    * @param rowNumber the row number (1 indexed)
    * @param content the string to search for
    */
   public static int findColNum(String filePath, String sheetName, int rowNumber, String content) throws Exception {
      setOds(filePath);
      Sheet sheet = spreadSheet.getSheet(sheetName);
      return findColNum(sheet, rowNumber, content);
   }

   /**
    * @param sheet the Sheet to search
    * @param rowNumber the row number (1 indexed)
    * @param content the string to search for
    * @return return the column as seen in excel (e.g. using column(), 1 is the first column etc
    */
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
}
