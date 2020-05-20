package se.alipsa.excelutils;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;

public class ExcelUtil {

   public static String stringCellVal(Cell cell, FormulaEvaluator evaluator, DataFormatter formatter) {
      if (cell == null) return null;
      switch (cell.getCellType()) {
         case BLANK:
            return "";
         case BOOLEAN:
            return String.valueOf(cell.getBooleanCellValue());
         case ERROR:
            return String.valueOf(cell.getErrorCellValue());
         case FORMULA:
            return readFormattedCellValue(cell, evaluator, formatter);
         case NUMERIC:
            return String.valueOf(cell.getNumericCellValue());
         case STRING:
            return cell.getStringCellValue();
         default:
            return "Unknown type!";
      }
   }

   public static String readFormattedCellValue(Cell cell, FormulaEvaluator evaluator, DataFormatter formatter) {
      try {
         return formatter.formatCellValue(cell, evaluator);
      } catch (RuntimeException e) {
         return e.getMessage(); // Error from evaluator, for example "Don't know how to evaluate name 'xxx'" if we have =xxx() in cell
      }
   }
}
