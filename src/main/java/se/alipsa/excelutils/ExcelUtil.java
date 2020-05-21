package se.alipsa.excelutils;

public class ExcelUtil {

   public static int toColumnNumber(String name) {
      if (name == null) {
         return 0;
      }
      String colName = name.toUpperCase();
      int number = 0;
      for (int i = 0; i < colName.length(); i++) {
         number = number * 26 + (colName.charAt(i) - ('A' - 1));
      }
      return number;
   }

   public static String toColumnName(int number) {
      StringBuilder sb = new StringBuilder();
      while (number-- > 0) {
         sb.append((char)('A' + (number % 26)));
         number /= 26;
      }
      return sb.reverse().toString();
   }
}
