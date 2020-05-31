package se.alipsa.excelutils;

import org.apache.poi.ss.usermodel.DateUtil;

import java.time.format.DateTimeFormatter;

public class SpreadsheetUtil {

   public static final DateTimeFormatter dateTimeFormatter = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss.SSS");

   public static int asColumnNumber(String name) {
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

   public static int toPoiColumnNumber(String name) {
      return asColumnNumber(name) -1;
   }

   public static String asColumnName(int number) {
      StringBuilder sb = new StringBuilder();
      while (number-- > 0) {
         sb.append((char)('A' + (number % 26)));
         number /= 26;
      }
      return sb.reverse().toString();
   }

   public static String asDateString(double dateNumber, String... pattern) {
      DateTimeFormatter formatter = pattern.length > 0 ? DateTimeFormatter.ofPattern(pattern[0]) : dateTimeFormatter;
      return formatter.format(DateUtil.getLocalDateTime(dateNumber));
   }
}
