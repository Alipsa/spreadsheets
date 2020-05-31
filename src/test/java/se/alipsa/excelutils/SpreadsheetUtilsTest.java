package se.alipsa.excelutils;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertEquals;

public class SpreadsheetUtilsTest {

   @Test
   public void testColumnConversion() {
      assertEquals(14, SpreadsheetUtil.asColumnNumber("N"));
      assertEquals(32, SpreadsheetUtil.asColumnNumber("AF"));
      assertEquals(704, SpreadsheetUtil.asColumnNumber("AAB"));

      assertEquals("N", SpreadsheetUtil.asColumnName(14));
      assertEquals("AF", SpreadsheetUtil.asColumnName(32));
      assertEquals("AAB", SpreadsheetUtil.asColumnName(704));
   }
}
