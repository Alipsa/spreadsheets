package se.alipsa.excelutils;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertEquals;

public class SpreadsheetUtilsTest {

   @Test
   public void testColumnConversion() {
      assertEquals(14, SpreadsheetUtil.toColumnNumber("N"));
      assertEquals(32, SpreadsheetUtil.toColumnNumber("AF"));
      assertEquals(704, SpreadsheetUtil.toColumnNumber("AAB"));

      assertEquals("N", SpreadsheetUtil.toColumnName(14));
      assertEquals("AF", SpreadsheetUtil.toColumnName(32));
      assertEquals("AAB", SpreadsheetUtil.toColumnName(704));
   }
}
