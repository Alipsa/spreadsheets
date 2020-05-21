package se.alipsa.excelutils;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertEquals;

public class ExcelUtilsTest {

   @Test
   public void testColumnConversion() {
      assertEquals(14, ExcelUtil.toColumnNumber("N"));
      assertEquals(32, ExcelUtil.toColumnNumber("AF"));
      assertEquals(704, ExcelUtil.toColumnNumber("AAB"));

      assertEquals("N", ExcelUtil.toColumnName(14));
      assertEquals("AF", ExcelUtil.toColumnName(32));
      assertEquals("AAB", ExcelUtil.toColumnName(704));
   }
}
