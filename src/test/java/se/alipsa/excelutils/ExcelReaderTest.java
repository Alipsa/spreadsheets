package se.alipsa.excelutils;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertEquals;

public class ExcelReaderTest {

   @Test
   public void testFindRowNum() throws Exception {
      int rowNum = ExcelReader.findRowNum("df.xlsx",1,1,"Iris");
      assertEquals(36, rowNum);
   }

   @Test
   public void testFindRowNumRainy() throws Exception {
      int rowNum = ExcelReader.findRowNum("df.xlsx", 1,1,"NOthing that exist");
      assertEquals(-1, rowNum);
   }

   @Test
   public void testFindColNum() throws Exception {
      int colNum = ExcelReader.findColNum("df.xlsx",1,37,"Petal.Length");
      assertEquals(3, colNum);
      assertEquals(ExcelUtil.toColumnNumber("C"), colNum);

      colNum = ExcelReader.findColNum("df.xlsx",1,36,"test");
      assertEquals(ExcelUtil.toColumnNumber("L"), colNum);
   }

   @Test
   public void testFindColNumRainy() throws Exception {
      int colNum = ExcelReader.findColNum("df.xlsx",1,17,"Foff");
      assertEquals(-1, colNum);
   }
}
