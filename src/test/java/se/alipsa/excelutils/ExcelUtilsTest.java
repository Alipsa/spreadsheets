package se.alipsa.excelutils;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertEquals;

public class ExcelUtilsTest {

   @Test
   public void testFindRowNum() throws Exception {
      int rowNum = ExcelReader.findRowNum("df.xlsx",0,0,"Iris");
      assertEquals(35, rowNum);
   }

   @Test
   public void testFindRowNumRainy() throws Exception {
      int rowNum = ExcelReader.findRowNum("df.xlsx", 0,0,"NOthing that exist");
      assertEquals(-1, rowNum);
   }

   @Test
   public void testFindColNum() throws Exception {
      int colNum = ExcelReader.findColNum("df.xlsx",0,36,"Petal.Length");
      assertEquals(2, colNum);
   }

   @Test
   public void testFindColNumRainy() throws Exception {
      int colNum = ExcelReader.findColNum("df.xlsx",0,16,"Foff");
      assertEquals(-1, colNum);
   }
}
