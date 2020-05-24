package se.alipsa.excelutils;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertEquals;

public class OdsReaderTest {

   @Test
   public void testFindRowNum() throws Exception {
      int rowNum = OdsReader.findRowNum("df.ods",1,1,"Iris");
      assertEquals(36, rowNum);
   }

   @Test
   public void testFindRowNumRainy() throws Exception {
      int rowNum = OdsReader.findRowNum("df.ods", 1,1,"NOthing that exist");
      assertEquals(-1, rowNum);
   }

   @Test
   public void testFindColNum() throws Exception {
      int colNum = OdsReader.findColNum("df.ods",1,37,"Petal.Length");
      assertEquals(3, colNum);
      assertEquals(SpreadsheetUtil.toColumnNumber("C"), colNum);

      colNum = OdsReader.findColNum("df.ods",1,36,"test");
      assertEquals(SpreadsheetUtil.toColumnNumber("L"), colNum);
   }

   @Test
   public void testFindColNumRainy() throws Exception {
      int colNum = OdsReader.findColNum("df.ods",1,17,"Foff");
      assertEquals(-1, colNum);
   }
}
