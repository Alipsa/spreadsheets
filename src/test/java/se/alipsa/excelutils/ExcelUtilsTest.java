package se.alipsa.excelutils;

import org.junit.jupiter.api.Test;
import org.renjin.sexp.*;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;

public class ExcelUtilsTest {

   @Test
   public void testFindRowNum() throws Exception {
      ExcelReader util = new ExcelReader().setExcel("df.xlsx");
      int rowNum = util.findRowNum(0,0,"Iris");
      util.close();
      assertEquals(35, rowNum);
   }

   @Test
   public void testFindRowNumRainy() throws Exception {
      ExcelReader util = new ExcelReader().setExcel("df.xlsx");
      int rowNum = util.findRowNum(0,0,"NOthing that exist");
      util.close();
      assertEquals(-1, rowNum);
   }
}
