package se.alipsa.excelutils;

import org.junit.jupiter.api.Test;
import org.renjin.sexp.*;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;

public class ExcelImporterTest {

   @Test
   public void testExcelImportWithHeaderRow() throws Exception {
      ListVector vec = ExcelImporter.importExcel("df.xlsx", 0,1, 33, 0, 11, true);
      //System.out.println(vec);

      List<String> columnList = toHeaderList(vec);
      // should be 32 rows and 11 variables
      assertEquals(11, columnList.size(), "Number of columns");

      for(int i = 0; i < vec.length(); i++) {
         assertEquals(32, vec.getElementAsVector(vec.getName(i)).length(), "Number of rows");
      }
      assertEquals("mpg", columnList.get(0), "First column name");
      assertEquals("carb", columnList.get(10), "11:th column name");

      assertEquals("2.76", vec.getElementAsVector("drat").getElementAsString(5));
   }

   @Test
   public void testExcelImportNoHeaderRow() throws Exception {
      ListVector vec = ExcelImporter.importExcel("df.xlsx", 0,2, 33, 0, 11, false);
      //System.out.println(vec);

      List<String> columnList = toHeaderList(vec);
      // should be 32 rows and 11 variables
      assertEquals(11, columnList.size(), "Number of columns");

      for(int i = 0; i < vec.length(); i++) {
         assertEquals(32, vec.getElementAsVector(vec.getName(i)).length(), "Number of rows");
      }
      String qsec = "7";
      assertEquals(19.47, vec.getElementAsVector(qsec).getElementAsDouble(17), 0.00001);
   }

   @Test
   public void testExcelImportWithHeaderList() throws Exception {
      Vector headerList = StringVector.newBuilder().addAll(Arrays.asList("1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11")).build();
      ListVector vec = ExcelImporter.importExcel("df.xlsx", 0,2, 33, 0, 11, headerList);
      //System.out.println(vec);

      List<String> columnList = toHeaderList(vec);
      // should be 32 rows and 11 variables
      assertEquals(11, columnList.size(), "Number of columns");

      for(int i = 0; i < vec.length(); i++) {
         assertEquals(32, vec.getElementAsVector(vec.getName(i)).length(), "Number of rows");
      }
      assertEquals("1", columnList.get(0), "First column name");
      assertEquals("11", columnList.get(10), "11:th column name");
      assertEquals(109, vec.getElementAsVector("4").getElementAsInt(31));
   }

   public static List<String> toHeaderList(ListVector df) {
      List<String> colList = new ArrayList<>();
      if (df.hasAttributes()) {
         AttributeMap attributes = df.getAttributes();
         Map<Symbol, SEXP> attrMap = attributes.toMap();
         Symbol s = attrMap.keySet().stream().filter(p -> "names".equals(p.getPrintName())).findAny().orElse(null);
         Vector colNames = (Vector) attrMap.get(s);
         if (colNames != null) {
            for (int i = 0; i < colNames.length(); i++) {
               colList.add(colNames.getElementAsString(i));
            }
         }
      }
      return colList;
   }
}
