package se.alipsa.excelutils;

import org.junit.jupiter.api.Test;
import org.renjin.sexp.*;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;

public class OdsImporterTest {

   @Test
   public void testExcelImportWithHeaderRow() throws Exception {
      ListVector vec = OdsImporter.importOds("df.ods", 1,2, 34, 1, 11, true);
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
      ListVector vec = OdsImporter.importOds("df.ods", 1,3, 34, 1, 11, false);
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
      ListVector vec = OdsImporter.importOds("df.ods", 1,3, 34, 1, 11, headerList);
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

   @Test
   public void testImportComplexOds() throws Exception {
      ListVector vec = OdsImporter.importOds(
         "complex.ods",
         1,
         1,
         7,
         "A",
         "F",
         true
      );
      int row = 2;

      LocalDate theDate = LocalDate.from(SpreadsheetUtil.dateTimeFormatter.parse(vec.getElementAsVector("date").getElementAsString(row)));
      assertEquals("2020-05-03", DateTimeFormatter.ofPattern("yyyy-MM-dd").format(theDate));

      LocalDateTime localDateTime = LocalDateTime.parse(vec.getElementAsVector("datetime").getElementAsString(row),
         SpreadsheetUtil.dateTimeFormatter);
      assertEquals("2020-05-03 15:43:12", DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss").format(localDateTime));
      assertEquals(new BigDecimal("102").intValue(), new BigDecimal(vec.getElementAsVector("integer").getElementAsString(row)).intValue());
      assertEquals("5.222", vec.getElementAsVector("decimal").getElementAsString(row));
      assertEquals("three", vec.getElementAsVector("string").getElementAsString(row));
      assertEquals("96.778", vec.getElementAsVector("Numdiff").getElementAsString(row));
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
