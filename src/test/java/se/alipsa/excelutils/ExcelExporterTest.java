package se.alipsa.excelutils;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.jupiter.api.Test;
import org.renjin.script.RenjinScriptEngine;
import org.renjin.script.RenjinScriptEngineFactory;
import org.renjin.sexp.ListVector;

import javax.script.ScriptException;

import java.io.File;
import java.io.IOException;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

public class ExcelExporterTest {

   @Test
   public void testExport() throws ScriptException, IOException {
      RenjinScriptEngine engine = new RenjinScriptEngineFactory().getScriptEngine();
      ListVector mtcars = (ListVector)engine.eval("mtcars");
      assertEquals(11, mtcars.length());

      File file = new File("mtcars.xlsx");
      if (file.exists()) file.delete();

      ExcelExporter.exportExcel(mtcars, file.getName(), false);
      assertTrue(file.exists());
      ExcelExporter.exportExcel(mtcars, file.getName(), true);


      try(Workbook workbook = WorkbookFactory.create(file)) {
         Sheet sheet = workbook.getSheetAt(0);
         Row lastRow = sheet.getRow(32);
         assertEquals(21.4, lastRow.getCell(0).getNumericCellValue(), 0.00001);
         assertEquals(4, (int)Math.round(lastRow.getCell(1).getNumericCellValue()));
         assertEquals(2.78, lastRow.getCell(5).getNumericCellValue(), 0.00001);
         assertEquals(2, (int)Math.round(lastRow.getCell(10).getNumericCellValue()));
      }

      ListVector iris = (ListVector)engine.eval("iris");
      ExcelExporter.exportExcel(iris, file.getName(), "iris", true);


   }
}
