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
import java.time.LocalDate;
import java.time.LocalDateTime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

public class ExcelExporterTest {

   @Test
   public void testExport() throws Exception {
      RenjinScriptEngine engine = new RenjinScriptEngineFactory().getScriptEngine();
      ListVector mtcars = (ListVector)engine.eval("mtcars");
      assertEquals(11, mtcars.length());

      File file = File.createTempFile("mtcars", ".xlsx");
      if (file.exists()) file.delete();

      ExcelExporter.exportExcel(mtcars, file.getAbsolutePath());
      assertTrue(file.exists());
      ExcelExporter.exportExcel(mtcars, "Sheet0", file.getAbsolutePath());


      try(Workbook workbook = WorkbookFactory.create(file)) {
         Sheet sheet = workbook.getSheetAt(0);
         Row lastRow = sheet.getRow(32);
         assertEquals(21.4, lastRow.getCell(0).getNumericCellValue(), 0.00001);
         assertEquals(4, (int)Math.round(lastRow.getCell(1).getNumericCellValue()));
         assertEquals(2.78, lastRow.getCell(5).getNumericCellValue(), 0.00001);
         assertEquals(2, (int)Math.round(lastRow.getCell(10).getNumericCellValue()));
      }

      ListVector iris = (ListVector)engine.eval("iris");
      ExcelExporter.exportExcel(iris, "iris", file.getAbsolutePath());
      try(Workbook workbook = WorkbookFactory.create(file)) {
         assertEquals(2, workbook.getNumberOfSheets(), "Number of sheets");
         assertEquals(0, workbook.getSheetIndex("Sheet0"), "mtcars sheet index");
         assertEquals(1, workbook.getSheetIndex("iris"), "iris sheet index");
         Sheet irisSheet = workbook.getSheet("iris");
         Row lastRow = irisSheet.getRow(150);
         assertEquals(5.9, lastRow.getCell(0).getNumericCellValue(), 0.00001);
         assertEquals(3, (int)Math.round(lastRow.getCell(1).getNumericCellValue()));
         assertEquals(5.1, lastRow.getCell(2).getNumericCellValue(), 0.00001);
         assertEquals(1.8, lastRow.getCell(3).getNumericCellValue(), 0.00001);
         assertEquals("virginica", lastRow.getCell(4).getStringCellValue());
      }
      assertEquals(52, ExcelReader.findRowNum(file.getAbsolutePath(), "iris", SpreadsheetUtil.toColumnNumber("E"), "versicolor"));
      file.deleteOnExit();
   }
}
