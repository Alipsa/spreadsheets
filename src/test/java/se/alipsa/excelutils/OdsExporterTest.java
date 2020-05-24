package se.alipsa.excelutils;

import com.github.miachm.sods.Sheet;
import com.github.miachm.sods.SpreadSheet;
import org.junit.jupiter.api.Test;
import org.renjin.script.RenjinScriptEngine;
import org.renjin.script.RenjinScriptEngineFactory;
import org.renjin.sexp.ListVector;

import javax.script.ScriptException;
import java.io.File;
import java.io.IOException;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

public class OdsExporterTest {

   @Test
   public void testExport() throws ScriptException, IOException {
      RenjinScriptEngine engine = new RenjinScriptEngineFactory().getScriptEngine();
      ListVector mtcars = (ListVector)engine.eval("mtcars");
      assertEquals(11, mtcars.length());

      File file = new File("mtcars.ods");
      if (file.exists()) file.delete();

      System.out.println("Saving to " + file.getAbsolutePath());

      OdsExporter.exportOds(mtcars, file.getName());
      assertTrue(file.exists());
      OdsExporter.exportOds(mtcars, file.getName(), "1");

      SpreadSheet spreadSheet = new SpreadSheet(file);
      assertEquals(1, spreadSheet.getNumSheets(), "Number of sheets");
      Sheet sheet = spreadSheet.getSheet(0);
      OdsValueExtractor ext = new OdsValueExtractor(sheet);
      int lastRow = 32;
      assertEquals(21.4, ext.getDouble(lastRow, 0), 0.00001);
      assertEquals(4, ext.getInt(lastRow,1));
      assertEquals(2.78, ext.getDouble(lastRow, 5), 0.00001);
      assertEquals(2, ext.getInt(lastRow,10));

      ListVector iris = (ListVector)engine.eval("iris");
      OdsExporter.exportOds(iris, file.getName(), "iris");

      spreadSheet = new SpreadSheet(file);

      assertEquals(2, spreadSheet.getNumSheets(), "Number of sheets");
      assertEquals("1", spreadSheet.getSheet(0).getName(), "mtcars sheet index");
      assertEquals("iris", spreadSheet.getSheet(1).getName(), "iris sheet index");
      Sheet irisSheet = spreadSheet.getSheet("iris");
      ext = new OdsValueExtractor(irisSheet);
      lastRow = 150;
      assertEquals(5.9, ext.getDouble(lastRow,0), 0.00001);
      assertEquals(3, ext.getInt(lastRow, 1));
      assertEquals(5.1, ext.getDouble(lastRow, 2), 0.00001);
      assertEquals(1.8, ext.getDouble(lastRow,3), 0.00001);
      assertEquals("virginica", ext.getString(lastRow,4));

      try {
         assertEquals(52, OdsReader.findRowNum(file.getName(), "iris", SpreadsheetUtil.toColumnNumber("E"), "versicolor"));
      } catch (Exception e) {
         e.printStackTrace();
      }

   }
}
