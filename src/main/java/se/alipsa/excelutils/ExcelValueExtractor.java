package se.alipsa.excelutils;

import org.apache.poi.ss.usermodel.*;

public class ExcelValueExtractor extends AbstractValueExtractor {

   private final Sheet sheet;
   private final FormulaEvaluator evaluator;
   private final DataFormatter dataFormatter;

   public ExcelValueExtractor(Sheet sheet, DataFormatter... dataFormatterOpt) {
      this.sheet = sheet;
      evaluator = sheet.getWorkbook().getCreationHelper().createFormulaEvaluator();
      if (dataFormatterOpt.length > 0) {
         dataFormatter = dataFormatterOpt[0];
      } else {
         dataFormatter = new DataFormatter();
      }
   }


   public double getDouble(int row, int column) {
      return getDouble(sheet.getRow(row), column);
   }

   public double getDouble(Row row, int column) {
      return getDouble(getObject(row.getCell(column)));
   }

   public float getFloat(int row, int column) {
      return (float) getDouble(sheet.getRow(row), column);
   }

   public float getFloat(Row row, int column) {
      return (float) getDouble(row, column);
   }

   public int getInt(int row, int column) {
      return getInt(sheet.getRow(row), column);
   }

   public int getInt(Row row, int column) {
      return getInt(getObject(row.getCell(column)));
   }

   public String getString(int row, int column) {
      return getString(sheet.getRow(row), column);
   }

   public String getString(Row row, int column) {
      return String.valueOf(getObject(row.getCell(column)));
   }

   public String getString(Cell cell) {
      return getString(getObject(cell));
   }

   public Long getLong(int row, int column) {
      return getLong(sheet.getRow(row), column);
   }

   public Long getLong(Row row, int column) {
      return(getLong(getObject(row.getCell(column))));
   }

   public Boolean getBoolean(int row, int column) {
      return getBoolean(sheet.getRow(row), column);
   }

   public Boolean getBoolean(Row row, int column) {
      return getBoolean(getObject(row.getCell(column)));
   }

   /** get the value from a Excel cell */
   public Object getObject(Cell cell) {
      if (cell == null) {
         return null;
      }
      switch (cell.getCellType()) {
         case BLANK:
            return null;
         case NUMERIC:
            if (DateUtil.isCellDateFormatted(cell)) {
               return SpreadsheetUtil.dateTimeFormatter.format(cell.getLocalDateTimeCellValue());
            }
            return cell.getNumericCellValue();
         case BOOLEAN:
            return cell.getBooleanCellValue();
         case STRING:
            return cell.getStringCellValue();
         case FORMULA:
            return getValueFromFormulaCell(cell);
         default:
            return dataFormatter.formatCellValue(cell);
      }
   }

   private Object getValueFromFormulaCell(Cell cell) {
      switch (evaluator.evaluateFormulaCell(cell)) {
         case BLANK:
            return null;
         case NUMERIC:
            if (DateUtil.isCellDateFormatted(cell)) {
               return cell.getLocalDateTimeCellValue();
            }
            return evaluator.evaluate(cell).getNumberValue();
         case BOOLEAN:
            return evaluator.evaluate(cell).getBooleanValue();
         case STRING:
            return evaluator.evaluate(cell).getStringValue();
         default:
            return dataFormatter.formatCellValue(cell);
      }
   }
}
