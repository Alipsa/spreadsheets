package se.alipsa.excelutils;

import org.apache.poi.ss.usermodel.*;

/**
 * A value extractor specialized in extracting info from an Excel file
 */
public class ExcelValueExtractor extends ValueExtractor {

   private final Sheet sheet;
   private final FormulaEvaluator evaluator;
   private final DataFormatter dataFormatter;

   public ExcelValueExtractor(Sheet sheet, DataFormatter... dataFormatterOpt) {
      if (sheet == null) {
         throw new IllegalArgumentException("Sheet is null, will not be able to extract any values");
      }
      this.sheet = sheet;
      evaluator = sheet.getWorkbook().getCreationHelper().createFormulaEvaluator();
      if (dataFormatterOpt.length > 0) {
         dataFormatter = dataFormatterOpt[0];
      } else {
         dataFormatter = new DataFormatter();
      }
   }


   public Double getDouble(int row, int column) {
      return getDouble(sheet.getRow(row), column);
   }

   public Double getDouble(Row row, int column) {
      return row == null ? null : getDouble(getObject(row.getCell(column)));
   }

   public Float getFloat(int row, int column) {
      Double d = getDouble(sheet.getRow(row), column);
      return d == null ? null : d.floatValue();
   }

   public Float getFloat(Row row, int column) {
      if (row == null) return null;
      Double d = getDouble(row, column);
      return d == null ? null : d.floatValue();
   }

   public Integer getInteger(int row, int column) {
      return getInteger(sheet.getRow(row), column);
   }

   public Integer getInteger(Row row, int column) {
      return row == null ? null : getInt(getObject(row.getCell(column)));
   }

   public String getString(int row, int column) {
      return getString(sheet.getRow(row), column);
   }

   public String getString(Row row, int column) {
      if (row == null) return null;
      Object val = getObject(row.getCell(column));
      return val == null ? null : String.valueOf(val);
   }

   public String getString(Cell cell) {
      return getString(getObject(cell));
   }

   public Long getLong(int row, int column) {
      return getLong(sheet.getRow(row), column);
   }

   public Long getLong(Row row, int column) {
      return row == null ? null : getLong(getObject(row.getCell(column)));
   }

   public Boolean getBoolean(int row, int column) {
      return getBoolean(sheet.getRow(row), column);
   }

   public Boolean getBoolean(Row row, int column) {
      return row == null ? null : getBoolean(getObject(row.getCell(column)));
   }

   /**
    * get the value from a Excel cell
    * @param cell the cell to extract the value from
    * @return the value
    */
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
