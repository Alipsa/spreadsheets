package se.alipsa.excelutils;

import org.apache.poi.ss.usermodel.*;

public class ValueExtractor {

   private final Sheet sheet;
   private final FormulaEvaluator evaluator;
   private final DataFormatter dataFormatter;

   public ValueExtractor(Sheet sheet, DataFormatter... dataFormatterOpt) {
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
      Object val = getObject(row.getCell(column));
      if (val == null) {
         return 0;
      }
      if (val instanceof Double) {
         return (Double) val;
      }
      try {
         return Double.parseDouble(val.toString());
      } catch (NumberFormatException e) {
         return 0;
      }
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
      Object objVal = getObject(row.getCell(column));
      if (objVal == null) {
         return Integer.MIN_VALUE;
      }
      if (objVal instanceof Double) {
         return (int)(Math.round((Double) objVal));
      }
      if (objVal instanceof Boolean) {
         return (boolean)objVal ? 1 : 0;
      }
      return Integer.parseInt(objVal.toString());
   }

   public String getString(int row, int column) {
      return getString(sheet.getRow(row), column);
   }

   public String getString(Row row, int column) {
      return String.valueOf(getObject(row.getCell(column)));
   }

   public String getString(Cell cell) {
      return String.valueOf(getObject(cell));
   }

   public Long getLong(int row, int column) {
      return getLong(sheet.getRow(row), column);
   }

   public Long getLong(Row row, int column) {
      Object objVal = getObject(row.getCell(column));
      if (objVal == null) {
         return Long.MIN_VALUE;
      }
      if (objVal instanceof Double) {
         return (Math.round((Double) objVal));
      }
      if (objVal instanceof Boolean) {
         return (boolean)objVal ? 1L : 0L;
      }
      return Long.parseLong(objVal.toString());
   }

   public Boolean getBoolean(int row, int column) {
      return getBoolean(sheet.getRow(row), column);
   }

   public Boolean getBoolean(Row row, int column) {
      Object val = getObject(row.getCell(column));
      if (val == null || "".equals(val)) {
         return null;
      }
      return getBoolean(row, column);
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
            return evaluator.evaluate(cell).getNumberValue();
         case BOOLEAN:
            return evaluator.evaluate(cell).getBooleanValue();
         case STRING:
            return evaluator.evaluate(cell).getStringValue();
         default:
            return dataFormatter.formatCellValue(cell);
      }
   }

   public Object getObject(Row row, int colIdx) {
      return getObject(row.getCell(colIdx));
   }
}
