package se.alipsa.excelutils;

/**
 * A ValueExtractor is a helper class that makes it easier to get values from a spreadsheet.
 */
public abstract class ValueExtractor {

   public Double getDouble(Object val) {
      if (val == null) {
         //return 0;
         return null;
      }
      if (val instanceof Double) {
         return (Double) val;
      }
      try {
         return Double.parseDouble(val.toString());
      } catch (NumberFormatException e) {
         //return 0;
         return null;
      }
   }

   public Integer getInt(Object objVal) {
      if (objVal == null) {
         //return Integer.MIN_VALUE;
         return null;
      }
      if (objVal instanceof Double) {
         return (int)(Math.round((Double) objVal));
      }
      if (objVal instanceof Boolean) {
         return (boolean)objVal ? 1 : 0;
      }
      return Integer.parseInt(objVal.toString());
   }

   public Long getLong(Object objVal) {
      if (objVal == null) {
         //return Long.MIN_VALUE;
         return null;
      }
      if (objVal instanceof Double) {
         return (Math.round((Double) objVal));
      }
      if (objVal instanceof Boolean) {
         return (boolean)objVal ? 1L : 0L;
      }
      return Long.parseLong(objVal.toString());
   }

   public Boolean getBoolean(Object val) {
      if (val == null || "".equals(val)) {
         return null;
      }
      if (val instanceof Boolean) {
         return (Boolean) val;
      } else if (val instanceof Number) {
         int num = (int)Math.round(((Number)val).doubleValue());
         return num == 1;
      } else {
         String strVal = String.valueOf(val).toLowerCase();
         switch (strVal) {
            case "j":
            case "y":
            case "ja":
            case "yes":
            case "1":
            case "true":
            case "on":
               return true;
            default:
               return false;
         }
      }
   }

   public String getString(Object val) {
      return val == null ? null : String.valueOf(val);
   }
}
