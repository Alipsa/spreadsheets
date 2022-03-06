package se.alipsa.excelutils;

import java.text.NumberFormat;
import java.text.ParseException;

/**
 * A ValueExtractor is a helper class that makes it easier to get values from a spreadsheet.
 */
public abstract class ValueExtractor {

   protected NumberFormat percentFormat = NumberFormat.getPercentInstance();

   public Double getDouble(Object val) {
      if (val == null) {
         //return 0;
         return null;
      }
      if (val instanceof Double) {
         return (Double) val;
      }
      String strVal = val.toString();
      try {
         return Double.parseDouble(strVal);
      } catch (NumberFormatException e) {
         try {
            percentFormat.parse(strVal).doubleValue();
         } catch (ParseException ignored) {
            // do nothing
         }
         //return 0;
         return null;
      }
   }

   public Double getPercentage(Object val) {
      if (val == null) {
         return null;
      }
      if (val instanceof Double) {
         return (Double) val;
      }
      String strVal = val.toString();
      if (strVal.contains("%")) {
         try {
            return percentFormat.parse(strVal).doubleValue();
         } catch (ParseException e) {
            return null;
         }
      } else {
         return Double.parseDouble(strVal);
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
