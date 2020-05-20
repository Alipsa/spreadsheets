package se.alipsa.excelutils;

public class Barfer {

   public void barf(String msg) throws Exception {
     throw new Exception(msg);
   }
}