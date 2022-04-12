package se.alipsa.excelutils;

import org.junit.jupiter.api.Test;
import org.renjin.script.RenjinScriptEngine;
import org.renjin.script.RenjinScriptEngineFactory;
import org.renjin.sexp.AttributeMap;
import org.renjin.sexp.SEXP;

import javax.script.ScriptException;

public class SyntaxTest {

  @Test
  public void testSyntax() throws ScriptException {
    RenjinScriptEngineFactory factory = new RenjinScriptEngineFactory();
    RenjinScriptEngine engine = factory.getScriptEngine();

    SEXP list = (SEXP)engine.eval("list('Sheet1'=c(1, 1212, 1, 17), 'Sheet2'=c(5,123,2,10))");
    System.out.println(list.getClass());
    AttributeMap attributes = list.getAttributes();
    attributes.toMap().forEach((symbol, sexp) -> System.out.println(symbol + " = " + sexp));
  }
}
