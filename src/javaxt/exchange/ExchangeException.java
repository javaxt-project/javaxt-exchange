package javaxt.exchange;

//******************************************************************************
//**  Exception Class
//******************************************************************************
/**
 *   Enter class description here
 *
 ******************************************************************************/

public class ExchangeException extends Exception{


  //**************************************************************************
  //** Constructor
  //**************************************************************************
  /** Creates a new instance of Exception. */

    public ExchangeException(String error) {
        super(error);
    }

    public ExchangeException() {
        super();
    }


    
    protected static String parseError(org.w3c.dom.Document xml){
        System.out.println("Parse Error!");
        return null;
    }

}