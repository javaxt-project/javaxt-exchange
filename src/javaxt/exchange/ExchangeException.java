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
        super(parseError(error));
    }

    public ExchangeException() {
        super();
    }


    
    private static String parseError(String error){
        return error;
    }

}