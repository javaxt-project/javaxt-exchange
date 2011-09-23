package javaxt.exchange;

//******************************************************************************
//**  Connection Class
//******************************************************************************
/**
 *   Used to encapsulate raw connection information
 *
 ******************************************************************************/

public class Connection {


    private String username;
    private String password;
    private String ews;


  //**************************************************************************
  //** Constructor
  //**************************************************************************
  /** Creates a new instance of Connection. 
   *  @param host URL to the Exchange Web Services (EWS) endpoint. The url
   *  typically ends with "Exchange.asmx".
   *  @param username
   *  @param password
   */
    public Connection(String host, String username, String password) {
        this.ews = host;
        this.username = username;
        this.password = password;
    }






  //**************************************************************************
  //** execute
  //**************************************************************************
  /** Used to send a SOAP message to the Exchange Web Services.
   *  @param soap Raw SOAP message
   *  @return Raw HTTP response
   */
    public org.w3c.dom.Document execute(String soap) throws ExchangeException {

        javaxt.http.Request request = new javaxt.http.Request(ews);
        request.setCredentials(username, password);
        request.setHeader("Accept", "*/*");
        request.setHeader("Content-Type", "text/xml");
        request.write(soap);

        javaxt.http.Response response = request.getResponse();
        int status = response.getStatus();
        if (status>=400){
            throw new ExchangeException(response.getErrorMessage());
        }

        org.w3c.dom.Document xml = response.getXML();
        String error = ExchangeException.parseError(xml);
        if (error!=null) throw new ExchangeException(error);
        
        return xml;
    }

    
    public String toString(){
        return "host:  " + ews + "\r\nusername:  " + username + "\r\npassword:  " + password;
    }

}