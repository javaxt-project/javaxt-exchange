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
   *  @param username Username. Typically this is an email address.
   *  @param password
   */
    public Connection(String host, String username, String password) {
        this.ews = host;
        this.username = username;
        this.password = password;
    }

    public String getHost(){
        return ews;
    }
    
    public String getUserName(){
        return username;
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
        if (status<100) throw new ExchangeException("Failed to connect to Exchange Web Service.");
        if (status>=400){

            String errorMessage = response.getText();
            org.w3c.dom.Document xml = javaxt.xml.DOM.createDocument(errorMessage);
            if (xml!=null){
                org.w3c.dom.Node[] responseMessages = javaxt.xml.DOM.getElementsByTagName("faultstring", xml);
                if (responseMessages.length>0){
                    errorMessage = responseMessages[0].getTextContent();
                }
            }
            if (errorMessage.trim().length()==0) errorMessage = "Failed to connect to Exchange Web Service: " + status + " - " + response.getMessage();
            throw new ExchangeException(errorMessage);
            
        }

        org.w3c.dom.Document xml = response.getXML();
        String error = parseError(xml);
        if (error!=null) throw new ExchangeException(error);
        //new javaxt.io.File("/temp/exchange-execute-" + new java.util.Date().getTime() + ".xml").write(xml);
        return xml;
    }

    
    public String toString(){
        return "host:  " + ews + "\r\nusername:  " + username + "\r\npassword:  " + password;
    }



    private String parseError(org.w3c.dom.Document xml){

        if (xml==null) return ("Invalid Server Response");
        org.w3c.dom.Node[] responseMessages = javaxt.xml.DOM.getElementsByTagName("ResponseMessages", xml);
        if (responseMessages.length>0){
            org.w3c.dom.Node responseMessage = responseMessages[0];
            org.w3c.dom.NodeList childNodes = responseMessage.getChildNodes();
            for (int i=0; i<childNodes.getLength(); i++){
                org.w3c.dom.Node node = childNodes.item(i);
                String nodeName = node.getNodeName();
                if (nodeName.toLowerCase().endsWith("responsemessage")){
                /*
                if (nodeName.contains(":")) nodeName = nodeName.substring(nodeName.indexOf(":")+1);
                if (nodeName.equalsIgnoreCase("FindItemResponseMessage") ||
                    nodeName.equalsIgnoreCase("CreateItemResponseMessage") ||
                    nodeName.equalsIgnoreCase("UpdateItemResponseMessage") ||
                    nodeName.equalsIgnoreCase("DeleteItemResponseMessage"))
                */
                    String responseClass = javaxt.xml.DOM.getAttributeValue(node, "ResponseClass");
                    if (responseClass.equalsIgnoreCase("Error")){

                        String MessageText = "Unknown error.";
                        String ResponseCode = "";
                        String MessageXml = "";

                        org.w3c.dom.NodeList responseDetails = node.getChildNodes();
                        for (int j=0; j<responseDetails.getLength(); j++){
                            org.w3c.dom.Node details = responseDetails.item(j);
                            String attrName = details.getNodeName();
                            String attrValue = javaxt.xml.DOM.getNodeValue(details);

                            if (attrName.contains(":")) attrName = attrName.substring(attrName.indexOf(":")+1);

                            if (attrName.equalsIgnoreCase("MessageText")) MessageText = attrValue;
                            if (attrName.equalsIgnoreCase("ResponseCode")) ResponseCode = attrValue;
                            if (attrName.equalsIgnoreCase("MessageXml")) MessageXml = javaxt.xml.DOM.getText(details.getChildNodes());


                        }


                        return (MessageText + ": " + MessageXml);
                    }

                    break;
                }


            }

        }

        return null;
    }
}