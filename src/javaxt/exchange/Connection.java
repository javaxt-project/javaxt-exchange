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


  //**************************************************************************
  //** getHost
  //**************************************************************************
  /** Returns the URL to the Exchange Web Services (EWS) endpoint.
   */
    public String getHost(){
        return ews;
    }


  //**************************************************************************
  //** getUserName
  //**************************************************************************
  /** Returns the username used to bind the Exchange Web Services.
   */
    public String getUserName(){
        return username;
    }


  //**************************************************************************
  //** createRequest
  //**************************************************************************
  /** Returns an HTTP request object configured to execute a web service
   *  requests. To actually execute a request, users must send a SOAP message
   *  using one of the Request.write() methods.
   */
    protected javaxt.http.Request createRequest(){
        javaxt.http.Request request = new javaxt.http.Request(ews);
        request.validateSSLCertificates(true);
        request.setCredentials(username, password);
        request.setHeader("Accept", "*/*");
        request.setHeader("Content-Type", "text/xml");
        request.setHeader("Accept-Encoding", "gzip,deflate");
        return request;
    }


  //**************************************************************************
  //** getResponse
  //**************************************************************************
  /** Returns a response from a web service service request. The response can
   *  be either an XML Document or the raw HTTP Response object, depending on
   *  the parseResponse parameter.
   * 
   *  @param request An HTTP Request that has already been "written" to.
   *
   *  @param parseResponse If true, parses the HTTP Response from the server
   *  and returns an XML Document. If false, does not parse the response and
   *  returns the raw HTTP Response object.
   */
    protected Object getResponse(javaxt.http.Request request, boolean parseResponse) throws ExchangeException {
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

        if (!parseResponse) return response;

        String text = response.getText();
        org.w3c.dom.Document xml = javaxt.xml.DOM.createDocument(text);
        if (xml==null){ //Possible illegal characters in XML (e.g. "&#x17")
            xml = updateXML(text);
        }

        String error = parseError(xml);
        if (error!=null) throw new ExchangeException(error);
        //new javaxt.io.File("/temp/exchange-execute-" + new java.util.Date().getTime() + ".xml").write(xml);
        return xml;
    }


  //**************************************************************************
  //** execute
  //**************************************************************************
  /** Used to send a SOAP message to the Exchange Web Services.
   *  @param soap Raw SOAP message
   *  @param parseResponse Boolean indicating whether to parse the response.
   *  If true, will return an org.w3c.dom.Document. If false, returns an
   *  javaxt.http.Response.
   */
    public Object execute(String soap, boolean parseResponse) throws ExchangeException {
        javaxt.http.Request request = createRequest();
        request.write(soap);
        return getResponse(request, parseResponse);
    }


  //**************************************************************************
  //** execute
  //**************************************************************************
  /** Used to send a SOAP message to the Exchange Web Services and parse the
   *  response. Returns an return an org.w3c.dom.Document.
   *  @param soap Raw SOAP message
   */
    public org.w3c.dom.Document execute(String soap) throws ExchangeException {
        return (org.w3c.dom.Document) this.execute(soap, true);
    }


  //**************************************************************************
  //** toString
  //**************************************************************************
  /** Returns the host, username, and password.
   */
    public String toString(){
        return "host:  " + ews + "\r\nusername:  " + username + "\r\npassword:  " + password;
    }


  //**************************************************************************
  //** parseError
  //**************************************************************************
  /** Used to parse any error messages. Returns null if no error messages are
   *  found.
   */
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


  //**************************************************************************
  //** updateXML
  //**************************************************************************
  /** For whatever reason, Exchange Web Services return XML documents with
   *  illegal XML characters (e.g. "&#x17"). This is commonly seen in the
   *  GetItem response messages with HTML (e.g. email messages). To circumvent
   *  this issue, this method tries to wrap the invalid characters in a CDATA
   *  block.
   */
    private org.w3c.dom.Document updateXML(String text){

      //Find error. See if its an illegal character.
        while (true)
        try{
            java.io.InputStream is = new java.io.ByteArrayInputStream(text.getBytes("UTF-8"));
            javax.xml.parsers.DocumentBuilderFactory builderFactory = javax.xml.parsers.DocumentBuilderFactory.newInstance();
            javax.xml.parsers.DocumentBuilder builder = builderFactory.newDocumentBuilder();
            return builder.parse(is);
        }
        catch(Exception e){

          //Parse the error message. Look for errors specific to invalid characters.
          //Note this extremely error prone and has only been tested on Java 1.6 in the US Locale...
            String error = e.getMessage().trim();
            if (error.startsWith("Character reference \"") && error.endsWith("\" is an invalid XML character.")){
                String illegalChar = error.substring(error.indexOf("\"")+1, error.lastIndexOf("\""));
                String a = text.substring(0, text.indexOf(illegalChar));
                String b = text.substring(text.indexOf(illegalChar));

                a = a.substring(0, a.lastIndexOf(">")+1) + "<![CDATA[" + a.substring(a.lastIndexOf(">")+1);
                b = b.substring(0, b.indexOf("<")) + "]]>" + b.substring(b.indexOf("<")) ;
                text = a + b;
            }
            else{
                return null;
            }
        }
    }
}