package javaxt.exchange;
import java.io.IOException;

//******************************************************************************
//**  Attachment Class
//******************************************************************************
/**
 *   Used to represent an attachment associated with a FolderItem.
 *
 ******************************************************************************/

public class Attachment {

    private String id;
    private String name;
    private String contentType;


  //**************************************************************************
  //** Constructor
  //**************************************************************************
  /** Creates a new instance of this class using an attachment ID. Initiates
   *  a "GetAttachment" request but stops downloading and parsing the response
   *  as soon as it reaches the "Content" tag.
   */
    public Attachment(String id, Connection conn) throws ExchangeException, java.io.IOException {
        java.io.InputStream inputStream = getAttachment(id, conn);
        parseResponse(inputStream, false);
        inputStream.close();
        this.id = id;
    }


  //**************************************************************************
  //** Constructor
  //**************************************************************************
    protected Attachment(org.w3c.dom.Node node){
        org.w3c.dom.NodeList childNodes = node.getChildNodes();
        for (int j=0; j<childNodes.getLength(); j++){
            org.w3c.dom.Node childNode = childNodes.item(j);
            if (childNode.getNodeType()==1){
                String nodeName = childNode.getNodeName();
                if (nodeName.contains(":")) nodeName = nodeName.substring(nodeName.indexOf(":")+1);

                if (nodeName.equalsIgnoreCase("AttachmentId")){
                    id = javaxt.xml.DOM.getAttributeValue(childNode, "Id");
                }
                else if(nodeName.equalsIgnoreCase("Name")){
                    name = javaxt.xml.DOM.getNodeValue(childNode);
                }
                else if(nodeName.equalsIgnoreCase("ContentType")){
                    contentType = javaxt.xml.DOM.getNodeValue(childNode);
                }
            }
        }
    }
    
    public String getID(){
        return id;
    }

    public String getName(){
        return name;
    }

    public String getContentType(){
        return contentType;
    }

    public String toString(){
        return name;
    }

    public int hashCode(){
        return id.hashCode();
    }

    public boolean equals(Object obj){
        if (obj==null) return false;
        if (obj instanceof Attachment){
            return ((Attachment) obj).id.equals(id);
        }
        return false;
    }


  //**************************************************************************
  //** download
  //**************************************************************************
  /** Used to download the attachment. Returns an input stream with the
   *  attachment.
   */
    public java.io.InputStream download(Connection conn) throws ExchangeException {
        try{
            return parseResponse(getAttachment(id, conn), true);
        }
        catch(java.io.IOException e){
            throw new ExchangeException(e.getLocalizedMessage());
        }
    }


  //**************************************************************************
  //** getAttachment
  //**************************************************************************
  /** Executes the "GetAttachment" request and returns the raw response as an
   *  inputstream.
   */
    private java.io.InputStream getAttachment(String id, Connection conn)
        throws java.io.IOException, ExchangeException {

        String msg =
        "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
        + "<soap:Envelope xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\">"
        + "<soap:Body>"
        + "<GetAttachment xmlns=\"http://schemas.microsoft.com/exchange/services/2006/messages\" xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\">"
        + "<AttachmentShape/>"
        + "<AttachmentIds><t:AttachmentId Id=\"" + id + "\"/></AttachmentIds>"
        + "</GetAttachment>"
        + "</soap:Body>"
        + "</soap:Envelope>";

        javaxt.http.Response response = (javaxt.http.Response) conn.execute(msg, false);
        
        String contentEncoding = response.getHeader("Content-Encoding");
        if (contentEncoding!=null && contentEncoding.equalsIgnoreCase("gzip")){
            return new java.util.zip.GZIPInputStream(response.getInputStream());
        }
        else{
            return response.getInputStream();
        }
    }


  //**************************************************************************
  //** parseResponse
  //**************************************************************************
  /** Used to parse the "GetAttachmentResponse" response message. Instead of
   *  using a DOM parser, this method parses the response XML as a stream.
   *  This is important in order to minimize memory usage when downloading
   *  large attachments.
   */
    private java.io.InputStream parseResponse(java.io.InputStream inputStream, boolean getContent)
        throws ExchangeException, java.io.IOException {

        boolean concat = true;
        StringBuffer tag = new StringBuffer();
        int x=0;
        while ( (x = inputStream.read()) != -1) {

            char c = (char) x;

            if (c == '<'){
                concat = true;
            }


            if (concat==true){
                tag.append(c);
            }


            if (c==('>') && concat==true){
                concat = false;

                if (tag.indexOf("</")==-1){
                    String node = tag.toString().trim();
                    String nodeName = node.substring(1).trim();
                    if (nodeName.endsWith("/>")) nodeName = nodeName.substring(0, nodeName.length()-2).trim();
                    else nodeName = nodeName.substring(0, nodeName.length()-1).trim();
                    if (nodeName.contains(" ")) nodeName = nodeName.substring(0, nodeName.indexOf(" "));
                    if (nodeName.contains(":")) nodeName = nodeName.substring(nodeName.indexOf(":")+1);

                    
                    if (nodeName.equalsIgnoreCase("Name")){
                        name = getNodeValue(inputStream);
                    }
                    else if(nodeName.equalsIgnoreCase("ContentType")){
                        contentType = getNodeValue(inputStream);
                    }
                    
                    if (nodeName.toLowerCase().endsWith("responsemessage")){
                        if (node.toLowerCase().contains("error")){
                            if (!node.endsWith("/>")) node = node.substring(0, node.length()-1) + "/>";
                            org.w3c.dom.Document XMLDoc = javaxt.xml.DOM.createDocument(node);
                            String responseClass = javaxt.xml.DOM.getAttributeValue(XMLDoc.getFirstChild(), "ResponseClass");
                            if (responseClass.equalsIgnoreCase("Error")){
                                inputStream.close();
                                throw new ExchangeException("Failed to download attachment.");
                            }
                        }
                    }
                    else if(nodeName.equalsIgnoreCase("Content")){
                        if (getContent){
                            return new javaxt.utils.Base64.InputStream(
                                new ContentInputStream(inputStream)
                            );
                        }
                        else{
                            return null;
                        }
                    }
                }


                tag = new StringBuffer();
            }
        }
        return null;
    }


  //**************************************************************************
  //** getNodeValue
  //**************************************************************************
  /** Returns the value of a node by parsing an input stream.
   */
    private String getNodeValue(java.io.InputStream inputStream) throws java.io.IOException {
        StringBuffer str = new StringBuffer();
        int x = 0;
        while ( (x = inputStream.read()) != -1) {

            char c = (char) x;

            if (c == '<'){
                break;
            }

            str.append(c);
        }
        return str.toString();
    }


  //**************************************************************************
  //** ContentInputStream
  //**************************************************************************
  /** Provides an input stream for reading the "Content" tag found in a
   *  "GetAttachmentResponse" message.
   */
    public static class ContentInputStream extends java.io.InputStream {

        private java.io.InputStream inputStream;

        protected ContentInputStream(java.io.InputStream inputStream){
            this.inputStream = inputStream;
        }


      /** Returns false. This stream does not support the mark and reset
       *  methods.
       */
        public boolean markSupported(){
            return false;
        }


      /** Reads the next byte of data from the "Content" tag. The byte is
       *  returned as a positive integer. Returns a value of -1 if there is
       *  no more data to read (e.g. found whitespace, start of new xml tag,
       *  or end of stream).
       */
        public int read() throws IOException {

            int x = inputStream.read();
            if (x==-1) return -1;
            else{
                char c = (char) x;
                if (c == '<'){
                    return -1;
                }
                else if (c==' '){
                    while ( (x = inputStream.read()) != -1) {
                        c = (char) x;
                        if (c == '<') return -1;
                        if (c != ' ') return x;
                    }
                    return -1;
                }
                else return x;
            }
        }
    }

}