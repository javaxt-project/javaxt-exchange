package javaxt.exchange;

//******************************************************************************
//**  Folder Class
//******************************************************************************
/**
 *   Used to represent a folder
 *
 ******************************************************************************/

public class Folder {

    private String id;
    private Integer count;

    protected Folder(){}


  //**************************************************************************
  //** getFolder
  //**************************************************************************
  /** Retrieves the folder id and item count for a given folder name.
   *  @param folderName Name of the exchange folder (e.g. inbox, contacts, etc).
   */
    public Folder(String folderName, Connection conn){
        String msg =
        "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
        + "<soap:Envelope xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\">"
        + "<soap:Body>"
        + "<GetFolder xmlns=\"http://schemas.microsoft.com/exchange/services/2006/messages\" xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\">"
        + "<FolderShape><t:BaseShape>Default</t:BaseShape></FolderShape>"
        + "<FolderIds><t:DistinguishedFolderId Id=\"" + folderName + "\"/></FolderIds></GetFolder>"
        + "</soap:Body>"
        + "</soap:Envelope>";



        javaxt.http.Response response = conn.execute(msg);
        System.out.println(response.getStatus() + " - " + response.getMessage());
        parseXML(response.getXML());
    }



  //**************************************************************************
  //** Constructor
  //**************************************************************************
  /** Creates a new instance of Folder. */

    public Folder(String xml) {
        this(javaxt.xml.DOM.createDocument(xml));
    }

    public Folder(org.w3c.dom.Document xml){
        parseXML(xml);
    }


    private void parseXML(org.w3c.dom.Document xml){
        this.id = javaxt.xml.DOM.getAttributeValue(xml.getElementsByTagName("t:FolderId").item(0), "Id");
        String count = javaxt.xml.DOM.getNodeValue(xml.getElementsByTagName("t:TotalCount").item(0));

        try{
            this.count = javaxt.utils.string.cint(count);
        }
        catch(Exception e){
        }
    }

    public String getID(){
        return id;
    }

    public Integer getCount(){
        return count;
    }


    
    protected org.w3c.dom.Document getItems(Connection conn, int maxEntries, int offset){

        String msg =
        "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
        + "<soap:Envelope xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\" xmlns:m=\"http://schemas.microsoft.com/exchange/services/2006/messages\">"
        //+ "<soap:Header><t:RequestServerVersion Version=\"Exchange2007_SP1\"/></soap:Header>"
        + "<soap:Body>"
        + "<m:FindItem Traversal=\"Shallow\">"
        + "<m:ItemShape><t:BaseShape>Default</t:BaseShape>"
        + "<t:AdditionalProperties><t:FieldURI FieldURI=\"item:ItemClass\"/></t:AdditionalProperties>"
        + "</m:ItemShape>"
        + "<m:IndexedPageItemView MaxEntriesReturned=\"" + maxEntries + "\" Offset=\"" + offset + "\" BasePoint=\"Beginning\"/><m:ParentFolderIds>"
        + "<t:FolderId Id=\"" + id + "\"/>"
        + "</m:ParentFolderIds>"
        + "</m:FindItem>"
        + "</soap:Body>"
        + "</soap:Envelope>";

        javaxt.http.Response response = conn.execute(msg);
        return response.getXML();
    }


    protected static org.w3c.dom.Document getItem(String itemID, Connection conn){
        String msg =
        "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
        + "<soap:Envelope xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\" "
            + "xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\" "
            + "xmlns:m=\"http://schemas.microsoft.com/exchange/services/2006/messages\">"
        + "<soap:Body>"
        + "<m:GetItem>"
        + "<m:ItemShape><t:BaseShape>AllProperties</t:BaseShape></m:ItemShape>"
        + "<m:ItemIds><t:ItemId Id=\"" + itemID + "\"/></m:ItemIds>"
        + "</m:GetItem>"
        + "</soap:Body>"
        + "</soap:Envelope>";

        javaxt.http.Response response = conn.execute(msg);
        return response.getXML();
    }
}