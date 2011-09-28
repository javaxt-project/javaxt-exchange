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
    private String changeKey;
    private Integer count;

    protected Folder(){}


  //**************************************************************************
  //** getFolder
  //**************************************************************************
  /** Retrieves the folder id and item count for a given folder name.
   *  @param folderName Name of the exchange folder (e.g. inbox, contacts, etc).
   */
    public Folder(String folderName, Connection conn) throws ExchangeException {

        folderName = getDistinguishedFolderId(folderName);

        String msg =
        "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
        + "<soap:Envelope xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\">"
        + "<soap:Body>"
        + "<GetFolder xmlns=\"http://schemas.microsoft.com/exchange/services/2006/messages\" xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\">"
        + "<FolderShape><t:BaseShape>Default</t:BaseShape></FolderShape>"
        + "<FolderIds><t:DistinguishedFolderId Id=\"" + folderName + "\"/></FolderIds></GetFolder>"
        + "</soap:Body>"
        + "</soap:Envelope>";



        parseXML(conn.execute(msg));
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

        org.w3c.dom.Node folderID = javaxt.xml.DOM.getElementsByTagName("FolderId", xml)[0]; //xml.getElementsByTagName("t:FolderId").item(0);
        this.id = javaxt.xml.DOM.getAttributeValue(folderID, "Id");
        this.changeKey = javaxt.xml.DOM.getAttributeValue(folderID, "ChangeKey");
        
        try{
            String count = javaxt.xml.DOM.getNodeValue(xml.getElementsByTagName("t:TotalCount").item(0));
            this.count = javaxt.utils.string.cint(count);
        }
        catch(Exception e){
        }
    }

    public String getID(){
        return id;
    }

    public String getChangeKey(){
        return changeKey;
    }

    public Integer getCount(){
        return count;
    }


    
    protected org.w3c.dom.Document getItems(Connection conn, int maxEntries, int offset) throws ExchangeException {

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
        
        return conn.execute(msg);
    }


    protected static org.w3c.dom.Document getItem(String itemID, Connection conn) throws ExchangeException {
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
        
        return conn.execute(msg);
    }


    public static String[] getDistinguishedFolderIds(){
        return DistinguishedFolderIds;
    }

    public static String getDistinguishedFolderId(String folderName){
        for (String folderID : DistinguishedFolderIds){
            if (folderID.equalsIgnoreCase(folderName)) return folderID;
        }
        return null;
    }

    //http://msdn.microsoft.com/en-us/library/exchangewebservices.distinguishedfolderidnametype%28v=exchg.140%29.aspx
    private static String[] DistinguishedFolderIds = new String[]{
        "archivedeleteditems",
        "archivemsgfolderroot",
        "archiverecoverableitemsdeletions",
        "archiverecoverableitemspurges",
        "archiverecoverableitemsroot",
        "archiverecoverableitemsversions",
        "archiveroot",
        "calendar",  //Represents the Calendar folder.
        "contacts",  //Represents the Contacts folder.
        "deleteditems",  //Represents the Deleted Items folder.
        "drafts",  //Represents the Drafts folder.
        "inbox",  //Represents the Inbox folder.
        "journal",  //Represents the Journal folder.
        "junkemail",  //Represents the Junk E-mail folder.
        "msgfolderroot",  //Represents the message folder root.
        "notes",  //Represents the Notes folder.
        "outbox",  //Represents the Outbox folder.
        "publicfoldersroot",
        "recoverableitemsdeletions",
        "recoverableitemspurges",
        "recoverableitemsroot",
        "recoverableitemsversions",
        "root",  //Represents the root of the mailbox.
        "searchfolders",  //Represents the Search Folders folder. This is also an alias for the Finder folder.
        "sentitems",  //Represents the Sent Items folder.
        "tasks",  //Represents the Tasks folder.
        "voicemail"  //Represents the Voice Mail folder.
    };
}