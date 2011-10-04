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
    private Connection conn;
    //private String changeKey;
    //private Integer count;

    protected Folder(){}


  //**************************************************************************
  //** getFolder
  //**************************************************************************
  /** Retrieves the folder id and item count for a given folder name.
   *  @param folderName Name of the exchange folder (e.g. inbox, contacts, etc).
   */
    public Folder(String folderName, Connection conn) throws ExchangeException {

        this.conn = conn;

        String folderID = getDistinguishedFolderId(folderName);
        if (folderID==null) folderID = folderName;

        String msg =
        "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
        + "<soap:Envelope xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\">"
        + "<soap:Body>"
        + "<GetFolder xmlns=\"http://schemas.microsoft.com/exchange/services/2006/messages\" xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\">"
        + "<FolderShape><t:BaseShape>Default</t:BaseShape></FolderShape>"

        
        //This returns the LastModifiedTime for the folder. Unfortunately, this
        //does not reflect changes made to items in this folder
        /*
        + "<FolderShape>"
                + "<t:BaseShape>Default</t:BaseShape>"
                + "<t:AdditionalProperties>"
                + "<t:ExtendedFieldURI PropertyTag=\"0x3008\" PropertyType=\"SystemTime\" />" 
                + "</t:AdditionalProperties>"
        + "</FolderShape>"
        */


        + "<FolderIds><t:DistinguishedFolderId Id=\"" + folderID + "\"/></FolderIds></GetFolder>"
        + "</soap:Body>"
        + "</soap:Envelope>";
        parseXML(conn.execute(msg));
    }


    /*
    public Folder(String xml) {
        this(javaxt.xml.DOM.createDocument(xml));
    }

    public Folder(org.w3c.dom.Document xml){
        parseXML(xml);
    }
    */

    private void parseXML(org.w3c.dom.Document xml) throws ExchangeException {
        org.w3c.dom.Node[] nodes = javaxt.xml.DOM.getElementsByTagName("FolderId", xml);
        if (nodes.length>0){
            org.w3c.dom.Node folderID = nodes[0];
            this.id = javaxt.xml.DOM.getAttributeValue(folderID, "Id");
        }
        if (this.id==null || this.id.length()==0) throw new ExchangeException("Failed to parse FolderId.");
    }

  //**************************************************************************
  //** getID
  //**************************************************************************
  /** Returns the unique Exchange ID for this folder. */

    public String getExchangeID(){
        return id;
    }

    /*
    public String getChangeKey(){
        return changeKey;
    }

    public Integer getCount(){
        return count;
    }
    */


  //**************************************************************************
  //** getItems
  //**************************************************************************
  /** Returns an XML document with shallow representations of items found in
   *  this folder.
   */
    protected org.w3c.dom.Document getItems(int maxEntries, int offset) throws ExchangeException {

        String msg =
        "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
        + "<soap:Envelope xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\" xmlns:m=\"http://schemas.microsoft.com/exchange/services/2006/messages\">"
        + "<soap:Body>"
        + "<m:FindItem Traversal=\"Shallow\">"
        /*
        + "<m:ItemShape><t:BaseShape>Default</t:BaseShape>" //<--"Default" vs "AllProperties"
        + "<t:AdditionalProperties><t:FieldURI FieldURI=\"item:ItemClass\"/></t:AdditionalProperties>"
        + "</m:ItemShape>"
        */
        + "<m:ItemShape>"
                + "<t:BaseShape>Default</t:BaseShape>" //<--"Default" vs "AllProperties"
                + "<t:AdditionalProperties>"
                + "<t:FieldURI FieldURI=\"item:ItemClass\"/>"
                //+ "<t:FieldURI FieldURI=\"item:LastModifiedTime\"/>" //value="item:LastModifiedTime" //<--This doesn't work...
                + "<t:ExtendedFieldURI PropertyTag=\"0x3008\" PropertyType=\"SystemTime\" />" //<--This returns the LastModifiedTime!
                + "</t:AdditionalProperties>"
        + "</m:ItemShape>"

        + "<m:IndexedPageItemView MaxEntriesReturned=\"" + maxEntries + "\" Offset=\"" + offset + "\" BasePoint=\"Beginning\"/><m:ParentFolderIds>"
        + "<t:FolderId Id=\"" + id + "\"/>"
        + "</m:ParentFolderIds>"
        + "</m:FindItem>"
        + "</soap:Body>"
        + "</soap:Envelope>";
        
        return conn.execute(msg);
    }


  //**************************************************************************
  //** getIndex
  //**************************************************************************
  /** Returns a hashmap of all the items found in this folder. The hashmap key
   *  is the item id and the corresponding value is the last modification date.
   */
    public java.util.HashMap<String, javaxt.utils.Date> getIndex() throws ExchangeException {

        String msg =
        "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
        + "<soap:Envelope xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\" xmlns:m=\"http://schemas.microsoft.com/exchange/services/2006/messages\">"
        + "<soap:Body>"
        + "<m:FindItem Traversal=\"Shallow\">"
        + "<m:ItemShape>"
                + "<t:BaseShape>IdOnly</t:BaseShape>" //<--"Default" vs "AllProperties"
                + "<t:AdditionalProperties>"
                + "<t:ExtendedFieldURI PropertyTag=\"0x3008\" PropertyType=\"SystemTime\" />" //<--This returns the LastModifiedTime!
                + "</t:AdditionalProperties>"
        + "</m:ItemShape>"

        + "<m:SortOrder>"
        + "<t:FieldOrder Order=\"Descending\">"
            + "<t:ExtendedFieldURI PropertyTag=\"0x3008\" PropertyType=\"SystemTime\" />"
        + "</t:FieldOrder>"
        + "</m:SortOrder>"
            
        //+ "<m:IndexedPageItemView MaxEntriesReturned=\"1\" Offset=\"0\" BasePoint=\"Beginning\"/>"
        + "<m:ParentFolderIds>"
        + "<t:FolderId Id=\"" + id + "\"/>"
        + "</m:ParentFolderIds>"
        + "</m:FindItem>"
        + "</soap:Body>"
        + "</soap:Envelope>";

        org.w3c.dom.Document xml = conn.execute(msg);
        //new javaxt.io.File("/temp/exchange-sort.xml").write(xml);

        java.util.HashMap<String, javaxt.utils.Date> index = new java.util.HashMap<String, javaxt.utils.Date>();

        org.w3c.dom.Node[] items = javaxt.xml.DOM.getElementsByTagName("Items", xml);
        if (items.length>0){
            org.w3c.dom.NodeList nodes = items[0].getChildNodes();
            for (int i=0; i<nodes.getLength(); i++){
                org.w3c.dom.Node node = nodes.item(i);
                if (node.getNodeType()==1){
                    FolderItem item = new FolderItem(node);
                    index.put(item.getID(), item.getLastModifiedTime());
                }
            }

        }        

        return index;
    }


    protected void findItem(){
    /*
      <m:FindItem Traversal="Shallow"
          xmlns:m=".../messages"
          xmlns:t=".../types">
          <m:ItemShape>
            <t:BaseShape>IdOnly</t:BaseShape>
            <t:AdditionalProperties>
              <t:FieldURI FieldURI="item:Subject" />
              <t:FieldURI FieldURI="calendar:CalendarItemType" />
            </t:AdditionalProperties>
          </m:ItemShape>
          <m:Restriction>
            <t:And>
              <t:IsGreaterThan>
                <t:FieldURI FieldURI="calendar:Start" />
                <t:FieldURIOrConstant>
                  <t:Constant Value="2006-10-16T00:00:00-08:00" />
                </t:FieldURIOrConstant>
              </t:IsGreaterThan>
              <t:IsLessThan>
                <t:FieldURI FieldURI="calendar:End" />
                <t:FieldURIOrConstant>
                  <t:Constant Value="2006-10-20T23:59:59-08:00" />
                </t:FieldURIOrConstant>
              </t:IsLessThan>
            </t:And>
          </m:Restriction>
          <m:ParentFolderIds>
            <t:DistinguishedFolderId Id="calendar"/>
          </m:ParentFolderIds>
        </m:FindItem>
     */
    }


  //**************************************************************************
  //** getDistinguishedFolderIds
  //**************************************************************************
  /** Returns a list of Distinguished Folder IDs.
   */
    public static String[] getDistinguishedFolderIds(){
        return DistinguishedFolderIds;
    }

    
  //**************************************************************************
  //** getDistinguishedFolderId
  //**************************************************************************
  /** Returns the DistinguishedFolderID for a given folder.
   */
    public static String getDistinguishedFolderId(String folderName){
        for (String folderID : DistinguishedFolderIds){
            if (folderID.equalsIgnoreCase(folderName)) return folderID;
        }
        return null;
    }

    
  //**************************************************************************
  //** DistinguishedFolderIds
  //**************************************************************************
  /** Static list of Distinguished Folder IDs. Source:
   *  http://msdn.microsoft.com/en-us/library/exchangewebservices.distinguishedfolderidnametype%28v=exchg.140%29.aspx
   */
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