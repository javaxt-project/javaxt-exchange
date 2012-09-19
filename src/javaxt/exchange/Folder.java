package javaxt.exchange;

//******************************************************************************
//**  Folder Class
//******************************************************************************
/**
 *   Used to represent a folder (e.g. "Inbox", "Contacts", etc.)
 *
 ******************************************************************************/

public class Folder {

    private String id;
    private String name;
    private String changeKey;
    private Integer totalCount;
    private Integer unreadCount;
    private Connection conn;


  //**************************************************************************
  //** Constructor
  //**************************************************************************
  /** This constructor is provided for application developers who wish to
   *  extend this class.
   */
    protected Folder(){}


  //**************************************************************************
  //** Constructor
  //**************************************************************************
  /** Creates a new instance of this class using a folder name/id.
   *  @param folderName Name of the exchange folder (e.g. inbox, contacts, etc).
   */
    public Folder(String folderName, Connection conn) throws ExchangeException {

        this.conn = conn;

        String folderID = getDistinguishedFolderId(folderName);
        if (folderID==null){
            folderID = "<t:FolderId Id=\"" + folderName + "\"/>";
        }
        else{
            folderID = "<t:DistinguishedFolderId Id=\"" + folderID + "\"/>";
        }

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


        + "<FolderIds>" + folderID + "</FolderIds></GetFolder>"
        + "</soap:Body>"
        + "</soap:Envelope>";
        parseXML(conn.execute(msg));
    }


  //**************************************************************************
  //** Constructor
  //**************************************************************************
  /** Private constructor used by the getFolders() method.
   */
    private Folder(org.w3c.dom.Node folder, Connection conn) throws ExchangeException {
        this.conn = conn;
        parseFolderNode(folder);
    }


  //**************************************************************************
  //** getFolders
  //**************************************************************************
  /** Returns an array of folders found in this folder. Returns a zero length
   *  array of no folders are found.
   */
    public Folder[] getFolders() throws ExchangeException {
        java.util.ArrayList<Folder> folders = new java.util.ArrayList<Folder>();

        String msg =
        "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
        + "<soap:Envelope xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\">"
        + "<soap:Body>"
        + "<FindFolder Traversal=\"Shallow\" xmlns=\"http://schemas.microsoft.com/exchange/services/2006/messages\" xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\">"
        + "<FolderShape><t:BaseShape>Default</t:BaseShape></FolderShape>"
        + "<ParentFolderIds><t:FolderId Id=\"" + id + "\"/></ParentFolderIds>"
        + "</FindFolder>"
        + "</soap:Body>"
        + "</soap:Envelope>";

        org.w3c.dom.Document response = conn.execute(msg);
        for (org.w3c.dom.Node node : javaxt.xml.DOM.getElementsByTagName("Folder", response)){
            folders.add(new Folder(node, conn));
        }
        return folders.toArray(new Folder[folders.size()]);
    }


  //**************************************************************************
  //** createFolder
  //**************************************************************************
  /** Used to create a folder within this folder.
   */
    public Folder createFolder(String name) throws ExchangeException {
        name = name.trim();
        String msg =
        "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
        + "<soap:Envelope xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\">"
        + "<soap:Body>"
        + "<CreateFolder xmlns=\"http://schemas.microsoft.com/exchange/services/2006/messages\" xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\">"
        + "<ParentFolderId><t:FolderId Id=\"" + id + "\"/></ParentFolderId>"
        + "<Folders><t:Folder><t:DisplayName>" + name + "</t:DisplayName></t:Folder></Folders>"
        + "</CreateFolder>"
        + "</soap:Body>"
        + "</soap:Envelope>";

        org.w3c.dom.Document response = conn.execute(msg);
        org.w3c.dom.Node[] nodes = javaxt.xml.DOM.getElementsByTagName("Folder", response);
        if (nodes.length>0){
            Folder folder = new Folder(nodes[0], conn);
            folder.name = name;
            folder.totalCount = 0;
            folder.unreadCount = 0;
            return folder;
        }
        else throw new ExchangeException("Failed to create folder.");
    }


  //**************************************************************************
  //** delete
  //**************************************************************************
  /** Used to delete this folder.
   */
    public void delete() throws ExchangeException {
        String msg =
        "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
        + "<soap:Envelope xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\">"
        + "<soap:Body>"
        + "<DeleteFolder DeleteType=\"HardDelete\" xmlns=\"http://schemas.microsoft.com/exchange/services/2006/messages\" xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\">"
        + "<FolderIds><t:FolderId Id=\"" + id + "\"/></FolderIds>"
        + "</DeleteFolder>"
        + "</soap:Body>"
        + "</soap:Envelope>";
        conn.execute(msg);
    }



    private void parseXML(org.w3c.dom.Document xml) throws ExchangeException {
        org.w3c.dom.Node[] nodes = javaxt.xml.DOM.getElementsByTagName("FolderId", xml);
        if (nodes.length>0) parseFolderNode(nodes[0].getParentNode());
        else throw new ExchangeException("Failed to parse FolderId.");
    }


    private void parseFolderNode(org.w3c.dom.Node folder) throws ExchangeException {
        org.w3c.dom.NodeList properties = folder.getChildNodes();
        for (int i=0; i<properties.getLength(); i++){
            org.w3c.dom.Node property = properties.item(i);
            if (property.getNodeType()==1){
                String key = property.getNodeName();
                String value = property.getTextContent().trim();
                if (key.contains(":")) key = key.substring(key.indexOf(":")+1);
                if (key.equalsIgnoreCase("FolderId")){
                    id = javaxt.xml.DOM.getAttributeValue(property, "Id");
                    changeKey = javaxt.xml.DOM.getAttributeValue(property, "ChangeKey");
                }
                if (key.equalsIgnoreCase("DisplayName")) name = value;
                if (key.equalsIgnoreCase("TotalCount")) totalCount = cint(value);
                //if (key.equalsIgnoreCase("ChildFolderCount")) totalCount = cint(value);
                if (key.equalsIgnoreCase("UnreadCount")) unreadCount = cint(value);
            }
        }

        if (id==null || id.length()==0) throw new ExchangeException("Failed to parse Folder.");
    }


    private Integer cint(String str){
        try{
            return Integer.parseInt(str);
        }
        catch(Exception e){
            return null;
        }
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
    */


    public String getName(){
        return name;
    }

    public Integer getTotalCount(){
        return totalCount;
    }
    
    public Integer getUnreadCount(){
        return unreadCount;
    }

    public String toString(){
        return name;
    }



  //**************************************************************************
  //** getItems
  //**************************************************************************
  /** Returns an XML document with shallow representations of items found in
   *  this folder.
   *  @param maxEntries Maximum number of items to return.
   *  @param offset Item offset. 0 implies no offset.
   */
    protected org.w3c.dom.Document getItems(int maxEntries, int offset, java.util.ArrayList<String> additionalProperties, String orderBy) throws ExchangeException {
        return getItems("<m:IndexedPageItemView MaxEntriesReturned=\"" + maxEntries + "\" Offset=\"" + offset + "\" BasePoint=\"Beginning\"/>", additionalProperties, orderBy);
    }


  //**************************************************************************
  //** getItems
  //**************************************************************************
  /** Returns an XML document with shallow representations of items found in
   *  this folder.
   *
   *  @param view XML node representing a page view (e.g. IndexedPageItemView,
   *  FractionalPageItemView, CalendarView, ContactsView).
   *
   *  @param additionalProperties By default, this method returns a shallow
   *  representation of each item found in this folder. You can retrieve
   *  additional attributes by providing a list of properties
   *  (e.g. "calendar:TimeZone", "item:Sensitivity", etc).
   *
   *  @param orderBy SQL-style order by clause used to sort the results
   *  (e.g. "item:DateTimeReceived DESC").
   */
    protected org.w3c.dom.Document getItems(String view, java.util.ArrayList<String> additionalProperties, String orderBy) throws ExchangeException {

        String sort = "";
        if (orderBy!=null){

            for (String str : orderBy.split(",")){
                str = str.trim();
                String direction = "Ascending";
                if (str.toUpperCase().endsWith(" ASC")){
                    str = str.substring(0, str.lastIndexOf(" "));
                }
                else if (str.toUpperCase().endsWith(" DESC")){
                    str = str.substring(0, str.lastIndexOf(" "));
                    direction = "Descending";
                }
                if (str.length()>0){
                    sort += "<t:FieldOrder Order=\"" + direction + "\">"
                    + "<t:FieldURI FieldURI=\"" + str + "\" /></t:FieldOrder>";
                }
            }

            if (sort.length()>0){
                sort = "<m:SortOrder>" + sort + "</m:SortOrder>";
            }
        }


      //Update the view xml node. Make sure the node name is prefixed with a "m:" namespace
        if (view==null) view = "";
        else{
            view = view.trim();

            String nodeName = view.substring(1, view.indexOf(">"));
            if (nodeName.endsWith("/")) nodeName = nodeName.substring(0, nodeName.length()-1);
            if (nodeName.contains(" ")) nodeName = nodeName.substring(0, nodeName.indexOf(" "));
            nodeName = nodeName.trim();
            if (nodeName.contains(":")){
                String ns = nodeName.substring(0, nodeName.indexOf(":"));
                if (!ns.equals("m")){
                    String newNodeName = nodeName.substring(ns.length()+1);
                    view = view.replace("<" + nodeName, "<m:" + newNodeName);
                    view = view.replace("</" + nodeName, "</m:" + newNodeName);
                }
            }
            else{
                view = view.replace("<" + nodeName, "<m:" + nodeName);
                view = view.replace("</" + nodeName, "</m:" + nodeName);
            }
        }

        StringBuffer props = new StringBuffer();
        if (additionalProperties!=null){
            for (String prop : additionalProperties){
                props.append("<t:FieldURI FieldURI=\"" + prop + "\"/>");
            }
        }

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
                + props.toString() 
                + "</t:AdditionalProperties>"
        + "</m:ItemShape>"

        + view
        + sort

        + "<m:ParentFolderIds>"
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