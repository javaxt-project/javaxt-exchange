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
    private String parentID;
    private String name;
    private String changeKey;
    private Integer totalCount;
    private Integer unreadCount;
    private Integer folderCount;
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
  /** Creates a new instance of this class using a folder
   */
    protected Folder(Folder folder){
        this.id = folder.id;
        this.parentID = folder.parentID;
        this.name = folder.name;
        this.changeKey = folder.changeKey;
        this.totalCount = folder.totalCount;
        this.unreadCount = folder.unreadCount;
        this.folderCount = folder.folderCount;
        this.conn = folder.conn;
    }


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
        + "<FolderShape>"
            + "<t:BaseShape>Default</t:BaseShape>"
            + "<t:AdditionalProperties>"

            //The following returns the LastModifiedTime for the folder. Unfortunately,
            //this does not reflect changes made to items in this folder.
            //+ "<t:ExtendedFieldURI PropertyTag=\"0x3008\" PropertyType=\"SystemTime\" />"

            + "<t:FieldURI FieldURI=\"folder:ParentFolderId\" />"
            + "</t:AdditionalProperties>"

        + "</FolderShape>"
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
        return getFolders("Shallow");
    }


  //**************************************************************************
  //** getFolders
  //**************************************************************************
  /** Returns an array of folders found in this folder. Returns a zero length
   *  array of no folders are found.
   *  @param Traversal Possible values include "Deep", "Shallow", "SoftDeleted".
   *  Defaults to "Shallow" if a null or invalid string is used.
   */
    public Folder[] getFolders(String Traversal) throws ExchangeException {
        java.util.ArrayList<Folder> folders = new java.util.ArrayList<Folder>();

        if (Traversal==null) Traversal = "";
        if (Traversal.equalsIgnoreCase("Deep")) Traversal = "Deep";
        else if (Traversal.equalsIgnoreCase("SoftDeleted")) Traversal = "SoftDeleted";
        else Traversal = "Shallow";


        String msg =
        "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
        + "<soap:Envelope xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\">"
        + "<soap:Body>"
        + "<FindFolder Traversal=\"" + Traversal + "\" xmlns=\"http://schemas.microsoft.com/exchange/services/2006/messages\" xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\">"
        + "<FolderShape>"
            + "<t:BaseShape>Default</t:BaseShape>"
            + "<t:AdditionalProperties>"
            + "<t:FieldURI FieldURI=\"folder:ParentFolderId\" />"
            + "</t:AdditionalProperties>"
        + "</FolderShape>"
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
            folder.folderCount = 0;
            folder.parentID = id;
            return folder;
        }
        else throw new ExchangeException("Failed to create folder.");
    }


  //**************************************************************************
  //** rename
  //**************************************************************************
  /** Used to rename this folder.
   */
    public void rename(String name) throws ExchangeException {
        changeKey = getChangeKey(conn);
        name = name.trim();
        String msg =
        "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
        + "<soap:Envelope xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\">"
        + "<soap:Body>"
        + "<UpdateFolder xmlns=\"http://schemas.microsoft.com/exchange/services/2006/messages\" xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\">"
        + "<FolderChanges>"
        + "<t:FolderChange>"
                + "<t:FolderId Id=\"" + id + "\" ChangeKey=\"" + changeKey + "\"/>"
                + "<t:Updates>"
                    + "<t:SetFolderField>"
                    + "<t:FieldURI FieldURI=\"folder:DisplayName\" />"
                    + "<t:Folder><t:DisplayName>" + name + "</t:DisplayName></t:Folder>"
                    + "</t:SetFolderField>"
                + "</t:Updates>"
        + "</t:FolderChange>"
        + "</FolderChanges>"
        + "</UpdateFolder>"
        + "</soap:Body>"
        + "</soap:Envelope>";
        conn.execute(msg);
        this.name = name;
    }


  //**************************************************************************
  //** move
  //**************************************************************************
  /** Used to move this folder to another folder.
   */
    public void move(Folder destination) throws ExchangeException {
        String msg =
        "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
        + "<soap:Envelope xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\">"
        + "<soap:Body>"
        + "<MoveFolder xmlns=\"http://schemas.microsoft.com/exchange/services/2006/messages\" xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\">"
        + "<ToFolderId><t:FolderId Id=\"" + destination.id + "\"/></ToFolderId>"
        + "<FolderIds><t:FolderId Id=\"" + id + "\"/></FolderIds>"
        + "</MoveFolder>"
        + "</soap:Body>"
        + "</soap:Envelope>";
        conn.execute(msg);
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


  //**************************************************************************
  //** getChangeKey
  //**************************************************************************
  /** Used to retrieve the latest ChangeKey for this folder. This method is
   *  required to update an item.
   */
    protected String getChangeKey(Connection conn) throws ExchangeException {
        changeKey = new Folder(id, conn).changeKey;
        return changeKey;
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
                else if(key.equalsIgnoreCase("ParentFolderId")){
                    parentID = javaxt.xml.DOM.getAttributeValue(property, "Id");
                    //javaxt.xml.DOM.getAttributeValue(property, "ChangeKey");
                }
                else if(key.equalsIgnoreCase("DisplayName")) name = value;
                else if(key.equalsIgnoreCase("TotalCount")) totalCount = cint(value);
                else if(key.equalsIgnoreCase("ChildFolderCount")) folderCount = cint(value);
                else if(key.equalsIgnoreCase("UnreadCount")) unreadCount = cint(value);
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

    public String getID(){
        return id;
    }


  //**************************************************************************
  //** getParentID
  //**************************************************************************
  /** Returns the unique ID for the parent folder. */

    public String getParentID(){
        return parentID;
    }

    /*
    public String getChangeKey(){
        return changeKey;
    }
    */

  //**************************************************************************
  //** getName
  //**************************************************************************
  /** Returns the name of the folder. */

    public String getName(){
        return name;
    }


  //**************************************************************************
  //** getTotalCount
  //**************************************************************************
  /** Returns the total number of items found in this folder. */

    public Integer getTotalCount(){
        return totalCount;
    }


  //**************************************************************************
  //** getUnreadCount
  //**************************************************************************
  /** Returns the total number of unread items found in this folder. */

    public Integer getUnreadCount(){
        return unreadCount;
    }


  //**************************************************************************
  //** getChildFolderCount
  //**************************************************************************
  /** Returns the total number of folders found in this this folder. */

    public Integer getChildFolderCount(){
        return folderCount;
    }


  //**************************************************************************
  //** toString
  //**************************************************************************
  /** Returns the name of the folder. */

    public String toString(){
        return name;
    }


  //**************************************************************************
  //** getItems
  //**************************************************************************
  /** Returns an XML document with shallow representations of items found in
   *  this folder.
   *  @param offset Item offset. 0 implies no offset.
   *  @param limit Maximum number of items to return.
   */
    protected org.w3c.dom.Document getItems(int offset, int limit, java.util.HashSet<FieldURI> additionalProperties, String where, FieldOrder[] sortOrder) throws ExchangeException {
        if (offset<1) offset = 0;
        if (limit<1) limit = 1;
        return getItems("<m:IndexedPageItemView MaxEntriesReturned=\"" + limit + "\" Offset=\"" + offset + "\" BasePoint=\"Beginning\"/>",
            additionalProperties, where, sortOrder);
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
   *  @param sortOrder SQL-style order by clause used to sort the results
   *  (e.g. "item:DateTimeReceived DESC").
   */
    protected org.w3c.dom.Document getItems(String view, java.util.HashSet<FieldURI> additionalProperties, String where, FieldOrder[] sortOrder) throws ExchangeException {


      //Parse order by statement
        String sort = "";
        if (sortOrder!=null){
            for (FieldOrder field : sortOrder){
                sort += field.toXML();
            }
            if (sort.length()>0){
                sort = "<m:SortOrder>" + sort + "</m:SortOrder>";
            }
        }


      //Parse where clasue and create restriction
        if (where==null) where = "";
        else where = where.trim();
        String restriction = where;


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

      //Pa
        StringBuffer props = new StringBuffer();
        if (additionalProperties!=null){
            java.util.Iterator<FieldURI> it = additionalProperties.iterator();
            while (it.hasNext()){
                FieldURI prop = it.next();
                if (prop!=null){
                    props.append(prop.toXML("t")); //("<t:FieldURI FieldURI=\"" + prop + "\"/>");
                }
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
        + restriction
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

        java.util.HashMap<String, javaxt.utils.Date> index = new java.util.HashMap<String, javaxt.utils.Date>();

        int offset = 0;
        int limit = 1000;


        while (true){


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
            + "<m:IndexedPageItemView MaxEntriesReturned=\"" + limit + "\" Offset=\"" + offset + "\" BasePoint=\"Beginning\"/>"


            + "<m:ParentFolderIds>"
            + "<t:FolderId Id=\"" + id + "\"/>"
            + "</m:ParentFolderIds>"
            + "</m:FindItem>"
            + "</soap:Body>"
            + "</soap:Envelope>";

            org.w3c.dom.Document xml = conn.execute(msg);

            org.w3c.dom.Node[] items = javaxt.xml.DOM.getElementsByTagName("Items", xml);
            int numItems = 0;
            if (items.length>0){
                org.w3c.dom.NodeList nodes = items[0].getChildNodes();
                for (int i=0; i<nodes.getLength(); i++){
                    org.w3c.dom.Node node = nodes.item(i);
                    if (node.getNodeType()==1){
                        FolderItem item = new FolderItem(node);
                        index.put(item.getID(), item.getLastModifiedTime());
                        numItems++;
                    }
                }

            }

            offset+=numItems;

            //System.out.println("Found " + numItems + " items");

            if (numItems==0 || numItems<limit) break;
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