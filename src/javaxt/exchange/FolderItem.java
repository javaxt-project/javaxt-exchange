package javaxt.exchange;

//******************************************************************************
//**  FolderItem Class
//******************************************************************************
/**
 *   Represents an item in a folder. This is a base class for contacts, calendar
 *   events, etc.
 *
 ******************************************************************************/

public class FolderItem {

/*
  //This is an ordered list of item properties.
    private String MimeContent
    private String ItemId
    private String ParentFolderId
    private String ItemClass
    private String Subject
    private String Sensitivity
    private String Body
    private String Attachments
    private String DateTimeReceived
    private String Size
    private String Categories
    private String Importance
    private String InReplyTo
    private String IsSubmitted
    private String IsDraft
    private String IsFromMe
    private String IsResend
    private String IsUnmodified
    private String InternetMessageHeaders
    private String DateTimeSent
    private String DateTimeCreated
    private String ResponseObjects
    private String ReminderDueBy
    private String ReminderIsSet
    private String ReminderMinutesBeforeStart
    private String DisplayCc
    private String DisplayTo
    private String HasAttachments
    private String ExtendedProperty
    private String Culture
    private String EffectiveRights
    private String LastModifiedName
    private String LastModifiedTime
    private String IsAssociated
    private String WebClientReadFormQueryString
    private String WebClientEditFormQueryString
    private String ConversationId
    private String UniqueBody
 */

    protected String id;
    protected String changeKey;
    protected String subject;
    protected String body;
    protected java.util.HashSet<String> categories = new java.util.HashSet<String>();
    protected javaxt.utils.Date lastModified;
    protected java.util.HashMap<String, String> updates = new java.util.HashMap<String, String>();

    private org.w3c.dom.Node node;

    
    protected FolderItem(){}


  //**************************************************************************
  //** Constructor
  //**************************************************************************
  /** Creates a new instance of this class
   */
    protected FolderItem(String exchangeID, Connection conn) throws ExchangeException{
        String msg =
        "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
        + "<soap:Envelope xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\" "
            + "xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\" "
            + "xmlns:m=\"http://schemas.microsoft.com/exchange/services/2006/messages\">"
        //+ "<soap:Header><t:RequestServerVersion Version=\"Exchange2007\"/></soap:Header>"
        + "<soap:Body>"
        + "<m:GetItem>"
        + "<m:ItemShape>"
                + "<t:BaseShape>AllProperties</t:BaseShape>"
                + "<t:AdditionalProperties>"
                + "<t:FieldURI FieldURI=\"item:ItemClass\"/>"
                //+ "<t:FieldURI FieldURI=\"item:LastModifiedTime\"/>" //value="item:LastModifiedTime" //<--This doesn't work...
                + "<t:ExtendedFieldURI PropertyTag=\"0x3008\" PropertyType=\"SystemTime\" />" //<--This returns the LastModifiedTime!
                + "</t:AdditionalProperties>"
        + "</m:ItemShape>"
        + "<m:ItemIds><t:ItemId Id=\"" + exchangeID + "\"/></m:ItemIds>"
        + "</m:GetItem>"
        + "</soap:Body>"
        + "</soap:Envelope>";


        org.w3c.dom.Document xml = conn.execute(msg);

        org.w3c.dom.Node[] items = javaxt.xml.DOM.getElementsByTagName("Items", xml);
        boolean foundItem = false;
        if (items.length>0){
            org.w3c.dom.NodeList nodes = items[0].getChildNodes();
            for (int i=0; i<nodes.getLength(); i++){
                org.w3c.dom.Node node = nodes.item(i);
                if (node.getNodeType()==1){
                    parseNode(node);
                    foundItem = true;
                }
            }

        }

        if (!foundItem) throw new ExchangeException("Failed to find item " + exchangeID);
    }


  //**************************************************************************
  //** Constructor
  //**************************************************************************
  /** Creates a new instance of this class using
   *  @param node Item node (e.g. "Contact", "CalendarItem", etc).
   */
    protected FolderItem(org.w3c.dom.Node node) {
        parseNode(node);
    }


  //**************************************************************************
  //** parseNode
  //**************************************************************************
  /** Used to parse an xml node with item information.
   */
    private void parseNode(org.w3c.dom.Node node){
        this.node = node;
        org.w3c.dom.NodeList outerNodes = node.getChildNodes();
        for (int i=0; i<outerNodes.getLength(); i++){
            org.w3c.dom.Node outerNode = outerNodes.item(i);
            if (outerNode.getNodeType()==1){
                String nodeName = outerNode.getNodeName();
                if (nodeName.contains(":")) nodeName = nodeName.substring(nodeName.indexOf(":")+1);
                if (nodeName.equalsIgnoreCase("ItemId")){
                    id = javaxt.xml.DOM.getAttributeValue(outerNode, "Id");
                    changeKey = javaxt.xml.DOM.getAttributeValue(outerNode, "ChangeKey");
                }
                else if(nodeName.equalsIgnoreCase("Subject")){
                    subject = javaxt.xml.DOM.getNodeValue(outerNode);
                }
                else if(nodeName.equalsIgnoreCase("Body")){
                    body = javaxt.xml.DOM.getNodeValue(outerNode);
                }
                else if (nodeName.equalsIgnoreCase("Categories")){
                    org.w3c.dom.NodeList childNodes = outerNode.getChildNodes();
                    for (int j=0; j<childNodes.getLength(); j++){
                        org.w3c.dom.Node childNode = childNodes.item(j);
                        if (childNode.getNodeType()==1){
                            categories.add(childNode.getTextContent());
                        }
                    }
                }
                else if(nodeName.equalsIgnoreCase("LastModifiedTime")){
                    try{
                        lastModified = new javaxt.utils.Date(javaxt.xml.DOM.getNodeValue(outerNode));
                    }
                    catch(java.text.ParseException e){}
                }
                else if (nodeName.equalsIgnoreCase("ExtendedProperty")){

                    org.w3c.dom.Node ExtendedFieldURI = null;
                    org.w3c.dom.Node Value = null;

                    org.w3c.dom.NodeList childNodes = outerNode.getChildNodes();
                    for (int j=0; j<childNodes.getLength(); j++){
                        org.w3c.dom.Node childNode = childNodes.item(j);
                        if (childNode.getNodeType()==1){

                            String childNodeName = childNode.getNodeName();
                            if (childNodeName.contains(":")) childNodeName = childNodeName.substring(childNodeName.indexOf(":")+1);

                            if (childNodeName.equalsIgnoreCase("ExtendedFieldURI")){
                                ExtendedFieldURI = childNode;
                            }
                            else if (childNodeName.equalsIgnoreCase("Value")){
                                Value = childNode;
                            }
                        }
                    }

                    if (ExtendedFieldURI!=null){
                        String PropertyTag = javaxt.xml.DOM.getAttributeValue(ExtendedFieldURI, "PropertyTag");

                      //Extract last mod date
                        if (PropertyTag.equalsIgnoreCase("0x3008")){
                            try{
                                lastModified = new javaxt.utils.Date(javaxt.xml.DOM.getNodeValue(Value));
                            }
                            catch(java.text.ParseException e){}
                        }
                    }
                }

            }
        }
    }


  //**************************************************************************
  //** getChildNodes
  //**************************************************************************
  /** Returns the children of the xml node used to instantiate this class.
   *  Classes that extend this class use the child nodes to extract additional
   *  item information.
   */
    protected org.w3c.dom.NodeList getChildNodes(){
        return this.node.getChildNodes();
    }


  //**************************************************************************
  //** resetUpdates
  //**************************************************************************
  /** Calls to 'add' and 'set' methods in this class are recorded in a hashmap.
   *  The hashmap is later used when updating a contact via the save() method.
   *  Use this method to reset the list of updates.
   */
    protected void resetUpdates(){
        updates.clear();
    }
    
  //**************************************************************************
  //** getID
  //**************************************************************************
  /** Used to get the unique id associated with this item.
   */
    public String getID(){
        return id;
    }


  //**************************************************************************
  //** setID
  //**************************************************************************
  /** Used to set/update the id associated with this item. This method is
   *  available to application developers who with to extend this class.
   */
    protected void setID(String id){
        if (id!=null){
            id = id.trim();
            if (id.length()<25) id = null;
        }
        this.id = id;
    }


  //**************************************************************************
  //** getLastModifiedTime
  //**************************************************************************
  /** Returns the timestamp for when this item was last modified. Note that
   *  timestamp is set in the constructor and may not reflect the most recent
   *  value.
   */
    public javaxt.utils.Date getLastModifiedTime(){
        return lastModified;
    }


    public void setLastModifiedTime(javaxt.utils.Date lastModified){
        this.lastModified = lastModified;
    }

    
  //**************************************************************************
  //** getChangeKey
  //**************************************************************************
  /** Used to retrieve the ChangeKey for this contact. Note that the ChangeKey 
   *  is set in the constructor and may not reflect the most recent value.
   */
    public String getChangeKey(){
        return changeKey;
    }


  //**************************************************************************
  //** getChangeKey
  //**************************************************************************
  /** Used to retrieve the latest ChangeKey for this contact. This method is
   *  required to update an item.
   */
    protected String getChangeKey(Connection conn) throws ExchangeException {
        changeKey = new FolderItem(id, conn).changeKey;
        return changeKey;
    }


  //**************************************************************************
  //** getSubject
  //**************************************************************************
  /** Returns the subject of this item. Usually, this only applies to mail
   *  and calendar entries.
   */
    protected String getSubject(){
        subject = getValue(subject);
        return subject;
    }


    protected void setSubject(String subject){
        subject = getValue(subject);

        if (id!=null) {
            if (subject==null && this.subject!=null) updates.put("Subject", null);
            if (subject!=null && !subject.equals(this.subject)) updates.put("Subject", subject);
        }

        this.subject = subject;
    }


  //**************************************************************************
  //** getBody
  //**************************************************************************
  /** Returns the body of this item. Usually, this only applies to mail items
   *  and calendar entries.
   */
    protected String getBody(){
        body = getValue(body);
        return body;
    }

    protected void setBody(String body){
        body = getValue(body);

        if (id!=null) {
            if (body==null && this.body!=null) updates.put("Body", null);
            if (body!=null && !body.equals(this.body)) updates.put("Body", body);
        }

        this.body = body;
    }


  //**************************************************************************
  //** getCategories
  //**************************************************************************
  /** Returns a list of categories associated with this item.
   */
    public String[] getCategories(){
        return categories.toArray(new String[categories.size()]);
    }


  //**************************************************************************
  //** setCategories
  //**************************************************************************
  /** Used to add categories to a contact.
   */
    public void setCategories(String[] categories){

        if (categories==null || categories.length==0){
            removeCategories();
            return;
        }

      //See whether any updates are required
        int numMatches = 0;
        if (categories.length==this.categories.size()){
            for (String category : categories){
                if (category!=null){
                    if (this.categories.contains(category)) numMatches++;
                }
            }
        }

      //If the input array equals the current list of categories, do nothing...
        if (numMatches==categories.length){
            return;
        }
        else {
            this.categories.clear();
            for (String category : categories){
                addCategory(category);
            }
        }
    }


    public void setCategory(String category){
        setCategories(new String[]{category});
    }

  //**************************************************************************
  //** addCategory
  //**************************************************************************
  /** Used to add a category to a contact.
   */
    public void addCategory(String category){

        if (category!=null){
            category = category.trim();
            if (category.length()==0) category = null;
        }
        if (category==null) return;


        if (id!=null && !categories.contains(category)){

            categories.add(category);

            StringBuffer xml = new StringBuffer();
            java.util.Iterator<String> it = categories.iterator();
            while (it.hasNext()){
                xml.append("<t:String>" + it.next() + "</t:String>");
            }
            updates.put("Categories", xml.toString());
        }
        else{
            categories.add(category);
        }
    }


  //**************************************************************************
  //** removeCategory
  //**************************************************************************
  /**  Used to remove a category associated with this contact.
   */
    public void removeCategory(String category){

        if (category==null) return;
        category = category.trim();

        String obj = null;
        java.util.Iterator<String> it = categories.iterator();
        while (it.hasNext()){
            String val = it.next();
            if (val.equalsIgnoreCase(category)){
                obj = val;
                break;
            }
        }

        if (obj!=null){
            categories.remove(obj);
            setCategories(categories.toArray(new String[categories.size()]));
        }
    }


  //**************************************************************************
  //** removeCategories
  //**************************************************************************
  /**  Used to remove all categories associated with this contact.
   */
    public void removeCategories(){
        if (id!=null && !categories.isEmpty()){
            updates.put("Categories", null);
        }

        categories.clear();
    }


  //**************************************************************************
  //** delete
  //**************************************************************************
  /** Used to delete this item.
   */
    public void delete(Connection conn) throws ExchangeException {
        String msg =
        "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
        + "<soap:Envelope xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\" xmlns:m=\"http://schemas.microsoft.com/exchange/services/2006/messages\">"
        //+ "<soap:Header><t:RequestServerVersion Version=\"Exchange2007_SP1\"/></soap:Header>"
        + "<soap:Body>"
        + "<m:DeleteItem DeleteType=\"MoveToDeletedItems\">"
        + "<m:ItemIds>"
        + "<t:ItemId Id=\"" + id + "\"/></m:ItemIds>" //ChangeKey=\"EQAAABYAAAA9IPsqEJarRJBDywM9WmXKAAV+D0Dq\"
        + "</m:DeleteItem>"
        + "</soap:Body>"
        + "</soap:Envelope>";
        conn.execute(msg);
    }


  //**************************************************************************
  //** getValue
  //**************************************************************************
  /** Private method used to normalize a string. This method is called by all
   *  setters and getters that deal with string values.
   */
    protected String getValue(String val){
        if (val!=null){
            val = val.trim();
            if (val.length()==0) val = null;
        }
        return val;
    }

  //**************************************************************************
  //** formatDate
  //**************************************************************************
  /** Used to format a date into a string that Exchange Web Services can
   *  understand.
   */
    protected String formatDate(javaxt.utils.Date date){
        if (date==null) return null;
        String d = date.toString("yyyy-MM-dd HH:mm:ssZ").replace(" ", "T");
        return d.substring(0, d.length()-2) + ":" + d.substring(d.length()-2);
    }
}