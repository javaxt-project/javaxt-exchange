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
    //private String changeKey;
    protected String subject;
    protected String body;
    protected java.util.HashSet<String> categories = new java.util.HashSet<String>();
    protected java.util.HashMap<String, String> updates = new java.util.HashMap<String, String>();

    private org.w3c.dom.Node node;

    
    protected FolderItem(){}

  //**************************************************************************
  //** Constructor
  //**************************************************************************
  /** Creates a new instance of this class
   */
    protected FolderItem(String exchangeID, Connection conn) throws ExchangeException{
        this(Folder.getItem(exchangeID, conn));
    }


  //**************************************************************************
  //** Constructor
  //**************************************************************************
  /** Creates a new instance of this class using item node (e.g. "Contact", 
   * "CalendarItem", etc).
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
                    //changeKey = javaxt.xml.DOM.getAttributeValue(outerNode, "ChangeKey");
                }
                else if(nodeName.equalsIgnoreCase("Subject")){
                    subject = javaxt.xml.DOM.getNodeValue(outerNode);
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
  //** getExchangeID
  //**************************************************************************
  /** Used to get the unique id associated with this item.
   */
    public String getExchangeID(){
        return id;
    }


  //**************************************************************************
  //** setExchangeID
  //**************************************************************************
  /** Used to set/update the id associated with this item.
   */
    public void setExchangeID(String id){
        if (id!=null){
            id = id.trim();
            if (id.length()<25) id = null;
        }
        this.id = id;
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



  //**************************************************************************
  //** getChangeKey
  //**************************************************************************
  /** Used to retrieve the latest ChangeKey for this contact. The ChangeKey is
   *  required to update an existing contact.
   */
    protected String getChangeKey(Connection conn) throws ExchangeException {

        org.w3c.dom.Node node = Folder.getItem(id, conn);
        org.w3c.dom.NodeList outerNodes = node.getChildNodes();
        for (int j=0; j<outerNodes.getLength(); j++){
            org.w3c.dom.Node outerNode = outerNodes.item(j);
            if (outerNode.getNodeType()==1){
                String nodeName = outerNode.getNodeName();
                if (nodeName.contains(":")) nodeName = nodeName.substring(nodeName.indexOf(":")+1);
                if (nodeName.equalsIgnoreCase("ItemId")){
                    return javaxt.xml.DOM.getAttributeValue(outerNode, "ChangeKey");
                }
            }
        }

        throw new ExchangeException("Failed to retrieve ChangeKey");

    }
}