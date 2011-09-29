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
    protected java.util.HashSet<String> categories = new java.util.HashSet<String>();
    protected java.util.HashMap<String, String> updates = new java.util.HashMap<String, String>();

    private org.w3c.dom.Node node;

    
    protected FolderItem(){}

  //**************************************************************************
  //** Constructor
  //**************************************************************************
  /** Creates a new instance of this class
   */
    protected FolderItem(String exchangeID, String tagName, Connection conn) throws ExchangeException{
        org.w3c.dom.Document xml = Folder.getItem(exchangeID, conn);
        org.w3c.dom.Node[] nodes = javaxt.xml.DOM.getElementsByTagName(tagName, xml);
        if (nodes.length>0) {
            parseNode(nodes[0]);
        }
    }


  //**************************************************************************
  //** Constructor
  //**************************************************************************
  /** Creates a new instance of FolderItem. */

    protected FolderItem(org.w3c.dom.Node node) {
        parseNode(node);
    }

    protected org.w3c.dom.NodeList getChildNodes(){
        return this.node.getChildNodes();
    }

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

    
    
    public void setExchangeID(String id){
        if (id!=null){
            id = id.trim();
            if (id.length()<25) id = null;
        }
        this.id = id;
    }


    public String getExchangeID(){
        return id;
    }


    public String getSubject(){
        return subject;
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
  //** getCategories
  //**************************************************************************
  /** Used to add categories to a contact.
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


}