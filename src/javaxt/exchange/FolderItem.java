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
    protected String parentFolderID;
    protected String subject;
    protected String body;
    protected String bodyType;
    protected boolean hasAttachments = false;
    protected java.util.HashSet<String> categories = new java.util.HashSet<String>();
    protected java.util.HashSet<Attachment> attachments = new java.util.HashSet<Attachment>();
    protected javaxt.utils.Date lastModified;
    protected java.util.HashMap<String, ExtendedProperty> extendedProperties =
            new java.util.HashMap<String, ExtendedProperty>();

    /** Hashmap with a list of any pending updates to be made to this item. */
    protected java.util.HashMap<String, Object> updates = new java.util.HashMap<String, Object>();

    private org.w3c.dom.Node node;

    
    protected FolderItem(){}


  //**************************************************************************
  //** Constructor
  //**************************************************************************
  /** Creates a new instance of this class
   *  @param AdditionalProperties A list of attributes you wish to see
   */
    protected FolderItem(String exchangeID, Connection conn, ExtendedProperty[] AdditionalProperties) throws ExchangeException{

        if (exchangeID==null) throw new ExchangeException("Exchange ID is required.");
        if (conn==null) throw new ExchangeException("Exchange Web Services Connection is required.");

        StringBuffer str = new StringBuffer();

        str.append("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
        str.append("<soap:Envelope xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\" "
            + "xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\" "
            + "xmlns:m=\"http://schemas.microsoft.com/exchange/services/2006/messages\">");
        //+ "<soap:Header><t:RequestServerVersion Version=\"Exchange2007\"/></soap:Header>"
        str.append("<soap:Body>");
        str.append("<m:GetItem>");
        str.append("<m:ItemShape>");
        str.append("<t:BaseShape>AllProperties</t:BaseShape>");


        str.append("<t:AdditionalProperties>");
        //str.append("<t:FieldURI FieldURI=\"item:LastModifiedTime\"/>"); //<--This doesn't work...
        str.append("<t:ExtendedFieldURI PropertyTag=\"0x3008\" PropertyType=\"SystemTime\" />"); //<--This returns the LastModifiedTime!

        if (AdditionalProperties!=null){            
            for (ExtendedProperty property : AdditionalProperties){

                String name = property.getName();
                String type = property.getType();
                String guid = property.getID();

                String idAttr = (guid==null ? "" : "PropertySetId=\"" + guid + "\"");
                String nameAttr = (name.startsWith("0x") ? "PropertyTag=\"" + name + "\"" : "PropertyName=\"" + name + "\"");
                String typeAttr = "PropertyType=\"" + type + "\"";

                str.append("<t:ExtendedFieldURI " + nameAttr + " " + idAttr + " " + typeAttr + "/>");
            }
        }
        str.append("</t:AdditionalProperties>");


        str.append("</m:ItemShape>");
        str.append("<m:ItemIds><t:ItemId Id=\"" + exchangeID + "\"/></m:ItemIds>");
        str.append("</m:GetItem>");
        str.append("</soap:Body>");
        str.append("</soap:Envelope>");



        org.w3c.dom.Document xml = conn.execute(str.toString());
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
                if (nodeName.equalsIgnoreCase("ParentFolderId")){
                    parentFolderID = javaxt.xml.DOM.getAttributeValue(outerNode, "Id");
                    //changeKey = javaxt.xml.DOM.getAttributeValue(outerNode, "ChangeKey");
                }
                else if(nodeName.equalsIgnoreCase("Subject")){
                    subject = javaxt.xml.DOM.getNodeValue(outerNode);
                }
                else if(nodeName.equalsIgnoreCase("Body")){
                    body = javaxt.xml.DOM.getNodeValue(outerNode);
                    bodyType = javaxt.xml.DOM.getAttributeValue(outerNode, "BodyType");
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
                else if(nodeName.equalsIgnoreCase("HasAttachments")){
                    hasAttachments = javaxt.xml.DOM.getNodeValue(outerNode).equalsIgnoreCase("true");
                }
                else if (nodeName.equalsIgnoreCase("Attachments")){
                    org.w3c.dom.NodeList childNodes = outerNode.getChildNodes();
                    for (int j=0; j<childNodes.getLength(); j++){
                        org.w3c.dom.Node childNode = childNodes.item(j);
                        if (childNode.getNodeType()==1){
                            attachments.add(new Attachment(childNode));
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

                    ExtendedProperty prop = new ExtendedProperty(outerNode);

                    if (prop.getName().equalsIgnoreCase("0x3008")){
                        try{
                            lastModified = new javaxt.utils.Date(prop.getValue());
                        }
                        catch(java.text.ParseException e){}
                    }
                    else{
                        extendedProperties.put(prop.getName(), prop);
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
  //** getParentFolderID
  //**************************************************************************
  /** Returns the unique id of the item's folder.
   */
    public String getParentFolderID(){
        return parentFolderID;
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


  //**************************************************************************
  //** getLastModifiedTime
  //**************************************************************************
  /** Returns the timestamp for when this item was last modified.
   */
    protected javaxt.utils.Date getLastModifiedTime(Connection conn) throws ExchangeException {
        this.lastModified = new FolderItem(id, conn, null).getLastModifiedTime();
        return lastModified;
    }


    public void setLastModifiedTime(javaxt.utils.Date lastModified){
        this.lastModified = lastModified;
    }

    
  //**************************************************************************
  //** getChangeKey
  //**************************************************************************
  /** Used to retrieve the ChangeKey for this item. Note that the ChangeKey
   *  is set in the constructor and may not reflect the most recent value.
   */
    public String getChangeKey(){
        return changeKey;
    }


  //**************************************************************************
  //** getChangeKey
  //**************************************************************************
  /** Used to retrieve the latest ChangeKey for this item. This method is
   *  required to update an item.
   */
    protected String getChangeKey(Connection conn) throws ExchangeException {
        changeKey = new FolderItem(id, conn, null).changeKey;
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


  //**************************************************************************
  //** setBody
  //**************************************************************************
  /** @param format Text format ("Best", "HTML", or "Text").
   */
    protected void setBody(String body, String format){
        body = getValue(body);
        format = getValue(format);
        if (format==null) format = "Text";
        else{
            format = format.trim();
            if (format.equalsIgnoreCase("Best")) format = "Best";
            else if (format.equalsIgnoreCase("HTML")) format = "HTML";
            else if (format.equalsIgnoreCase("Text")) format = "Text";
        }



        if (id!=null) {
            if (body==null && this.body!=null) updates.put("Body", null);
            if (body!=null && (!body.equals(this.body) || !format.equals(this.bodyType))){

                StringBuffer xml = new StringBuffer();
                xml.append("<t:SetItemField><t:FieldURI FieldURI=\"item:Body\" /><t:Message><t:Body BodyType=\"" + format + "\">");
                xml.append(body);
                xml.append("</t:Body></t:Message></t:SetItemField>");
                updates.put("Body", xml.toString());
            }
        }

        this.body = body;
        this.bodyType = format;
    }


  //**************************************************************************
  //** getBodyType
  //**************************************************************************
  /** Returns the text encoding used in the body of this item. Possible values
   *  include "Best", "HTML", or "Text".
   */
    protected String getBodyType(){
        bodyType = getValue(bodyType);
        return bodyType;
    }


  //**************************************************************************
  //** getCategories
  //**************************************************************************
  /** Returns a list of categories associated with this item. Returns null if
   *  no categories are found.
   */
    public String[] getCategories(){
        if (categories.isEmpty()) return null;
        else return categories.toArray(new String[categories.size()]);
    }


  //**************************************************************************
  //** hasAttachments
  //**************************************************************************
  /** Returns a boolean used to indicate whether this item has attachments.
   */
    public boolean hasAttachments(){
        return hasAttachments;
    }


  //**************************************************************************
  //** getAttachments
  //**************************************************************************
  /** Returns an array of attachments associated with this item. Returns null
   *  if there are no attachments.
   */
    public Attachment[] getAttachments(){
        if (attachments.isEmpty()) return null;
        else return attachments.toArray(new Attachment[attachments.size()]);
    }
    
    public void removeAttachment(Attachment attachment){
        attachments.remove(attachment);
    }
    
    public void addAttachment(Attachment attachment){
        attachments.add(attachment);
    }


  //**************************************************************************
  //** setCategories
  //**************************************************************************
  /** Used to define categories for this item.
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
  /** Used to add a category to this item.
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
  /**  Used to remove a category associated with this item.
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
  /**  Used to remove all categories associated with this item.
   */
    public void removeCategories(){
        if (id!=null && !categories.isEmpty()){
            updates.put("Categories", null);
        }

        categories.clear();
    }

    
    public ExtendedProperty getExtendedProperty(String name){
        return extendedProperties.get(name);
    }


  //**************************************************************************
  //** getExtendedProperties
  //**************************************************************************
  /** Returns an array of ExtendedProperties associated with this item.
   */
    public ExtendedProperty[] getExtendedProperties(){

      //Only include non-null values in the array
        java.util.ArrayList<ExtendedProperty> arr = new java.util.ArrayList<ExtendedProperty>();
        java.util.Iterator<String> it = extendedProperties.keySet().iterator();
        while (it.hasNext()){
            String key = it.next();
            ExtendedProperty val = extendedProperties.get(key);
            if (val!=null) arr.add(val);
        }
        if (arr.isEmpty()) return null;
        else return arr.toArray(new ExtendedProperty[arr.size()]);
    }


  //**************************************************************************
  //** setExtendedProperties
  //**************************************************************************
  /** Bulk update of the ExtendedProperties associated with this item.
   */
    public void setExtendedProperties(ExtendedProperty[] extendedProperties) {

        if (extendedProperties==null || extendedProperties.length==0){
            removeExtendedProperties();
            return;
        }

      //See whether any updates are required
        int numMatches = 0;
        int total = 0;
        for (ExtendedProperty property : extendedProperties){
            if (property!=null){
                total++;
                String name = property.getName();
                if (this.extendedProperties.containsKey(name)){
                    if (this.extendedProperties.get(name).equals(property)) numMatches++;
                }
            }
        }

        int numExtendedProperties = 0;
        if (this.getExtendedProperties()!=null) numExtendedProperties = this.getExtendedProperties().length;

      //If the input array equals the current list of extendedProperties, do nothing...
        if (numMatches==total && numMatches==numExtendedProperties){
            return;
        }
        else {
            this.extendedProperties.clear();
            for (ExtendedProperty property : extendedProperties){
                setExtendedProperty(property);
            }
        }
    }


  //**************************************************************************
  //** setExtendedProperty
  //**************************************************************************
  /** Used to add or update an ExtendedProperty associated with this item.
   */
    public void setExtendedProperty(ExtendedProperty property){

      //Check whether this is a new property
        boolean update = false;
        if (extendedProperties.containsKey(property.getName())){
            java.util.Iterator<String> it = extendedProperties.keySet().iterator();
            while (it.hasNext()){
                String key = it.next();
                if (key.equals(property.getName())){
                    String value = null;
                    if (extendedProperties.get(key).getValue()!=null){
                        value = extendedProperties.get(key).getValue();
                    }

                    if (value==null || !value.equals(property.getValue())){
                        update = true;
                    }
                    break;
                }
            }
        }
        else{
            update = true;
        }


        if (update){
            extendedProperties.put(property.getName(), property);
            if (id!=null) updates.put("ExtendedProperties", getExtendedPropertyUpdates());
        }
    }


  //**************************************************************************
  //** removeExtendedProperty
  //**************************************************************************
  /** Used to remove a specific ExtendedProperty associated with this item.
   */
    public void removeExtendedProperty(ExtendedProperty property){

        String name = property.getName();
        if (extendedProperties.containsKey(name)){
            if (id==null) extendedProperties.remove(name);
            else{
                property = extendedProperties.get(name);
                property.setValue(null);
                extendedProperties.put(name, property);
            }

            if (id!=null) updates.put("ExtendedProperties", getExtendedPropertyUpdates());
        }
    }


  //**************************************************************************
  //** removeExtendedProperties
  //**************************************************************************
  /** Removes all ExtendedProperties associated with this item.
   */
    public void removeExtendedProperties(){
        ExtendedProperty[] extendedProperties = getExtendedProperties();
        if (extendedProperties!=null){
            for (ExtendedProperty property : extendedProperties){
                removeExtendedProperty(property);
            }
        }
    }


  //**************************************************************************
  //** getExtendedPropertyUpdates
  //**************************************************************************
  /** Used to generate an XML fragment used to update or delete
   *  ExtendedProperties.
   */
    private String getExtendedPropertyUpdates(){

        if (extendedProperties.isEmpty()) return "";
        else{
            StringBuffer xml = new StringBuffer();

            java.util.Iterator<String> it = extendedProperties.keySet().iterator();
            while (it.hasNext()){
                
                ExtendedProperty property = extendedProperties.get(it.next());
                String value = property.getValue();

                if (value==null){
                    if (id!=null){
                        xml.append(property.toXML("t", "delete"));
                    }
                }
                else{
                    xml.append(property.toXML("t", "update"));
                }

            }
            return xml.toString();
        }
    }


  //**************************************************************************
  //** updateLastModifiedTime
  //**************************************************************************
  /**  Used to update the LastModified attribute in Exchange by updating or
   *   deleting the "Subject" attribute.
   *   @param itemName Name of the item (e.g. "Contact")
   */
    protected javaxt.utils.Date updateLastModifiedTime(String itemName, javaxt.exchange.Connection ews) throws ExchangeException {

        StringBuffer msg = new StringBuffer();
        msg.append("<?xml version=\"1.0\" encoding=\"utf-8\" ?>");
        msg.append("<soap:Envelope xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\" xmlns:m=\"http://schemas.microsoft.com/exchange/services/2006/messages\">");
        msg.append("<soap:Body>");
        msg.append("<m:UpdateItem ConflictResolution=\"AutoResolve\">");
        msg.append("<m:ItemChanges>");
        msg.append("<t:ItemChange>");
        msg.append("<t:ItemId Id=\"" + getID() + "\" ChangeKey=\"" + getChangeKey(ews) + "\" />");
        msg.append("<t:Updates>");

        String namespace = "item";
        String key = "Subject";
        String value = this.getSubject();

        if (value==null){
            msg.append("<t:DeleteItemField>");
            msg.append("<t:FieldURI FieldURI=\"" + namespace + ":" + key + "\"/>");
            msg.append("</t:DeleteItemField>");
        }
        else{
            msg.append("<t:SetItemField>");
            msg.append("<t:FieldURI FieldURI=\"" + namespace + ":" + key + "\" />");
            msg.append("<t:" + itemName + ">");
            msg.append("<t:" + key + ">" + value + "</t:" + key + ">");
            msg.append("</t:" + itemName + ">");
            msg.append("</t:SetItemField>");
        }
        msg.append("</t:Updates>");
        msg.append("</t:ItemChange>");
        msg.append("</m:ItemChanges>");
        msg.append("</m:UpdateItem>");
        msg.append("</soap:Body>");
        msg.append("</soap:Envelope>");

        ews.execute(msg.toString());

        return getLastModifiedTime(ews);
    }


  //**************************************************************************
  //** update
  //**************************************************************************
  /** Used to update this item using the "UpdateItem" web method.
   *  http://msdn.microsoft.com/en-us/library/aa580254%28v=exchg.140%29.aspx
   *
   *  @param itemName Name/Type of item being updated (e.g. "Message",
   *  "Contact", "CalendarItem", etc.). A complete list of items can be
   *  found here:
   *  http://msdn.microsoft.com/en-us/library/aa565652%28v=exchg.140%29.aspx
   *
   *  @param options Hashmap containing key/value pairs representing update
   *  options. Valid keys include "ConflictResolution", "MessageDisposition",
   *  and "SendMeetingInvitationsOrCancellations".
   */
    protected void update(String itemName, java.util.HashMap<String, String> options, Connection conn) throws ExchangeException {

        if (updates.isEmpty()) return;
System.out.println("Update " + itemName);

      //Convert the options parameter into xml attributes. These attributes
      //are inserted into the "UpdateItem" node.
        String attr = "";
        if (options!=null){
            java.util.Iterator<String> it = options.keySet().iterator();
            while (it.hasNext()){
                String key = it.next();
                attr += " " + key + "=\"" + options.get(key) + "\"";
            }
        }

        StringBuffer msg = new StringBuffer();
        msg.append("<?xml version=\"1.0\" encoding=\"utf-8\" ?>");
        msg.append("<soap:Envelope xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\" xmlns:m=\"http://schemas.microsoft.com/exchange/services/2006/messages\">");
        msg.append("<soap:Body>");
        msg.append("<m:UpdateItem" + attr + ">");
        msg.append("<m:ItemChanges>");
        msg.append("<t:ItemChange>");
        msg.append("<t:ItemId Id=\"" + id + "\" ChangeKey=\"" + getChangeKey(conn) + "\" />"); //
        msg.append("<t:Updates>");


        java.util.Iterator<String> it = updates.keySet().iterator();
        while (it.hasNext()){
            String key = it.next();
            Object value = updates.get(key);
            String namespace = getNameSpace(key, itemName);

            if (value==null){
                System.out.println("Delete " + key);
                msg.append("<t:DeleteItemField>");
                msg.append("<t:FieldURI FieldURI=\"" + namespace + ":" + key + "\"/>");
                msg.append("</t:DeleteItemField>");
            }
            else{
                System.out.println("Update " + key);

                String str = value.toString().trim();
                if (str.startsWith("<t:SetItemField") || str.startsWith("<t:DeleteItemField")){
                    msg.append(value);
                }
                else{
                    msg.append("<t:SetItemField>");
                    msg.append("<t:FieldURI FieldURI=\"" + namespace + ":" + key + "\" />");
                    msg.append("<t:" + itemName + ">");
                    msg.append("<t:" + key + ">" + value + "</t:" + key + ">");
                    msg.append("</t:" + itemName + ">");
                    msg.append("</t:SetItemField>");
                }

            }
        }

        msg.append("</t:Updates>");
        msg.append("</t:ItemChange>");
        msg.append("</m:ItemChanges>");
        msg.append("</m:UpdateItem>");
        msg.append("</soap:Body>");
        msg.append("</soap:Envelope>");

System.out.println(msg + "\r\n");

        updates.clear();

        conn.execute(msg.toString());
    }


  //**************************************************************************
  //** delete
  //**************************************************************************
  /** Used to delete this item.
   *  @param options Hashmap containing key/value pairs representing delete
   *  options. These options are inserted as attributes into the DeleteItem node.
   *  Example: DeleteType="MoveToDeletedItems" and SendMeetingCancellations="SendToAllAndSaveCopy"
   */
    protected void delete(java.util.HashMap<String, String> options, Connection conn) throws ExchangeException {

        if (options==null) options = new java.util.HashMap<String, String>();
        if (!options.containsKey("DeleteType")) options.put("DeleteType", "MoveToDeletedItems");
        String attr = "";
        java.util.Iterator<String> it = options.keySet().iterator();
        while (it.hasNext()){
            String key = it.next();
            attr += " " + key + "=\"" + options.get(key) + "\"";
        }


        String msg =
        "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
        + "<soap:Envelope xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\" xmlns:m=\"http://schemas.microsoft.com/exchange/services/2006/messages\">"
        //+ "<soap:Header><t:RequestServerVersion Version=\"Exchange2007_SP1\"/></soap:Header>"
        + "<soap:Body>"
        + "<m:DeleteItem" + attr + ">"
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
    protected static String formatDate(javaxt.utils.Date date){
        if (date==null) return null;
        String d = date.toString("yyyy-MM-dd HH:mm:ssZ").replace(" ", "T");
        return d.substring(0, d.length()-2) + ":" + d.substring(d.length()-2);
    }


  //**************************************************************************
  //** getNameSpace
  //**************************************************************************
  /** Returns the namespace associated with a given field. This is critical
   *  when updating items.
   */
    public static String getNameSpace(String fieldName, String itemName){

        java.util.HashSet<String> namespace = namespaces.get(fieldName);
        if (namespace!=null){
            if (namespace.size()==1) return namespace.iterator().next();
            else{
                //System.out.println("Multiple values for " + fieldName + ". What matches with " + itemName + "?");

                String[] possibleValues = new String[]{"item"};
                if (itemName.equals("Item")){
                    possibleValues = new String[]{"item"};
                }
                if (itemName.equals("Message")){
                    possibleValues = new String[]{"message", "item"};
                }
                if (itemName.equals("Contact")){
                    possibleValues = new String[]{"contacts", "item"};
                }
                if (itemName.equals("CalendarItem")){
                    possibleValues = new String[]{"calendar", "item"};
                }

               /** TODO: Need to come up with more mappings! Here's a list of
                 * all possible namespaces:
                   - message
                   - folder
                   - distributionlist
                   - postitem
                   - task
                   - meeting
                   - item
                   - conversation
                   - meetingRequest
                   - contacts
                   - calendar

                   And here's a list of every possible "item":

                   - Item
                   - Message
                   - CalendarItem
                   - Contact
                   - DistributionList
                   - MeetingMessage
                   - MeetingRequest
                   - MeetingResponse
                   - MeetingCancellation
                   - Task
                   - ReplyToItem
                   - ForwardItem
                   - ReplyAllToItem
                   - AcceptItem
                   - TentativelyAcceptItem
                   - DeclineItem
                   - CancelCalendarItem
                   - RemoveItem
                   - PostReplyItem
                   - SuppressReadReceipt
                   - AcceptSharingInvitation
                 */

                for (String v : possibleValues){
                    if (namespace.contains(v)){
                        java.util.Iterator<String> it = namespace.iterator();
                        while (it.hasNext()){
                            String val = it.next();
                            if (val.equals(v)) return val;
                        }
                    }
                }
            }
        }


        return null;
    }

    private static final java.util.HashMap<String, java.util.HashSet<String>> namespaces = getNameSpaces();
    private static final java.util.HashMap<String, java.util.HashSet<String>> getNameSpaces(){
        String fieldURIs = "calendar:AdjacentMeetingCount,calendar:AdjacentMeetings,calendar:AllowNewTimeProposal,calendar:AppointmentReplyTime,calendar:AppointmentSequenceNumber,calendar:AppointmentState,calendar:CalendarItemType,calendar:ConferenceType,calendar:ConflictingMeetingCount,calendar:ConflictingMeetings,calendar:DateTimeStamp,calendar:DeletedOccurrences,calendar:Duration,calendar:End,calendar:EndTimeZone,calendar:FirstOccurrence,calendar:IsAllDayEvent,calendar:IsCancelled,calendar:IsMeeting,calendar:IsOnlineMeeting,calendar:IsRecurring,calendar:IsResponseRequested,calendar:LastOccurrence,calendar:LegacyFreeBusyStatus,calendar:Location,calendar:MeetingRequestWasSent,calendar:MeetingTimeZone,calendar:MeetingWorkspaceUrl,calendar:ModifiedOccurrences,calendar:MyResponseType,calendar:NetShowUrl,calendar:OptionalAttendees,calendar:Organizer,calendar:OriginalStart,calendar:Recurrence,calendar:RecurrenceId,calendar:RequiredAttendees,calendar:Resources,calendar:Start,calendar:StartTimeZone,calendar:TimeZone,calendar:UID,calendar:When,contacts:Alias,contacts:AssistantName,contacts:Birthday,contacts:BusinessHomePage,contacts:Children,contacts:Companies,contacts:CompanyName,contacts:CompleteName,contacts:ContactSource,contacts:Culture,contacts:Department,contacts:DirectoryId,contacts:DirectReports,contacts:DisplayName,contacts:EmailAddresses,contacts:FileAs,contacts:FileAsMapping,contacts:Generation,contacts:GivenName,contacts:HasPicture,contacts:HasPicture,contacts:ImAddresses,contacts:Initials,contacts:JobTitle,contacts:Manager,contacts:ManagerMailbox,contacts:MiddleName,contacts:Mileage,contacts:MSExchangeCertificate,contacts:Nickname,contacts:Notes,contacts:OfficeLocation,contacts:PhoneNumbers,contacts:PhoneticFirstName,contacts:PhoneticFullName,contacts:PhoneticLastName,contacts:Photo,contacts:PhysicalAddresses,contacts:PostalAddressIndex,contacts:Profession,contacts:SpouseName,contacts:Surname,contacts:UserSMIMECertificate,contacts:WeddingAnniversary,conversation:Categories,conversation:ConversationId,conversation:ConversationTopic,conversation:FlagStatus,conversation:GlobalCategories,conversation:GlobalFlagStatus,conversation:GlobalHasAttachments,conversation:GlobalImportance,conversation:GlobalItemClasses,conversation:GlobalItemIds,conversation:GlobalLastDeliveryTime,conversation:GlobalMessageCount,conversation:GlobalSize,conversation:GlobalUniqueRecipients,conversation:GlobalUniqueSenders,conversation:GlobalUniqueUnreadSenders,conversation:GlobalUnreadCount,conversation:HasAttachments,conversation:Importance,conversation:ItemClasses,conversation:ItemIds,conversation:LastDeliveryTime,conversation:MessageCount,conversation:Size,conversation:UniqueRecipients,conversation:UniqueSenders,conversation:UniqueUnreadSenders,conversation:UnreadCount,distributionlist:Members,folder:ChildFolderCount,folder:DisplayName,folder:EffectiveRights,folder:FolderClass,folder:FolderId,folder:ManagedFolderInformation,folder:ParentFolderId,folder:PermissionSet,folder:SearchParameters,folder:SharingEffectiveRights,folder:TotalCount,folder:UnreadCount,item:Attachments,item:Body,item:Categories,item:ConversationId,item:Culture,item:DateTimeCreated,item:DateTimeReceived,item:DateTimeSent,item:DisplayCc,item:DisplayTo,item:EffectiveRights,item:HasAttachments,item:Importance,item:InReplyTo,item:InternetMessageHeaders,item:IsAssociated,item:IsDraft,item:IsFromMe,item:IsResend,item:IsSubmitted,item:IsUnmodified,item:ItemClass,item:ItemId,item:LastModifiedName,item:LastModifiedTime,item:MimeContent,item:ParentFolderId,item:ReminderDueBy,item:ReminderIsSet,item:ReminderMinutesBeforeStart,item:ResponseObjects,item:Sensitivity,item:Size,item:Subject,item:UniqueBody,item:WebClientEditFormQueryString,item:WebClientReadFormQueryString,meeting:AssociatedCalendarItemId,meeting:HasBeenProcessed,meeting:IsDelegated,meeting:IsOutOfDate,meeting:ResponseType,meetingRequest:IntendedFreeBusyStatus,meetingRequest:MeetingRequestType,message:BccRecipients,message:CcRecipients,message:ConversationIndex,message:ConversationTopic,message:From,message:InternetMessageId,message:IsDeliveryReceiptRequested,message:IsRead,message:IsReadReceiptRequested,message:IsResponseRequested,message:References,message:ReplyTo,message:Sender,message:ToRecipients,postitem:PostedTime,task:ActualWork,task:AssignedTime,task:BillingInformation,task:ChangeCount,task:Companies,task:CompleteDate,task:Contacts,task:DelegationState,task:Delegator,task:DueDate,task:IsAssignmentEditable,task:IsComplete,task:IsRecurring,task:IsTeamTask,task:Mileage,task:Owner,task:PercentComplete,task:Recurrence,task:StartDate,task:Status,task:StatusDescription,task:TotalWork";

        java.util.HashMap<String, java.util.HashSet<String>> namespaces = new java.util.HashMap<String, java.util.HashSet<String>>();
        java.util.HashSet<String> uniqueKeys = new java.util.HashSet<String>();
        java.util.HashSet<String> duplicateKeys = new java.util.HashSet<String>();


        for (String row : fieldURIs.split(",")){
            String[] col = row.split(":");
            if (col.length>1){
                String key = col[0];
                String value = col[1];

                java.util.HashSet<String> keys = namespaces.get(value);
                if (keys==null){
                    keys = new java.util.HashSet<String>();
                    namespaces.put(value, keys);
                }
                keys.add(key);

                if (uniqueKeys.contains(value)) duplicateKeys.add(value);
                else uniqueKeys.add(value);

                //System.out.println(col[0] + "\t" + col[1]);
            }
        }

        /*
        java.util.Iterator<String> it = duplicateKeys.iterator();
        while (it.hasNext()){
            String key = it.next();
            System.out.println("\r\n" + key);
            java.util.Iterator<String> i2 = namespaces.get(key).iterator();
            while (i2.hasNext()){
                String ns = i2.next();
                System.out.println(" - " + ns);
            }
        }
        */


        return namespaces;
    }
}