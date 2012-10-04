package javaxt.exchange;

//******************************************************************************
//**  Email Message
//******************************************************************************
/**
 *   Used to represent a mail message found in a Mail Folder (e.g. "Inbox").
 *   http://msdn.microsoft.com/en-us/library/aa494306%28v=exchg.140%29.aspx
 *
 ******************************************************************************/

public class Email extends FolderItem {

    private Mailbox from;
    private java.util.HashSet<Mailbox> ToRecipients = new java.util.HashSet<Mailbox>();
    private java.util.HashSet<Mailbox> CcRecipients = new java.util.HashSet<Mailbox>();
    private java.util.HashSet<Mailbox> BccRecipients = new java.util.HashSet<Mailbox>();

    private String importance = "Normal";
    private String sensitivity = "Normal";
    private Integer size;
    private boolean isRead = false;
    private javaxt.utils.Date sent;
    private javaxt.utils.Date received;
    private String response;
    private javaxt.utils.Date responseDate;

  //The following parameters are for internal use only!
    private String referenceId;
    private String messageType = "Message";


  //**************************************************************************
  //** Constructor
  //**************************************************************************
  /** Used to create a new email message. The message will be saved in the
   *  "Drafts" folder.
   */
    public Email(){}


  //**************************************************************************
  //** Constructor
  //**************************************************************************
  /** Creates a new instance of this class using an existing Email,
   *  effectively creating a clone.
   */
    public Email(javaxt.exchange.Email message){
        this.init(message);
    }

    private void init(javaxt.exchange.Email message){

      //General information
        this.id = message.id;
        this.parentFolderID = message.parentFolderID;
        this.subject = message.subject;
        this.body = message.body;
        this.bodyType = message.bodyType;
        this.categories = message.categories;
        this.hasAttachments = message.hasAttachments;
        this.attachments = message.attachments;
        this.updates = message.updates;
        this.lastModified = message.lastModified;
        this.extendedProperties = message.extendedProperties;

      //Email specific information
        this.from = message.from;
        this.ToRecipients = message.ToRecipients;
        this.CcRecipients = message.CcRecipients;
        this.BccRecipients = message.BccRecipients;
        this.importance = message.importance;
        this.sensitivity = message.sensitivity;
        this.size = message.size;
        this.isRead = message.isRead;
        this.sent = message.sent;
        this.received = message.received;
        this.response = message.response;
        this.responseDate = message.responseDate;
        this.referenceId = message.referenceId;
        this.messageType = message.messageType;
    }


  //**************************************************************************
  //** Constructor
  //**************************************************************************
  /** Creates a new instance of Contact using a node from a
   *  FindItemResponseMessage.
   */
    protected Email(org.w3c.dom.Node messageNode) {
        super(messageNode);
        parseMessage();
    }


  //**************************************************************************
  //** Constructor
  //**************************************************************************
  /** Creates a new instance of this class
   */
    public Email(String exchangeID, Connection conn, ExtendedFieldURI[] AdditionalProperties) throws ExchangeException{
        super(exchangeID, conn, AdditionalProperties);
        parseMessage();
    }


  //**************************************************************************
  //** Constructor
  //**************************************************************************
  /** Creates a new instance of this class
   */
    public Email(String exchangeID, Connection conn) throws ExchangeException{
        this(exchangeID, conn, null);
    }


  //**************************************************************************
  //** parseMessage
  //**************************************************************************
  /** Used to parse an xml node with email information.
   */
    private void parseMessage(){

        boolean isDraft = false;

        org.w3c.dom.NodeList outerNodes = this.getChildNodes();
        for (int i=0; i<outerNodes.getLength(); i++){
            org.w3c.dom.Node outerNode = outerNodes.item(i);
            if (outerNode.getNodeType()==1){
                String nodeName = outerNode.getNodeName();
                if (nodeName.contains(":")) nodeName = nodeName.substring(nodeName.indexOf(":")+1);
                if (nodeName.equalsIgnoreCase("Sensitivity")){
                    sensitivity = javaxt.xml.DOM.getNodeValue(outerNode);
                }
                else if(nodeName.equalsIgnoreCase("Importance")){
                    importance = javaxt.xml.DOM.getNodeValue(outerNode);
                }
                else if(nodeName.equalsIgnoreCase("Size")){
                    try{
                        size = Integer.parseInt(javaxt.xml.DOM.getNodeValue(outerNode));
                    }
                    catch(Exception e){
                    }
                }
                else if(nodeName.equalsIgnoreCase("From")){
                    org.w3c.dom.Node[] nodes = javaxt.xml.DOM.getElementsByTagName("Mailbox", outerNode);
                    if (nodes.length>0){
                        try{
                            from = new Mailbox(nodes[0]);
                        }
                        catch(ExchangeException e){
                        }
                    }
                }
                else if(nodeName.equalsIgnoreCase("ToRecipients")){
                    for (org.w3c.dom.Node node : javaxt.xml.DOM.getElementsByTagName("Mailbox", outerNode)){
                        try{
                            ToRecipients.add(new Mailbox(node));
                        }
                        catch(ExchangeException e){
                        }
                    }
                }
                else if(nodeName.equalsIgnoreCase("CcRecipients")){
                    for (org.w3c.dom.Node node : javaxt.xml.DOM.getElementsByTagName("Mailbox", outerNode)){
                        try{
                            CcRecipients.add(new Mailbox(node));
                        }
                        catch(ExchangeException e){
                        }
                    }
                }
                else if(nodeName.equalsIgnoreCase("BccRecipients")){
                    for (org.w3c.dom.Node node : javaxt.xml.DOM.getElementsByTagName("Mailbox", outerNode)){
                        try{
                            BccRecipients.add(new Mailbox(node));
                        }
                        catch(ExchangeException e){
                        }
                    }
                }
                else if(nodeName.equalsIgnoreCase("IsRead")){
                    isRead = javaxt.xml.DOM.getNodeValue(outerNode).equalsIgnoreCase("true");
                }
                else if(nodeName.equalsIgnoreCase("IsDraft")){
                    isDraft = javaxt.xml.DOM.getNodeValue(outerNode).equalsIgnoreCase("true");
                }
                else if(nodeName.equalsIgnoreCase("DateTimeSent")){
                    try{
                        sent = new javaxt.utils.Date(javaxt.xml.DOM.getNodeValue(outerNode));
                    }
                    catch(java.text.ParseException e){}
                }
                else if(nodeName.equalsIgnoreCase("DateTimeReceived")){
                    try{
                        received = new javaxt.utils.Date(javaxt.xml.DOM.getNodeValue(outerNode));
                    }
                    catch(java.text.ParseException e){}
                }
            }
        }

      //Update the sent and received timestamps as needed
        if (isDraft) sent = received = null;


      //Find the PR_LAST_VERB_EXECUTED (0x10810003) extended MAPI property
        java.util.Iterator<ExtendedFieldURI> it = extendedProperties.keySet().iterator();
        while(it.hasNext()){
            ExtendedFieldURI property = it.next();
            if (property.getName().equalsIgnoreCase("0x1081")){
                Integer value = extendedProperties.get(property).toInteger();
                if (value==102) this.response = "Reply";
                else if(value == 103) this.response = "ReplyAll";
                else if(value == 104) this.response = "Forward";
                extendedProperties.remove(property);
                break;
            }
        }

        
      //Find the PR_LAST_VERB_EXECUTION_TIME (0x10820040) extended MAPI property
        it = extendedProperties.keySet().iterator();
        while(it.hasNext()){
            ExtendedFieldURI property = it.next();
            if (property.getName().equalsIgnoreCase("0x1082")){
                responseDate = extendedProperties.get(property).toDate();
                extendedProperties.remove(property);
                break;
            }
        }
    }



  //**************************************************************************
  //** getSubject
  //**************************************************************************
  /** Returns the subject associated with this message.
   */
    public String getSubject(){
        return super.getSubject();
    }


  //**************************************************************************
  //** setSubject
  //**************************************************************************
  /** Used to set/update the subject.
   */
    public void setSubject(String subject){
        super.setSubject(subject);
    }


  //**************************************************************************
  //** getBody
  //**************************************************************************
  /** Used to get the content of the message.
   */
    public String getBody(){
        return super.getBody();
    }


  //**************************************************************************
  //** setBody
  //**************************************************************************
  /** Used to set the content of the message.
   *  @param format Text format ("Best", "HTML", or "Text").
   */
    public void setBody(String description, String format){
        super.setBody(description, format);
    }


  //**************************************************************************
  //** getBodyType
  //**************************************************************************
  /** Returns the text encoding used in the body of this item. Possible values
   *  include "Best", "HTML", or "Text".
   */
    public String getBodyType(){
        return super.getBodyType();
    }


  //**************************************************************************
  //** getImportance
  //**************************************************************************
  /** Returns the importance assigned to this message. Possible values include
   *  "Low", "Normal", and "High".
   */
    public String getImportance(){
        return importance;
    }


  //**************************************************************************
  //** setImportance
  //**************************************************************************
  /** Used to assign an importance to this message.
   *  @param importance Possible values include "Low", "Normal", and "High"
   */
    public void setImportance(String importance){
        if (importance!=null) importance = importance.trim();
        if (importance==null || importance.length()==0) importance = "Normal";

        if (importance.equalsIgnoreCase("High")) importance = "High";
        else if (importance.equalsIgnoreCase("Low")) importance = "Low";
        else importance = "Normal";

        if (!this.importance.equals(importance)){
            this.importance = importance;
            updates.put("Importance", importance);
        }
    }


  //**************************************************************************
  //** getSensitivity
  //**************************************************************************
  /** Returns the sensitivity level associated with this message. Possible
   *  values include "Normal", "Personal", "Private", and "Confidential".
   */
    public String getSensitivity(){
        return sensitivity;
    }


  //**************************************************************************
  //** setSensitivity
  //**************************************************************************
  /** Used to set the sensitivity level associated with this message.
   *  @param sensitivity Possible values include "Normal", "Personal",
   *  "Private", and "Confidential".
   */
    public void setSensitivity(String sensitivity){
        if (sensitivity!=null) sensitivity = sensitivity.trim();
        if (sensitivity==null || sensitivity.length()==0) sensitivity = "Normal";

        if (sensitivity.equalsIgnoreCase("Personal")) sensitivity = "Personal";
        else if (sensitivity.equalsIgnoreCase("Private")) sensitivity = "Private";
        else if (sensitivity.equalsIgnoreCase("Confidential")) sensitivity = "Confidential";
        else sensitivity = "Normal";

        if (!this.sensitivity.equals(sensitivity)){
            this.sensitivity = sensitivity;
            updates.put("Sensitivity", sensitivity);
        }
    }


  //**************************************************************************
  //** getFrom
  //**************************************************************************
  /** Returns the Mailbox associated with the sender.
   */
    public Mailbox getFrom(){
        return from;
    }


  //**************************************************************************
  //** getDateTimeReceived
  //**************************************************************************
  /** Returns the date/time when the message was received. Returns a null if
   *  the message is a draft.
   */
    public javaxt.utils.Date getDateTimeReceived(){
        return received;
    }


  //**************************************************************************
  //** getDateTimeSent
  //**************************************************************************
  /** Returns the date/time when the message was sent. Returns a null if
   *  the message has not been sent (e.g. Draft message).
   */
    public javaxt.utils.Date getDateTimeSent(){
        return sent;
    }


  //**************************************************************************
  //** getSize
  //**************************************************************************
  /** Returns the size of the message.
   */
    public Integer getSize(){
        return size;
    }


  //**************************************************************************
  //** getResponse
  //**************************************************************************
  /** Returns the last action performed on this message. Possible values
   *  include "Forward", "Reply", "ReplyAll", or null.
   */
    public String getResponse(){
        return response;
    }


  //**************************************************************************
  //** getResponseDate
  //**************************************************************************
  /** Returns the date/time associated with the last action performed on this
   *  message. This, in conjunction with the getResponse() method can be used
   *  to generate messages like "You replied on 12/13/2012 5:38 PM." or
   *  "You forwarded this message on 12/13/2012 6:01 PM."
   */
    public javaxt.utils.Date getResponseDate(){
        return responseDate;
    }


  //**************************************************************************
  //** isRead
  //**************************************************************************
  /** Returns a boolean used to indicate whether the message has been read.
   */
    public boolean isRead(){
        return isRead;
    }


  //**************************************************************************
  //** setIsRead
  //**************************************************************************
  /** Returns a boolean used to indicate whether the message has been read.
   */
    public void setIsRead(boolean isRead){
        if (id!=null) {
            if (this.isRead!=isRead){
                this.isRead = isRead;
                updates.put("IsRead", isRead);
            }
        }
    }


  //**************************************************************************
  //** addRecipient
  //**************************************************************************
  /** Used to add a recipient to this message.
   *  @param list Possible values include "To", "Cc", or "Bcc".
   */
    public void addRecipient(String list, Mailbox recipient){
        updateRecipients("add", recipient, list);
    }


  //**************************************************************************
  //** removeRecipient
  //**************************************************************************
  /** Used to remove a recipient from this message.
   *  @param list Possible values include "To", "Cc", or "Bcc".
   */
    public void removeRecipient(String list, Mailbox recipient){
        updateRecipients("remove", recipient, list);
    }


  //**************************************************************************
  //** updateRecipients
  //**************************************************************************
  /** Private method used to add/remove a recipient from this message.
   */
    private void updateRecipients(String action, Mailbox recipient, String type){
        String updateNode = "";
        java.util.HashSet<Mailbox> recipients = null;
        if (type.equalsIgnoreCase("To")){
            recipients = ToRecipients;
            updateNode = "ToRecipients";
        }
        else if(type.equalsIgnoreCase("Cc")){
            recipients = CcRecipients;
            updateNode = "CcRecipients";
        }
        else if(type.equalsIgnoreCase("Bcc")){
            recipients = BccRecipients;
            updateNode = "BccRecipients";
        }
        else return; //Throw Exception?

        if (action.equals("add")){
            if (recipients.contains(recipient)) return;
            else recipients.add(recipient);
        }
        else if (action.equals("remove")){
            if (!recipients.contains(recipient)) return;
            else recipients.remove(recipient);
        }
        else{
            return;
        }

        if (id!=null) {
            StringBuffer xml = new StringBuffer();
            for (Mailbox r : recipients){
                xml.append(r.toXML("t"));
            }
            updates.put(updateNode, xml.toString());
        }
    }


  //**************************************************************************
  //** getToRecipients
  //**************************************************************************
  /** Returns an array of all the recipients on the "To" list.
   */
    public Mailbox[] getToRecipients(){
        if (ToRecipients.isEmpty()) return null;
        else return ToRecipients.toArray(new Mailbox[ToRecipients.size()]);
    }


  //**************************************************************************
  //** getCcRecipients
  //**************************************************************************
  /** Returns an array of all the recipients on the "CC" list.
   */
    public Mailbox[] getCcRecipients(){
        if (CcRecipients.isEmpty()) return null;
        else return CcRecipients.toArray(new Mailbox[CcRecipients.size()]);
    }


  //**************************************************************************
  //** getBccRecipients
  //**************************************************************************
  /** Returns an array of all the recipients on the "BCC" list.
   */
    public Mailbox[] getBccRecipients(){
        if (BccRecipients.isEmpty()) return null;
        else return BccRecipients.toArray(new Mailbox[BccRecipients.size()]);
    }


  //**************************************************************************
  //** delete
  //**************************************************************************
  /** Used to delete a message.
   *  @param MoveToDeletedItems If true, moves the item to the deleted items
   *  folder. If false, permanently deletes the message.
   */
    public void delete(boolean MoveToDeletedItems, Connection conn) throws ExchangeException {
        java.util.HashMap<String, String> options = new java.util.HashMap<String, String>();
        if (MoveToDeletedItems) options.put("DeleteType", "MoveToDeletedItems");
        else options.put("DeleteType", "HardDelete");
        super.delete(options, conn);
    }


  //**************************************************************************
  //** forward
  //**************************************************************************
  /** Creates a forwarded message and saves it to the drafts folder. The
   *  message won't be sent until the send() method is called.
   */
    public Email forward(Connection conn) throws ExchangeException {
        if (getID()==null) throw new ExchangeException("Can't forward message.");
        Email email = new Email();
        email.messageType = "ForwardItem";
        email.referenceId = this.getID();
        email.save(conn);
        return new Email(email.getID(), conn);
    }


  //**************************************************************************
  //** reply
  //**************************************************************************
  /** Creates a reply message and saves it to the drafts folder. The message
   *  won't be sent until the send() method is called.
   */
    public Email reply(Connection conn) throws ExchangeException {
        if (getID()==null) throw new ExchangeException("Can't reply to message.");
        Email email = new Email();
        email.messageType = "ReplyToItem";
        email.referenceId = this.getID();
        email.save(conn);
        return new Email(email.getID(), conn);
    }


  //**************************************************************************
  //** replyAll
  //**************************************************************************
  /** Creates a replyAll message and saves it to the drafts folder. The
   *  message won't be sent until the send() method is called.
   */
    public Email replyAll(Connection conn) throws ExchangeException {
        if (getID()==null) throw new ExchangeException("Can't reply to message.");
        Email email = new Email();
        email.messageType = "ReplyAllToItem";
        email.referenceId = this.getID();
        email.save(conn);
        return new Email(email.getID(), conn);
    }


  //**************************************************************************
  //** save
  //**************************************************************************
  /** Used to save this message. */

    public void save(Connection conn) throws ExchangeException {

      //Save the email message
        if (this.getID()==null) this.create(conn);
        else{
            java.util.HashMap<String, String> options = new java.util.HashMap<String, String>();
            options.put("ConflictResolution", "AutoResolve");
            options.put("MessageDisposition", "SaveOnly");
            super.update("Message", options, conn);
        }
        
      //Save attachments
        Attachment[] attachments = this.getAttachments();
        if (attachments!=null){
            for (Attachment attachment : attachments){
                if (attachment.getID()==null) attachment.save(conn);
            }
        }

      //Reset all the attributes of this item to reflect what's in Exchange
        init(new Email(id, conn));
    }


  //**************************************************************************
  //** send
  //**************************************************************************
  /** Used to send this message. A copy of this message will be saves in the
   *  "Sent" folder.
   */
    public void send(Connection conn) throws ExchangeException {

      //Make sure there's at least one recipient before sending the email
        if (this.getToRecipients()==null && this.getCcRecipients()==null &&
            this.getBccRecipients()==null){
            throw new ExchangeException("At least one recipient is required.");
        }

      //Save the message
        save(conn);

      //Send the message and create a copy in the "sentitems" folder
        StringBuffer msg = new StringBuffer();
        msg.append("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
        msg.append("<soap:Envelope xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\" xmlns:m=\"http://schemas.microsoft.com/exchange/services/2006/messages\">");
        msg.append("<soap:Body>");
        msg.append("<m:SendItem SaveItemToFolder=\"" + true + "\">");
        msg.append("<m:ItemIds>");
        msg.append("<t:ItemId Id=\"" + this.getID() + "\" ChangeKey=\"" + this.getChangeKey(conn) + "\" />"); //
        msg.append("</m:ItemIds>");
        msg.append("<m:SavedItemFolderId>");
        msg.append("<t:DistinguishedFolderId Id=\"" + "sentitems" + "\" />");
        msg.append("</m:SavedItemFolderId>");
        msg.append("</m:SendItem>");
        msg.append("</soap:Body>");
        msg.append("</soap:Envelope>");
        conn.execute(msg.toString());
    }


  //**************************************************************************
  //** create
  //**************************************************************************
  /** Used to create a new email message.
   */
    private void create(Connection conn) throws ExchangeException {

        String action = "SaveOnly";

        StringBuffer msg = new StringBuffer();
        msg.append("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
        msg.append("<soap:Envelope xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\" xmlns:m=\"http://schemas.microsoft.com/exchange/services/2006/messages\">");
        msg.append("<soap:Body>");
        msg.append("<m:CreateItem MessageDisposition=\"" + action + "\">");

        msg.append("<m:SavedItemFolderId>");
        if (this.getParentFolderID()!=null){
            msg.append("<t:FolderId Id=\"" + this.getParentFolderID() + "\"/>");
        }
        else{
            String folderName = "drafts";
            if (action.equals("SendOnly") || action.equals("SendAndSaveCopy")) folderName = "sentitems";
            msg.append("<t:DistinguishedFolderId Id=\"" + folderName + "\" />");
        }
        msg.append("</m:SavedItemFolderId>");
        msg.append("<m:Items>");

      /*
        Here's is an ordered list of all email properties. 

        WARNING -- ORDER IS VERY IMPORTANT!!! If you  mess up the order of the
        properties, the operation will fail - at least it did on my Exchange
        Server 2007 SP3 (8.3)

        MimeContent, ItemId, ParentFolderId, ItemClass, Subject, Sensitivity,
        Body, Attachments, DateTimeReceived, Size, Categories, Importance,
        InReplyTo, IsSubmitted, IsDraft, IsFromMe, IsResend, IsUnmodified,
        InternetMessageHeaders, DateTimeSent, DateTimeCreated, ResponseObjects,
        ReminderDueBy, ReminderIsSet, ReminderMinutesBeforeStart, DisplayCc,
        DisplayTo, HasAttachments, ExtendedProperty, Culture, Sender,
        ToRecipients, CcRecipients, BccRecipients, IsReadReceiptRequested,
        IsDeliveryReceiptRequested, ConversationIndex, ConversationTopic, From,
        InternetMessageId, IsRead, IsResponseRequested, References, ReplyTo,
        EffectiveRights, ReceivedBy, ReceivedRepresenting, LastModifiedName,
        LastModifiedTime, IsAssociated, WebClientReadFormQueryString,
        WebClientEditFormQueryString, ConversationId, UniqueBody
      */


        
        msg.append("<t:" + messageType + ">");
        if (this.getSubject()!=null) msg.append("<t:Subject>" + this.getSubject() + "</t:Subject>");
        if (referenceId==null) msg.append("<t:Sensitivity>" + this.getSensitivity() + "</t:Sensitivity>");

      //Set body
        if (getBody()!=null){
            msg.append("<t:Body BodyType=\"" + getBodyType() + "\">");
            msg.append(wrap(body));
            msg.append("</t:Body>");
        };


      //Set properties for new email messages
        if (referenceId==null){
            msg.append("<t:Importance>" + getImportance() + "</t:Importance>");
            if (from!=null) msg.append("<t:Sender>" + from.toXML("t") + "</t:Sender>"); //<--Doesn't seem to work!
        }
        
        Mailbox[] ToRecipients = this.getToRecipients();
        if (ToRecipients!=null){
            msg.append("<t:ToRecipients>");
            for (Mailbox recipient : ToRecipients){
                msg.append(recipient.toXML("t"));
            }
            msg.append("</t:ToRecipients>");
        }

        Mailbox[] CcRecipients = this.getCcRecipients();
        if (CcRecipients!=null){
            msg.append("<t:CcRecipients>");
            for (Mailbox recipient : CcRecipients){
                msg.append(recipient.toXML("t"));
            }
            msg.append("</t:CcRecipients>");
        }

        Mailbox[] BccRecipients = this.getBccRecipients();
        if (BccRecipients!=null){
            msg.append("<t:BccRecipients>");
            for (Mailbox recipient : BccRecipients){
                msg.append(recipient.toXML("t"));
            }
            msg.append("</t:BccRecipients>");
        }

        if (referenceId!=null){
            /*
               <IsReadReceiptRequested/>
               <IsDeliveryReceiptRequested/>
               <From/>
            */


            msg.append("<t:ReferenceItemId Id=\"" + referenceId + "\" ChangeKey=\"" + new Email(referenceId, conn).getChangeKey() + "\" />");
            /*
               <NewBodyContent/>
               <ReceivedBy/>
               <ReceivedRepresenting/>
             */
        }

        if (referenceId==null){
            if (from!=null) msg.append("<t:From>" + from.toXML("t") + "</t:From>"); //<--Doesn't seem to work!
        }
        
        msg.append("</t:" + messageType + ">");


        msg.append("</m:Items>");
        msg.append("</m:CreateItem>");
        msg.append("</soap:Body>");
        msg.append("</soap:Envelope>");



        org.w3c.dom.Document xml = conn.execute(msg.toString());

      //Parse the response. Note that send events don't return an Item ID
        org.w3c.dom.NodeList nodes = xml.getElementsByTagName("t:ItemId");
        if (nodes!=null && nodes.getLength()>0){
            id = javaxt.xml.DOM.getAttributeValue(nodes.item(0), "Id");
        }
    }

    public String toString(){
        return this.getSubject();
    }
}