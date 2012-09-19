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
    private String importance = "Normal";
    private String sensitivity = "Normal";
    private Integer size;
    private boolean isRead = false;
    private boolean hasAttachments = false;

/*
<t:ItemClass>IPM.Note</t:ItemClass>
 <t:Subject>doc rec</t:Subject>
 <t:Sensitivity>Normal</t:Sensitivity>
 <t:Size>2279</t:Size>
 <t:DateTimeSent>2012-09-17T14:26:04Z</t:DateTimeSent>
 <t:DateTimeCreated>2012-09-17T14:30:15Z</t:DateTimeCreated>
 <t:HasAttachments>false</t:HasAttachments>
 <t:ExtendedProperty>
    <t:ExtendedFieldURI PropertyType="SystemTime" PropertyTag="0x3008"/>
    <t:Value>2012-09-17T14:30:15Z</t:Value></t:ExtendedProperty>
 <t:From>
   <t:Mailbox><t:Name>Ann_Liu</t:Name></t:Mailbox>
 </t:From>
 <t:IsRead>false</t:IsRead>
 */
    
  //**************************************************************************
  //** Constructor
  //**************************************************************************
  /** This constructor is provided for application developers who wish to
   *  extend this class.
   */
    protected Email(){}


  //**************************************************************************
  //** Constructor
  //**************************************************************************
  /** Creates a new instance of this class using an existing Email,
   *  effectively creating a clone.
   */
    public Email(javaxt.exchange.Email message){

      //General information
        this.id = message.id;
        this.subject = message.subject;
        this.body = message.body;
        this.bodyType = message.bodyType;
        this.categories = message.categories;
        this.updates = message.updates;
        this.lastModified = message.lastModified;
        this.extendedProperties = message.extendedProperties;

      //Email specific information
        this.from = message.from;
        this.importance = message.importance;
        this.sensitivity = message.sensitivity;
        this.size = message.size;
        this.isRead = message.isRead;
        this.hasAttachments = message.hasAttachments;
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
    public Email(String exchangeID, Connection conn, ExtendedProperty[] AdditionalProperties) throws ExchangeException{
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
                        from = new Mailbox(nodes[0]);
                    }
                }
                else if(nodeName.equalsIgnoreCase("IsRead")){
                    isRead = javaxt.xml.DOM.getNodeValue(outerNode).equalsIgnoreCase("true");
                }
                else if(nodeName.equalsIgnoreCase("HasAttachments")){
                    hasAttachments = javaxt.xml.DOM.getNodeValue(outerNode).equalsIgnoreCase("true");
                }
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
  //** getSender
  //**************************************************************************
  /** Returns the Mailbox associated with the sender.
   */
    public Mailbox getSender(){
        return from;
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
        if (this.isRead!=isRead){
            this.isRead = isRead;
            updates.put("IsRead", isRead);
        }
    }


  //**************************************************************************
  //** hasAttachments
  //**************************************************************************
  /** Returns a boolean used to indicate whether the message has attachments.
   */
    public boolean hasAttachments(){
        return hasAttachments;
    }
}