package javaxt.exchange;

//******************************************************************************
//**  EmailFolder Class
//******************************************************************************
/**
 *   Used to represent an mail folder (e.g. "Inbox", "Sent Items", etc.).
 *
 ******************************************************************************/

public class EmailFolder extends Folder {

    private java.util.HashSet<FieldURI> props = new java.util.HashSet<FieldURI>();


  //**************************************************************************
  //** Constructor
  //**************************************************************************
  /** Creates a new instance of this class.
   *  @param folderName Distinguished name or unique ID of the folder (e.g. 
   *  "inbox", "drafts", "deleteditems", "junkemail", "outbox", etc.).
   */
    public EmailFolder(String folderName, Connection conn) throws ExchangeException {
        super(folderName, conn);
        props.add(new FieldURI("item:Importance"));
        props.add(new FieldURI("item:DateTimeReceived"));
        props.add(new FieldURI("item:IsDraft"));


      //Add the PR_LAST_VERB_EXECUTED (0x10810003) extended MAPI property. This
      //property represents the last action performed on individual mail items.
      //Possible values include 102 (replied), 103 (reply all), and 104 (forwarded).
        props.add(new ExtendedFieldURI(null, "0x1081", "Integer"));
        
      //Add the PR_LAST_VERB_EXECUTION_TIME (0x10820040) extended MAPI property.
      //This property represents the date/time associated with the last action
      //performed on individual mail items.
        props.add(new ExtendedFieldURI(null, "0x1082", "SystemTime"));
    }


  //**************************************************************************
  //** getMailFolders
  //**************************************************************************
  /** Returns an array of folders that contain, or are associated with, email
   *  messages (e.g. "Inbox", "Deleted Items", "Drafts", "Junk E-mail", 
   *  "Outbox", "Sent Items", etc). This is a convenience method and is 
   *  identical to calling:
   * <pre>new javaxt.exchange.Folder("msgfolderroot", conn).getFolders();</pre>
   */
    public static Folder[] getMailFolders(Connection conn) throws ExchangeException {
        return new javaxt.exchange.Folder("msgfolderroot", conn).getFolders();
    }


  //**************************************************************************
  //** getMessages
  //**************************************************************************
  /** Returns a shallow representation of email messages found in this folder.
   *
   *  @param offset Item offset. 0 implies no offset.
   *
   *  @param limit Maximum number of items to return.
   *
   *  @param additionalProperties By default, this method returns a shallow
   *  representation messages found in this folder. You can retrieve additional
   *  attributes by providing an array of FieldURI (including ExtendedFieldURIs).
   *  This parameter is optional. A null value will return no additional or
   *  extended attributes.
   *
   *  @param orderBy. SQL-style order by clause used to sort the response 
   *  (e.g. "item:DateTimeReceived DESC"). This parameter is optional. A null
   *  value implies no sort preference.
   */
    public Email[] getMessages(int offset, int limit, FieldURI[] additionalProperties, String orderBy) throws ExchangeException {

      //Merge additional properties
        java.util.HashSet<FieldURI> props = getDefaultProperties();
        if (additionalProperties!=null){
            for (FieldURI property : additionalProperties) props.add(property);
        }

        java.util.ArrayList<Email> messages = new java.util.ArrayList<Email>();
        org.w3c.dom.NodeList nodes = getItems(offset, limit, props, orderBy).getElementsByTagName("t:Message");
        for (int i=0; i<nodes.getLength(); i++){
            org.w3c.dom.Node node = nodes.item(i);
            if (node.getNodeType()==1){
                messages.add(new Email(node));
            }
        }
        return messages.toArray(new Email[messages.size()]);
    }


    private java.util.HashSet<FieldURI> getDefaultProperties(){
        java.util.HashSet<FieldURI> map = new java.util.HashSet<FieldURI>();
        map.addAll(props);
        return map;
    }
}