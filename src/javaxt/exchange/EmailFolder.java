package javaxt.exchange;

public class EmailFolder extends Folder {

    private Connection conn;
    private java.util.ArrayList<String> props = new java.util.ArrayList<String>();


  //**************************************************************************
  //** Constructor
  //**************************************************************************
  /** Creates a new instance of this class.
   *  @param folderName Distinguished name or unique ID of the folder (e.g. 
   *  "inbox", "drafts", "deleteditems", "junkemail", "outbox", etc.).
   */
    public EmailFolder(String folderName, Connection conn) throws ExchangeException {
        super(folderName, conn);
        props.add("item:Importance");
        this.conn = conn;
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
  //** getMail
  //**************************************************************************
  /** Returns a shallow representation of email messages found in this folder.
   *  @param limit Maximum number of items to return.
   *  @param offset Item offset. 0 implies no offset.
   *  @param orderBy. SQL-style order by clause (e.g. "item:DateTimeReceived DESC")
   */
    public Email[] getMail(int offset, int limit, String orderBy) throws ExchangeException {
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


  //**************************************************************************
  //** getMail
  //**************************************************************************
  /** Returns an email message associated with the given exchangeID
   */
    public Email getMail(String exchangeID) throws ExchangeException {
        return new Email(exchangeID, conn);
    }
}