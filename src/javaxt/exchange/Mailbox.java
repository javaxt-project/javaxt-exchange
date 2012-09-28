package javaxt.exchange;

//******************************************************************************
//**  Mailbox Class
//******************************************************************************
/**
 *   Used to represent a Mailbox element which is used to identify a mail-
 *   enabled Active Directory object.
 *   http://msdn.microsoft.com/en-us/library/aa565036%28v=EXCHG.140%29.aspx
 *
 ******************************************************************************/

public class Mailbox {

    private String Name;
    private EmailAddress EmailAddress;
    private String RoutingType;
    private String MailboxType;
    private String ItemId;


  //**************************************************************************
  //** Constructor
  //**************************************************************************
  /** Creates a new instance of Mailbox. */

    protected Mailbox(org.w3c.dom.Node mailboxNode) throws ExchangeException {
        org.w3c.dom.NodeList mailboxItems = mailboxNode.getChildNodes();

        for (int i=0; i<mailboxItems.getLength(); i++){
            org.w3c.dom.Node node = mailboxItems.item(i);
            if (node.getNodeType()==1){
                String nodeName = node.getNodeName();
                if (nodeName.contains(":")) nodeName =
                        nodeName.substring(nodeName.indexOf(":")+1);

                if (nodeName.equalsIgnoreCase("Name")){
                    Name = javaxt.xml.DOM.getNodeValue(node);
                }
                else if(nodeName.equalsIgnoreCase("EmailAddress")){
                    EmailAddress = new EmailAddress(javaxt.xml.DOM.getNodeValue(node));
                }
                else if(nodeName.equalsIgnoreCase("RoutingType")){
                    RoutingType = javaxt.xml.DOM.getNodeValue(node);
                }
                else if(nodeName.equalsIgnoreCase("MailboxType")){
                    MailboxType = javaxt.xml.DOM.getNodeValue(node);
                }
                else if(nodeName.equalsIgnoreCase("ItemId")){
                    ItemId = javaxt.xml.DOM.getNodeValue(node);
                }
            }
        }
    }

    public Mailbox(String name, EmailAddress email) {
        this.Name = name;
        this.EmailAddress = email;
    }

    public Mailbox(String name, String email) throws ExchangeException {
        this(name,  new EmailAddress(email));
    }

    public Mailbox(Contact contact){
        this.Name = contact.getFullName();
        this.EmailAddress = contact.getPrimaryEmailAddress();
    }


  //**************************************************************************
  //** toXML
  //**************************************************************************
  /** Returns an xml fragment used to save or update a mail item via Exchange
   *  Web Services (EWS):<br/>
   *  http://msdn.microsoft.com/en-us/library/aa563318%28v=exchg.140%29.aspx
   *
   *  @param namespace The namespace assigned to the "type". Typically this is
   *  "t" which corresponds to
   *  "http://schemas.microsoft.com/exchange/services/2006/types".
   *  Use a null value is you do not wish to append a namespace.
   */
    protected String toXML(String namespace){

      //Update namespace prefix
        if (namespace!=null){
            if (!namespace.endsWith(":")) namespace+=":";
        }
        else{
            namespace = "";
        }

        StringBuffer str = new StringBuffer();

        str.append("<" + namespace + "Mailbox>");
        if (Name!=null) str.append("<" + namespace + "Name>" + Name + "</" + namespace + "Name>");
        if (EmailAddress!=null)  str.append("<" + namespace + "EmailAddress>" + EmailAddress + "</" + namespace + "EmailAddress>");
        if (RoutingType!=null)  str.append("<" + namespace + "RoutingType>" + RoutingType + "</" + namespace + "RoutingType>");
        if (MailboxType!=null)  str.append("<" + namespace + "MailboxType>" + MailboxType + "</" + namespace + "MailboxType>");
        if (ItemId!=null)  str.append("<" + namespace + "ItemId>" + ItemId + "</" + namespace + "ItemId>");
        str.append("</" + namespace + "Mailbox>");

        return str.toString().trim();
    }


    public EmailAddress getEmailAddress(){
        return EmailAddress;
    }


    public String toString(){
        if (Name!=null && EmailAddress!=null) return Name + " <" + EmailAddress + ">";
        if (EmailAddress!=null) return EmailAddress.toString();
        if (Name!=null) return Name;        
        return ItemId; //?
    }


    public boolean equals(Object obj){
        if (obj instanceof Mailbox){
            return ((Mailbox) obj).hashCode()==this.hashCode();
        }
        return false;
    }

    public int hashCode(){
        return EmailAddress.hashCode();
    }

}