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
    private String domainAddress;

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
  //** getEmailAddress
  //**************************************************************************
  /** Returns the EmailAddress associated with this Mailbox, or null if the
   *  EmailAddress is undefined.
   */
    public EmailAddress getEmailAddress(){
        return EmailAddress;
    }


  //**************************************************************************
  //** setEmailAddress
  //**************************************************************************
  /** Used to set/update the EmailAddress associated with this Mailbox.
   */
    public void setEmailAddress(EmailAddress emailAddress){
        this.EmailAddress = emailAddress;
    }


  //**************************************************************************
  //** setEmailAddress
  //**************************************************************************
  /** Used to set/update the EmailAddress associated with this Mailbox.
   */
    public void setEmailAddress(String emailAddress) throws ExchangeException {
        setEmailAddress(new EmailAddress(emailAddress));
    }


  //**************************************************************************
  //** setDomainAddress
  //**************************************************************************
    protected void setDomainAddress(String domainAddress){
        this.domainAddress = domainAddress;
    }


  //**************************************************************************
  //** getDomainAddress
  //**************************************************************************
  /** Returns a string used to represent an email address of another Exchange
   *  account within your domain (e.g. "/o=org/ou=orgunit/cn=Recipients/cn=name").
   *  The domain address can be resolved to an email address using the
   *  resolveName() method.
   */
    public String getDomainAddress(){
        return domainAddress;
    }


  //**************************************************************************
  //** getName
  //**************************************************************************
  /** Returns the display name associated with this Mailbox (e.g. "John Smith").
   */
    public String getName(){
        return Name;
    }


  //**************************************************************************
  //** getID
  //**************************************************************************
    public String getID(){
        return ItemId;
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


  //**************************************************************************
  //** toString
  //**************************************************************************
  /** Returns a string representation of this Mailbox (e.g.
   *  "John Smith &lt;jsmith@acme.com&gt;").
   */
    public String toString(){
        if (Name!=null && EmailAddress!=null) return Name + " <" + EmailAddress + ">";
        if (EmailAddress!=null) return EmailAddress.toString();
        if (Name!=null) return Name;        
        return ItemId; //?
    }


  //**************************************************************************
  //** equals
  //**************************************************************************
  /** Used to compare this Mailbox to another. Returns true if the hashcodes
   *  match.
   */
    public boolean equals(Object obj){
        if (obj!=null){
            if (obj instanceof Mailbox){
                return ((Mailbox) obj).hashCode()==this.hashCode();
            }
        }
        return false;
    }


  //**************************************************************************
  //** hashCode
  //**************************************************************************
  /** Returns the hashcode associated with the EmailAddress. If the
   *  EmailAddress is undefined, returns the hashcode of the domain address.
   *  If both the EmailAddress and domain address are undefined, returns 0.
   */
    public int hashCode(){
        return (EmailAddress!=null) ? EmailAddress.hashCode() : 
            (domainAddress!=null? domainAddress.hashCode() : 0);
    }


  //**************************************************************************
  //** resolveName
  //**************************************************************************
  /** Attempts to resolve a user name, email address, or a domain address to a
   *  Mailbox.
   */
    public static Mailbox resolveName(String name, Connection conn) throws ExchangeException {
        StringBuffer str = new StringBuffer();

        str.append("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
        str.append("<soap:Envelope xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\" "
            + "xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\" "
            + "xmlns:m=\"http://schemas.microsoft.com/exchange/services/2006/messages\">");
        str.append("<soap:Body>");
        str.append("<m:ResolveNames ReturnFullContactData=\"false\" >");
        str.append("<m:UnresolvedEntry>");
        str.append(name);
        str.append("</m:UnresolvedEntry>");
        str.append("</m:ResolveNames>");
        str.append("</soap:Body>");
        str.append("</soap:Envelope>");

        org.w3c.dom.Document xml = conn.execute(str.toString());
        org.w3c.dom.Node[] items = javaxt.xml.DOM.getElementsByTagName("Resolution", xml);
        if (items.length>0){
            org.w3c.dom.NodeList nodes = items[0].getChildNodes();
            for (int i=0; i<nodes.getLength(); i++){
                org.w3c.dom.Node node = nodes.item(i);
                if (node.getNodeType()==1){
                    String nodeName = node.getNodeName();
                    if (nodeName.contains(":")) nodeName = nodeName.substring(nodeName.indexOf(":")+1);
                    if (nodeName.equalsIgnoreCase("Mailbox")){
                        return new Mailbox(node);
                    }
                    else if(nodeName.equalsIgnoreCase("Contact")){
                        //Contact contact = new Contact(node);
                    }
                }
            }

        }

      //If we're still here, throw an Exception
        throw new ExchangeException("No results were found.");
    }
}