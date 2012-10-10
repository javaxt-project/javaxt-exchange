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
    public EmailAddress getEmailAddress(){
        return EmailAddress;
    }

    public void setEmailAddress(EmailAddress emailAddress){
        this.EmailAddress = emailAddress;
    }

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
    public String getName(){
        return Name;
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




    public String toString(){
        if (Name!=null && EmailAddress!=null) return Name + " <" + EmailAddress + ">";
        if (EmailAddress!=null) return EmailAddress.toString();
        if (Name!=null) return Name;        
        return ItemId; //?
    }


    public boolean equals(Object obj){
        if (obj!=null){
            if (obj instanceof Mailbox){
                return ((Mailbox) obj).hashCode()==this.hashCode();
            }
        }
        return false;
    }

    public int hashCode(){
        return (EmailAddress != null) ? EmailAddress.hashCode() : 0;
    }


  /** Attempts to resolve a user name, email address, or an ADSI string to a
   *  Mailbox and Contact.
   *  @return
   */
    public static Mailbox resolveName(String name, Connection conn) throws ExchangeException {
        StringBuffer str = new StringBuffer();

        str.append("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
        str.append("<soap:Envelope xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\" "
            + "xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\" "
            + "xmlns:m=\"http://schemas.microsoft.com/exchange/services/2006/messages\">");
        //+ "<soap:Header><t:RequestServerVersion Version=\"Exchange2007\"/></soap:Header>"
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
                    //System.out.println(node.getNodeName());
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