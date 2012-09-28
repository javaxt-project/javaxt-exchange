package javaxt.exchange;

//******************************************************************************
//**  ContactFolder Class
//******************************************************************************
/**
 *   Used to represent a contact folder.
 *
 ******************************************************************************/

public class ContactsFolder extends Folder {

    private Connection conn;


  //**************************************************************************
  //** Constructor
  //**************************************************************************
  /** Creates a new instance of this class. */

    public ContactsFolder(Connection conn) throws ExchangeException {
        super("contacts", conn);
        this.conn = conn;
    }


  //**************************************************************************
  //** getContacts
  //**************************************************************************
  /** Returns an array of all contacts found in the contact folder.
   *
    public Contact[] getContacts() throws ExchangeException {

        java.util.ArrayList<Contact> contacts = new java.util.ArrayList<Contact>();
        
        int offset = 0;
        int maxRecords = 25;
        while(true){

            Contact[] arr = getContacts(offset, maxRecords);
            for (int i=0; i<arr.length; i++){
                contacts.add(arr[i]);
            }
            if (arr.length<maxRecords) break;
            else offset+=maxRecords;
        }

        return contacts.toArray(new Contact[contacts.size()]);
    }
*/

  //**************************************************************************
  //** getContacts
  //**************************************************************************
  /** Returns an array of contacts.
   *  @param limit Maximum number of items to return.
   *  @param offset Item offset. 0 implies no offset.
   */
    public Contact[] getContacts(int offset, int limit) throws ExchangeException {
        java.util.ArrayList<Contact> contacts = new java.util.ArrayList<Contact>();
        org.w3c.dom.NodeList nodes = getItems(offset, limit, null, null).getElementsByTagName("t:Contact");
        for (int i=0; i<nodes.getLength(); i++){
            org.w3c.dom.Node node = nodes.item(i);
            if (node.getNodeType()==1){
                contacts.add(new Contact(node));
            }
        }
        return contacts.toArray(new Contact[contacts.size()]);
    }


  //**************************************************************************
  //** getContact
  //**************************************************************************
  /** Returns a contact associated with the given exchangeID.
   */
    public Contact getContact(String exchangeID) throws ExchangeException {
        return new Contact(exchangeID, conn);
    }
}