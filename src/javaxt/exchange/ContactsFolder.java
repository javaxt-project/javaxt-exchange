package javaxt.exchange;

//******************************************************************************
//**  ContactFolder Class
//******************************************************************************
/**
 *   Enter class description here
 *
 ******************************************************************************/

public class ContactsFolder extends Folder {

    private Connection conn;

  //**************************************************************************
  //** Constructor
  //**************************************************************************
  /** Creates a new instance of ContactFolder. */

    public ContactsFolder(Connection conn) {
        super("contacts", conn);
        this.conn = conn;
    }

    
    public Contact[] getContacts(){

        java.util.ArrayList<Contact> contacts = new java.util.ArrayList<Contact>();
        
        int offset = 0;
        int maxRecords = 25;
        while(true){

            Contact[] arr = getContacts(maxRecords, offset);
            for (int i=0; i<arr.length; i++){
                contacts.add(arr[i]);
            }
            if (arr.length<maxRecords) break;
            else offset+=maxRecords;
        }

        return contacts.toArray(new Contact[contacts.size()]);
    }



    public Contact[] getContacts(int numEntries, int offset){

        java.util.ArrayList<Contact> contacts = new java.util.ArrayList<Contact>();
        org.w3c.dom.NodeList nodes = getItems(conn, numEntries, offset).getElementsByTagName("t:Contact");
        //org.w3c.dom.NodeList nodes = new javaxt.io.File("/temp/exchange-findItem.xml").getXML().getElementsByTagName("t:Contact");

        int numRecordsReturned = 0;
        for (int i=0; i<nodes.getLength(); i++){
            org.w3c.dom.Node node = nodes.item(i);
            if (node.getNodeType()==1){
                contacts.add(new Contact(node));
                numRecordsReturned++;
            }
        }
        return contacts.toArray(new Contact[contacts.size()]);
    }


  //**************************************************************************
  //** getContact
  //**************************************************************************
  /** GetItem request */
    
    public Contact getContact(String exchangeID){
        return new Contact(exchangeID, conn);
    }
}