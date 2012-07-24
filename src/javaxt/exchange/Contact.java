package javaxt.exchange;

//******************************************************************************
//**  Exchange Contact
//******************************************************************************
/**
 *   Used to represent a contact found in the Contacts Folder.
 *   http://msdn.microsoft.com/en-us/library/aa581315%28v=EXCHG.140%29.aspx
 *
 ******************************************************************************/

public class Contact extends FolderItem {


    /*
  //This is an ordered list of all the contact properties    
    private String FileAs
    private String FileAsMapping
    private String DisplayName
    private String GivenName
    private String Initials
    private String MiddleName
    private String Nickname
    private String CompleteName
    private String CompanyName
    private String EmailAddresses
    private String PhysicalAddresses
    private String PhoneNumbers
    private String AssistantName
    private String Birthday
    private String BusinessHomePage
    private String Children
    private String Companies
    private String ContactSource
    private String Department
    private String Generation
    private String ImAddresses
    private String JobTitle
    private String Manager
    private String Mileage
    private String OfficeLocation
    private String PostalAddressIndex
    private String Profession
    private String SpouseName
    private String Surname
    private String WeddingAnniversary
    private String HasPicture
     */


    private String firstName;
    private String lastName;
    private String fullName;
    private String company;
    private String title;
    private java.util.ArrayList<EmailAddress> emailAddresses = new java.util.ArrayList<EmailAddress>();
    private java.util.HashMap<String, PhoneNumber> phoneNumbers = new java.util.HashMap<String, PhoneNumber>();
    private java.util.HashMap<String, PhysicalAddress> physicalAddresses = new java.util.HashMap<String, PhysicalAddress>();
    private javaxt.utils.Date birthday;


  //**************************************************************************
  //** Constructor
  //**************************************************************************
  /** This constructor is provided for application developers who wish to
   *  extend this class.
   */
    protected Contact(){}


  //**************************************************************************
  //** Constructor
  //**************************************************************************
  /** Creates a new instance of this class using an existing Exchange Contact,
   *  effectively creating a clone.
   */
    public Contact(javaxt.exchange.Contact contact){

      //General information
        this.id = contact.id;
        this.categories = contact.categories;
        this.updates = contact.updates;
        this.lastModified = contact.lastModified;

      //Contact specific information
        this.firstName = contact.firstName;
        this.lastName = contact.lastName;
        this.fullName = contact.fullName;
        this.company = contact.company;
        this.title = contact.title;
        this.emailAddresses = contact.emailAddresses;
        this.phoneNumbers = contact.phoneNumbers;
        this.physicalAddresses = contact.physicalAddresses;
        this.birthday = contact.birthday;
    }


  //**************************************************************************
  //** Constructor
  //**************************************************************************
  /** Creates a new instance of this class
   */
    public Contact(String exchangeID, Connection conn, ExtendedProperty[] AdditionalProperties) throws ExchangeException{
        super(exchangeID, conn, AdditionalProperties);
        parseContact();
    }

  //**************************************************************************
  //** Constructor
  //**************************************************************************
  /** Creates a new instance of this class
   */
    public Contact(String exchangeID, Connection conn) throws ExchangeException{
        this(exchangeID, conn, null);
    }

  //**************************************************************************
  //** Constructor
  //**************************************************************************
  /** Creates a new instance of this class with a name
   */
    public Contact(String firstName, String lastName) {
        this.setName(firstName, lastName);
    }



  //**************************************************************************
  //** Constructor
  //**************************************************************************
  /** Creates a new instance of Contact using a node from a
   *  FindItemResponseMessage.
   */
    protected Contact(org.w3c.dom.Node contactNode) {
        super(contactNode);
        parseContact();
    }


  //**************************************************************************
  //** parseContact
  //**************************************************************************
  /** Used to parse an xml node with contact information.
   */
    private void parseContact(){
        org.w3c.dom.NodeList outerNodes = this.getChildNodes();
        for (int i=0; i<outerNodes.getLength(); i++){
            org.w3c.dom.Node outerNode = outerNodes.item(i);
            if (outerNode.getNodeType()==1){
                String nodeName = outerNode.getNodeName();
                if (nodeName.contains(":")) nodeName = nodeName.substring(nodeName.indexOf(":")+1);
                if (nodeName.equalsIgnoreCase("CompanyName")){
                    company = javaxt.xml.DOM.getNodeValue(outerNode);
                }
                else if(nodeName.equalsIgnoreCase("JobTitle")){
                    title = javaxt.xml.DOM.getNodeValue(outerNode);
                }
                else if(nodeName.equalsIgnoreCase("Birthday")){
                    try{
                        birthday = new javaxt.utils.Date(javaxt.xml.DOM.getNodeValue(outerNode));
                    }
                    catch(java.text.ParseException e){}
                }
                else if (nodeName.equalsIgnoreCase("CompleteName")){
                    org.w3c.dom.NodeList childNodes = outerNode.getChildNodes();
                    for (int j=0; j<childNodes.getLength(); j++){
                        org.w3c.dom.Node childNode = childNodes.item(j);
                        if (childNode.getNodeType()==1){
                            String childNodeName = childNode.getNodeName();
                            if (childNodeName.contains(":")) childNodeName = childNodeName.substring(childNodeName.indexOf(":")+1);
                            if (childNodeName.equalsIgnoreCase("FirstName")){
                                firstName = javaxt.xml.DOM.getNodeValue(childNode);
                            }
                            else if (childNodeName.equalsIgnoreCase("LastName")){
                                lastName = javaxt.xml.DOM.getNodeValue(childNode);
                            }
                            else if (childNodeName.equalsIgnoreCase("FullName")){
                                fullName = javaxt.xml.DOM.getNodeValue(childNode);
                            }
                        }
                    }
                }
                else if (nodeName.equalsIgnoreCase("EmailAddresses")){
                    org.w3c.dom.NodeList childNodes = outerNode.getChildNodes();
                    for (int j=0; j<childNodes.getLength(); j++){
                        org.w3c.dom.Node childNode = childNodes.item(j);
                        if (childNode.getNodeType()==1){
                            String childNodeName = childNode.getNodeName();
                            if (childNodeName.contains(":")) childNodeName = childNodeName.substring(childNodeName.indexOf(":")+1);
                            if (childNodeName.equalsIgnoreCase("Entry")){
                                String email = javaxt.xml.DOM.getNodeValue(childNode);
                                try{
                                    emailAddresses.add(new EmailAddress(email));
                                }
                                catch(ExchangeException e){}
                            }
                        }
                    }
                }
                else if (nodeName.equalsIgnoreCase("PhoneNumbers")){
                    org.w3c.dom.NodeList childNodes = outerNode.getChildNodes();
                    for (int j=0; j<childNodes.getLength(); j++){
                        try{
                            PhoneNumber phoneNumber = new PhoneNumber(childNodes.item(j));
                            phoneNumbers.put(phoneNumber.getType(), phoneNumber);
                        }
                        catch(ExchangeException e){}
                    }
                }
                else if (nodeName.equalsIgnoreCase("PhysicalAddresses")){
                    org.w3c.dom.NodeList childNodes = outerNode.getChildNodes();
                    for (int j=0; j<childNodes.getLength(); j++){
                        org.w3c.dom.Node childNode = childNodes.item(j);
                        if (childNode.getNodeType()==1){
                            PhysicalAddress address = new PhysicalAddress(childNode);
                            physicalAddresses.put(address.getType(), address);
                        }
                    }
                }
            }
        }
    }


  //**************************************************************************
  //** setName
  //**************************************************************************
  /** Used to set the first and last name for this contact.
   */
    protected void setName(String firstName, String lastName) {

        firstName = getValue(firstName);
        lastName = getValue(lastName);

        if (id!=null) {
            this.firstName = getFirstName();
            if (firstName==null && this.firstName!=null) updates.put("GivenName", null);
            if (firstName!=null && !firstName.equals(this.firstName)) updates.put("GivenName", firstName);

            this.lastName = getLastName();
            if (lastName==null && this.lastName!=null) updates.put("Surname", null);
            if (lastName!=null && !lastName.equals(this.lastName)) updates.put("Surname", lastName);

            if (updates.containsKey("GivenName") || updates.containsKey("Surname")){
                updates.put("FileAsMapping", "LastCommaFirst");
            }

        }

        this.firstName = firstName;
        this.lastName = lastName;
    }



    public String getFirstName(){
        firstName = getValue(firstName);
        return firstName;
    }

    public String getLastName(){
        lastName = getValue(lastName);
        return lastName;
    }



  //**************************************************************************
  //** setFullName
  //**************************************************************************
  /** Used to set the "Display Name" attribute for this contact.
   */
    public void setFullName(String fullName){

        fullName = getValue(fullName);

        if (id!=null) {
            this.fullName = getFullName();
            if (fullName==null && this.fullName!=null) updates.put("DisplayName", null);
            if (fullName!=null && !fullName.equals(this.fullName)) updates.put("DisplayName", fullName);
        }
        this.fullName = fullName;
    }

  //**************************************************************************
  //** getFullName
  //**************************************************************************
  /** Returns the "FullName" attribute for this contact.
   */
    public String getFullName(){
        if (fullName==null){
            String firstName = this.getFirstName();
            if (firstName==null) firstName = "";
            String lastName = this.getLastName();
            if (lastName==null) lastName="";
            fullName = (firstName + " " + lastName).trim();
            if (fullName.length()==0) fullName = null;
        }
        return fullName;
    }


  //**************************************************************************
  //** getPhysicalAddress
  //**************************************************************************
  /** Returns an array of mailing addresses associated with this contact.
   */
    public PhysicalAddress[] getPhysicalAddresses(){

      //Only include non-null values in the array
        java.util.ArrayList<PhysicalAddress> arr = new java.util.ArrayList<PhysicalAddress>();
        java.util.Iterator<String> it = physicalAddresses.keySet().iterator();
        while (it.hasNext()){
            String key = it.next();
            PhysicalAddress val = physicalAddresses.get(key);
            if (val!=null) arr.add(val);
        }
        if (arr.isEmpty()) return null;
        else return arr.toArray(new PhysicalAddress[arr.size()]);
    }


  //**************************************************************************
  //** setPhysicalAddresses
  //**************************************************************************
  /** Used to add phone numbers to a contact.
   */
    public void setPhysicalAddresses(PhysicalAddress[] physicalAddresses) {

        if (physicalAddresses==null || physicalAddresses.length==0) {
            removePhysicalAddresses();
            return;
        }

      //See whether any updates are required
        int numMatches = 0;
        int total = 0;
        for (PhysicalAddress physicalAddress : physicalAddresses){
            if (physicalAddress!=null){
                if (!physicalAddress.isEmpty()){
                    total++;
                    String type = physicalAddress.getType();
                    if (this.physicalAddresses.containsKey(type)){
                        if (this.physicalAddresses.get(type).equals(physicalAddress)) numMatches++;
                    }
                }
            }
        }

        int numAddresses = 0;
        if (this.getPhysicalAddresses()!=null) numAddresses = this.getPhysicalAddresses().length;

      //If the input array equals the current list of physicalAddresses, do nothing...
        if (numMatches==total && numMatches==numAddresses){
            return;
        }
        else {
            this.physicalAddresses.clear();
            for (PhysicalAddress physicalAddress : physicalAddresses){
                addPhysicalAddress(physicalAddress);
            }
        }
    }


  //**************************************************************************
  //** addPhysicalAddress
  //**************************************************************************
  /** Used to associate a mailing address with this contact.
   */
    public void addPhysicalAddress(PhysicalAddress address){

        if (address.isEmpty()){
            removePhysicalAddress(address);
        }
        else{

          //Check whether this is a new address
            boolean update = false;
            if (physicalAddresses.containsKey(address.getType())){
                java.util.Iterator<String> it = physicalAddresses.keySet().iterator();
                while (it.hasNext()){
                    String key = it.next();
                    if (key.equals(address.getType())){
                        PhysicalAddress addr = physicalAddresses.get(key);
                        if (addr==null || !addr.equals(address)){
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
                physicalAddresses.put(address.getType(), address);
                if (id!=null) updates.put("PhysicalAddresses", getAddressUpdates());
            }

        }
    }


  //**************************************************************************
  //** removePhysicalAddress
  //**************************************************************************
  /** Deletes an address associated with this contact.
   *  @param str Complete mailing address or an address type/category.
   */
    public void removePhysicalAddress(String str){

        java.util.Iterator<String> it = physicalAddresses.keySet().iterator();
        while (it.hasNext()){
            String type = it.next();
            if (str.equalsIgnoreCase(type)){
                physicalAddresses.put(type, null);
                if (id!=null) updates.put("PhysicalAddresses", getAddressUpdates());
                break;
            }
            else{
                PhysicalAddress address = physicalAddresses.get(type);
                if (address.equals(type)){
                    physicalAddresses.put(type, null);
                    if (id!=null) updates.put("PhysicalAddresses", getAddressUpdates());
                    break;
                }
            }
        }
    }


  //**************************************************************************
  //** removePhysicalAddress
  //**************************************************************************
  /** Deletes an address associated with this contact.
   *  @param address PhysicalAddress. Must be an exact address match.
   */
    public void removePhysicalAddress(PhysicalAddress address){
        java.util.Iterator<String> it = physicalAddresses.keySet().iterator();
        while (it.hasNext()){
            String type = it.next();
            if (physicalAddresses.get(type).equals(address)){
                physicalAddresses.put(type, null);
                if (id!=null) updates.put("PhysicalAddresses", getAddressUpdates());
                break;
            }
        }
    }


  //**************************************************************************
  //** removePhysicalAddresses
  //**************************************************************************
  /** Deletes all addresses associated with this contact.
   */
    public void removePhysicalAddresses(){
        PhysicalAddress[] addresses = getPhysicalAddresses();
        if (addresses!=null){
            for (PhysicalAddress address : addresses){
                removePhysicalAddress(address);
            }
        }
    }


  //**************************************************************************
  //** getAddressUpdates
  //**************************************************************************
  /** Returns an xml fragment with address updates or deletions. The xml
   *  fragment is later used in the updateContact() method.
   */
    private String getAddressUpdates(){

        if (physicalAddresses.isEmpty()) return "";
        else{
            StringBuffer xml = new StringBuffer();

            java.util.Iterator<String> it = physicalAddresses.keySet().iterator();
            while (it.hasNext()){
                String type = it.next();
                PhysicalAddress address = physicalAddresses.get(type);
                
                if (address==null){
                    if (id!=null){
                        xml.append(new PhysicalAddress(type).toXML("t", false));
                    }
                }
                else{
                    xml.append(address.toXML("t", false));
                }
                
            }
            return xml.toString();
        }
    }



    
  //**************************************************************************
  //** getPhoneNumbers
  //**************************************************************************
  /** Returns an array of PhoneNumbers associated with this contact.
   */
    public PhoneNumber[] getPhoneNumbers(){

      //Only include non-null values in the array
        java.util.ArrayList<PhoneNumber> arr = new java.util.ArrayList<PhoneNumber>();
        java.util.Iterator<String> it = phoneNumbers.keySet().iterator();
        while (it.hasNext()){
            String key = it.next();
            PhoneNumber val = phoneNumbers.get(key);
            if (val!=null) arr.add(val);
        }
        if (arr.isEmpty()) return null;
        else return arr.toArray(new PhoneNumber[arr.size()]);
    }


  //**************************************************************************
  //** setPhoneNumbers
  //**************************************************************************
  /** Used to add phone numbers to a contact.
   */
    public void setPhoneNumbers(PhoneNumber[] phoneNumbers) {

        if (phoneNumbers==null || phoneNumbers.length==0){
            removePhoneNumbers();
            return;
        }

      //See whether any updates are required
        int numMatches = 0;
        int total = 0;
        for (PhoneNumber phoneNumber : phoneNumbers){
            if (phoneNumber!=null){
                total++;
                String type = phoneNumber.getType();
                if (this.phoneNumbers.containsKey(type)){
                    if (this.phoneNumbers.get(type).equals(phoneNumber)) numMatches++;
                }
            }
        }

        int numPhoneNumbers = 0;
        if (this.getPhoneNumbers()!=null) numPhoneNumbers = this.getPhoneNumbers().length;

      //If the input array equals the current list of phoneNumbers, do nothing...
        if (numMatches==total && numMatches==numPhoneNumbers){
            return;
        }
        else {
            this.phoneNumbers.clear();
            for (PhoneNumber phoneNumber : phoneNumbers){
                addPhoneNumber(phoneNumber);
            }
        }
    }



  //**************************************************************************
  //** addPhoneNumber
  //**************************************************************************
  /** Used to associate a phone number with this contact.
   */
    public void addPhoneNumber(PhoneNumber phoneNumber){

      //Check whether this is a new phone number
        boolean update = false;
        if (phoneNumbers.containsKey(phoneNumber.getType())){
            java.util.Iterator<String> it = phoneNumbers.keySet().iterator();
            while (it.hasNext()){
                String key = it.next();
                if (key.equals(phoneNumber.getType())){
                    String number = null;
                    if (phoneNumbers.get(key)!=null){
                        number = phoneNumbers.get(key).getNumber();
                    }

                    if (number==null || !number.equals(phoneNumber.getNumber())){
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
            phoneNumbers.put(phoneNumber.getType(), phoneNumber);
            if (id!=null) updates.put("PhoneNumbers", getPhoneUpdates());
        }
    }



  //**************************************************************************
  //** removePhoneNumber
  //**************************************************************************

    public void removePhoneNumber(PhoneNumber phoneNumber){

      //Find any "types" associated with this phone number
        java.util.List<String> types = new java.util.ArrayList<String>();
        java.util.Iterator<String> it = phoneNumbers.keySet().iterator();
        while (it.hasNext()){
            String type = it.next();
            if (phoneNumbers.get(type).getNumber().equals(phoneNumber.getNumber())){
                types.add(type);
            }
        }

      //Update the hashmap of phone numbers
        if (!types.isEmpty()){
            it = types.iterator();
            while (it.hasNext()){
                String type = it.next();
                if (id==null) phoneNumbers.remove(type);
                else phoneNumbers.put(type, null);
            }
            if (id!=null) updates.put("PhoneNumbers", getPhoneUpdates());
        }
        
    }

  //**************************************************************************
  //** removePhoneNumbers
  //**************************************************************************

    public void removePhoneNumbers(){
        PhoneNumber[] phoneNumbers = getPhoneNumbers();
        if (phoneNumbers!=null){
            for (PhoneNumber phoneNumber : phoneNumbers){
                removePhoneNumber(phoneNumber);
            }
        }
    }


    private String getPhoneUpdates(){

        if (phoneNumbers.isEmpty()) return "";
        else{
            StringBuffer xml = new StringBuffer();

            java.util.Iterator<String> it = phoneNumbers.keySet().iterator();
            while (it.hasNext()){
                String type = it.next();
                PhoneNumber phoneNumber = phoneNumbers.get(type);

                if (phoneNumber==null){
                    if (id!=null){
                        xml.append("<t:DeleteItemField>");
                        xml.append("<t:IndexedFieldURI FieldURI=\"contacts:PhoneNumber\" FieldIndex=\"" + type + "\" /> ");
                        xml.append("</t:DeleteItemField>");
                    }
                }
                else{
                    String number = phoneNumber.getNumber();
                    xml.append("<t:SetItemField>");
                    xml.append("<t:IndexedFieldURI FieldURI=\"contacts:PhoneNumber\" FieldIndex=\"" + type + "\"/>");
                    xml.append("<t:Contact><t:PhoneNumbers><t:Entry Key=\"" + type + "\">" + number + "</t:Entry></t:PhoneNumbers></t:Contact>");
                    xml.append("</t:SetItemField>");
                }

            }
            return xml.toString();
        }
    }

    

  //**************************************************************************
  //** getEmailAddresses
  //**************************************************************************
  /** Returns an array of email addresses associated with this contact. Note
   *  that Exchange only allows 3 email addresses per contact.
   */
    public EmailAddress[] getEmailAddresses(){

      //Only include non-null values in the array
        java.util.ArrayList<EmailAddress> arr = new java.util.ArrayList<EmailAddress>();
        java.util.Iterator<EmailAddress> it = emailAddresses.iterator();
        while (it.hasNext()){
            EmailAddress emailAddress = it.next();
            if (emailAddress!=null) arr.add(emailAddress);
            
            if (arr.size()==3) break;
        }
        if (arr.isEmpty()) return null;
        else return arr.toArray(new EmailAddress[arr.size()]);
    }


  //**************************************************************************
  //** setEmailAddresses
  //**************************************************************************
  /** Used to add email Addresses to a contact.
   */
    public void setEmailAddresses(EmailAddress[] emailAddresses) {

        if (emailAddresses==null || emailAddresses.length==0){
            removeEmailAddresses();
            return;
        }

      //See whether any updates are required
        int numMatches = 0;
        int total = 0;
        for (EmailAddress emailAddress : emailAddresses){
            if (emailAddress!=null){
                total++;
                if (this.emailAddresses.contains(emailAddress)) numMatches++;
            }
        }

        int numEmails = 0;
        if (this.getEmailAddresses()!=null) numEmails = this.getEmailAddresses().length;

      //If the input array equals the current list of emailAddresses, do nothing...
        if (numMatches==total && numMatches==numEmails){
            return;
        }
        else {
            this.emailAddresses.clear();
            for (EmailAddress emailAddress : emailAddresses){
                addEmailAddress(emailAddress);
            }
        }
    }


    public void setEmailAddress(EmailAddress emailAddress){
        setEmailAddresses(new EmailAddress[]{emailAddress});
    }
    

  //**************************************************************************
  //** addEmailAddress
  //**************************************************************************
  /** Used to associate an email address with the contact. Note that Exchange
   *  only allows 3 email addresses per contact.
   */
    public void addEmailAddress(EmailAddress emailAddress) {

        if (emailAddresses.size()<3){

            if (id!=null && !emailAddresses.contains(emailAddress)){

                emailAddresses.add(emailAddress);
                updates.put("EmailAddresses", getEmailUpdates());
            }
            else{
                emailAddresses.add(emailAddress);
            }
        }
    }

  //**************************************************************************
  //** removeEmailAddress
  //**************************************************************************
  /** Used delete an email address associated with this contact.
   */
    public void removeEmailAddress(EmailAddress emailAddress){
        //emailAddresses.remove(emailAddress.toLowerCase());
    }


  //**************************************************************************
  //** removeEmailAddresses
  //**************************************************************************
  /** Used to remove all email addresses associated with this contact.
   */
    public void removeEmailAddresses(){
        EmailAddress[] addresses = getEmailAddresses();
        if (addresses!=null){
            for (EmailAddress address : addresses){
                removeEmailAddress(address);
            }
        }
    }
    

  //**************************************************************************
  //** getEmailUpdates
  //**************************************************************************
  /** Used to return an xml fragment used to update or delete email addresses.
   */
    private String getEmailUpdates(){

        if (emailAddresses.isEmpty()) return "";
        else{
            StringBuffer xml = new StringBuffer();

            int x = 1;
            java.util.Iterator<EmailAddress> it = emailAddresses.iterator();
            while (it.hasNext()){
                EmailAddress emailAddress = it.next();                
                if (emailAddress!=null){
                    if (x>3) break;

                    xml.append("<t:SetItemField>");
                    xml.append("<t:IndexedFieldURI FieldURI=\"contacts:EmailAddress\" FieldIndex=\"EmailAddress" + x + "\"/>");
                    xml.append("<t:Contact><t:EmailAddresses><t:Entry Key=\"EmailAddress" + x + "\">" + emailAddress + "</t:Entry></t:EmailAddresses></t:Contact>");
                    xml.append("</t:SetItemField>");

                    x++;
                }

            }

            
            for (int i=x; i<4; i++){
                if (id!=null){

                    xml.append("<t:DeleteItemField>");
                    xml.append("<t:IndexedFieldURI FieldURI=\"contacts:EmailAddress\" FieldIndex=\"EmailAddress" + i + "\" /> ");
                    xml.append("</t:DeleteItemField>");


                  //Apparently for emails, you have to delete property sets as well:
                  //http://social.technet.microsoft.com/Forums/en-US/exchangesvrdevelopment/thread/b57736cf-b007-49f6-b8c6-c4ba00b5cc23/
                    if (i==1){
                        for (String propertyID : new String[]{"32896", "32898", "32900"}){
                            xml.append("<t:DeleteItemField>");
                            xml.append("<t:ExtendedFieldURI PropertySetId=\"00062004-0000-0000-C000-000000000046\" PropertyId=\"" + propertyID + "\" PropertyType=\"String\"/>");
                            xml.append("</t:DeleteItemField>");
                        }                        
                        xml.append("<t:DeleteItemField>");
                        xml.append("<t:ExtendedFieldURI PropertySetId=\"00062004-0000-0000-C000-000000000046\" PropertyId=\"32901\" PropertyType=\"Binary\"/>");
                        xml.append("</t:DeleteItemField>");
                    }
                    if (i==2){
                        for (String propertyID : new String[]{"32912", "32914", "32916"}){
                            xml.append("<t:DeleteItemField>");
                            xml.append("<t:ExtendedFieldURI PropertySetId=\"00062004-0000-0000-C000-000000000046\" PropertyId=\"" + propertyID + "\" PropertyType=\"String\"/>");
                            xml.append("</t:DeleteItemField>");
                        }
                        xml.append("<t:DeleteItemField>");
                        xml.append("<t:ExtendedFieldURI PropertySetId=\"00062004-0000-0000-C000-000000000046\" PropertyId=\"32917\" PropertyType=\"Binary\"/>");
                        xml.append("</t:DeleteItemField>");
                    }
                    if (i==3){
                        for (String propertyID : new String[]{"32928", "32930", "32932"}){
                            xml.append("<t:DeleteItemField>");
                            xml.append("<t:ExtendedFieldURI PropertySetId=\"00062004-0000-0000-C000-000000000046\" PropertyId=\"" + propertyID + "\" PropertyType=\"String\"/>");
                            xml.append("</t:DeleteItemField>");
                        }
                        xml.append("<t:DeleteItemField>");
                        xml.append("<t:ExtendedFieldURI PropertySetId=\"00062004-0000-0000-C000-000000000046\" PropertyId=\"32933\" PropertyType=\"Binary\"/>");
                        xml.append("</t:DeleteItemField>");
                    }
                }
            }

            return xml.toString();
        }
    }


  //**************************************************************************
  //** setCompanyName
  //**************************************************************************
  /** Used to associate the contact with a company or organization.
   */
    public void setCompanyName(String company){
        company = getValue(company);

        if (id!=null) {
            if (company==null && this.company!=null) updates.put("CompanyName", null);
            if (company!=null && !company.equals(this.company)) updates.put("CompanyName", company);
        }
        this.company = company;
    }


  //**************************************************************************
  //** getCompanyName
  //**************************************************************************
  /** Returns the company or organization associated with this contact.
   */
    public String getCompanyName(){
        company = getValue(company);
        return company;
    }



  //**************************************************************************
  //** setBirthDay
  //**************************************************************************
  /** Used to set the date of birth using a java.util.Date.
   */
    public void setBirthDay(java.util.Date birthday){
        setBirthDay(new javaxt.utils.Date(birthday));
    }


  //**************************************************************************
  //** setBirthDay
  //**************************************************************************
  /** Used to set the date of birth using a javaxt.utils.Date.
   */
    public void setBirthDay(javaxt.utils.Date birthday){

      //Validate the date
        if (birthday!=null){
            if (birthday.getYear()<1900 || birthday.isAfter(new javaxt.utils.Date())){
                birthday = null;
            }
        }

      //Flip the timezone to UTC as needed.
        if (birthday!=null && !birthday.hasTimeStamp()){
            birthday.setTimeZone("UTC", true);
        }

      //Update birthday
        if (id!=null){
            if (birthday==null && this.birthday!=null) updates.put("Birthday", null);
            if (birthday!=null && !formatDate(birthday).equals(formatDate(this.birthday))){
                updates.put("Birthday", formatDate(birthday));
            }
        }
        this.birthday = birthday;
    }


  //**************************************************************************
  //** setBirthDay
  //**************************************************************************
  /** Used to set the date of birth using a String. If the method fails to
   *  parse the string, the new value will be ignored.
   */
    public void setBirthDay(String birthday){
        javaxt.utils.Date date = null;
        try{
            date = new javaxt.utils.Date(birthday);
        }
        catch(java.text.ParseException e){}
        setBirthDay(date);
    }


  //**************************************************************************
  //** getBirthDay
  //**************************************************************************
  /** Returns the date of birth.
   */
    public javaxt.utils.Date getBirthDay(){
        return birthday;
    }


  //**************************************************************************
  //** getAge
  //**************************************************************************
  /** Returns the age of the contact, in years.
   */
    public Integer getAge(){
        if (birthday!=null) return (int) -birthday.compareTo(new java.util.Date(), "YEAR");
        return null;
    }


  //**************************************************************************
  //** setTitle
  //**************************************************************************
  /** Used to set the job title.
   */
    public void setTitle(String title){
        title = getValue(title);

        if (id!=null) {
            this.title = getTitle();
            if (title==null && this.title!=null) updates.put("JobTitle", null);
            if (title!=null && !title.equals(this.title)) updates.put("JobTitle", title);
        }
        this.title = title;
    }
    

  //**************************************************************************
  //** getTitle
  //**************************************************************************
  /** Returns the job title associated with this contact.
   */
    public String getTitle(){
        title = getValue(title);
        return title;
    }


  //**************************************************************************
  //** save
  //**************************************************************************
  /**  Used to save/update a contact. Returns the Exchange ID for the item.
   */
    public String save(Connection conn) throws ExchangeException {

        java.util.HashMap<String, String> options = new java.util.HashMap<String, String>();
        options.put("ConflictResolution", "AutoResolve");

        if (id==null) create(conn);
        else update("Contact", "contacts", options, conn);
        return id;
    }


  //**************************************************************************
  //** delete
  //**************************************************************************
  /**  Used to delete a contact.
   */
    public void delete(Connection conn) throws ExchangeException {
        super.delete(null, conn);
    }


  //**************************************************************************
  //** create
  //**************************************************************************
  /** Used to create a new contact. 
   */
    private void create(Connection conn) throws ExchangeException {

        StringBuffer msg = new StringBuffer();        
        msg.append("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
        msg.append("<soap:Envelope xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\" xmlns:m=\"http://schemas.microsoft.com/exchange/services/2006/messages\">");
        //"<soap:Header>"<t:RequestServerVersion Version=\"Exchange2007_SP1\"/></soap:Header>"
        msg.append("<soap:Body><m:CreateItem>");
        msg.append("<m:SavedItemFolderId>");
        msg.append("<t:DistinguishedFolderId Id=\"contacts\" />"); //msg.append("<t:FolderId Id=\"" + folderID + "\"/>");
        msg.append("</m:SavedItemFolderId>");
        msg.append("<m:Items>");
        msg.append("<t:Contact>");

      
      /*
        Here's is an ordered list of all the contact properties. The first set
        is pretty generic and applies to other Exchange "Items". The second set
        is specific to contacts. WARNING -- ORDER IS VERY IMPORTANT!!! If you
        mess up the order of the properties, the save will fail - at least it
        did on my Exchange Server 2007 SP3 (8.3)

        MimeContent, ItemId, ParentFolderId, ItemClass, Subject, Sensitivity,
        Body, Attachments, DateTimeReceived, Size, Categories, Importance,
        InReplyTo, IsSubmitted, IsDraft, IsFromMe, IsResend, IsUnmodified,
        InternetMessageHeaders, DateTimeSent, DateTimeCreated, ResponseObjects,
        ReminderDueBy, ReminderIsSet, ReminderMinutesBeforeStart, DisplayCc,
        DisplayTo, HasAttachments, ExtendedProperty, Culture, EffectiveRights,
        LastModifiedName, LastModifiedTime, IsAssociated,
        WebClientReadFormQueryString, WebClientEditFormQueryString,
        ConversationId, UniqueBody

        FileAs, FileAsMapping, DisplayName, GivenName, Initials, MiddleName,
        Nickname, CompleteName, CompanyName, EmailAddresses, PhysicalAddresses,
        PhoneNumbers, AssistantName, Birthday, BusinessHomePage, Children,
        Companies, ContactSource, Department, Generation, ImAddresses, JobTitle,
        Manager, Mileage, OfficeLocation, PostalAddressIndex, Profession,
        SpouseName, Surname, WeddingAnniversary, HasPicture
      */


      //Add categories
        String[] categories = this.getCategories();
        if (categories!=null){
            msg.append("<t:Categories>");
            for (String category : categories){
                msg.append("<t:String>" + category + "</t:String>");
            }
            msg.append("</t:Categories>");
        }


      //Add extended properties
        ExtendedProperty[] properties = this.getExtendedProperties();
        if (properties!=null){
            for (ExtendedProperty property : properties){
                msg.append(property.toXML("t", "create"));
            }
        }


        msg.append("<t:FileAsMapping>LastCommaFirst</t:FileAsMapping>");
        msg.append("<t:DisplayName>" + getFullName() + "</t:DisplayName>");
        if (this.getFirstName()!=null) msg.append("<t:GivenName>" + firstName + "</t:GivenName>");
        if (this.getCompanyName()!=null) msg.append("<t:CompanyName>" + company + "</t:CompanyName>");

      //Add Email Addresses
        EmailAddress[] emails = getEmailAddresses();
        if (emails!=null){
            msg.append("<t:EmailAddresses>");
            for (int i=0; i<emails.length; i++){
                msg.append("<t:Entry Key=\"EmailAddress" + (i+1) + "\">" + emails[i] + "</t:Entry>");
            }
            msg.append("</t:EmailAddresses>");
        }


      //Add Physical Addresses
        PhysicalAddress[] addresses = getPhysicalAddresses();
        if (addresses!=null){
            msg.append("<t:PhysicalAddresses>");
            for (PhysicalAddress address : addresses){
                msg.append(address.toXML("t", true));
            }
            msg.append("</t:PhysicalAddresses>");
        }
        

      //Add Phone Numbers
        PhoneNumber[] phoneNumbers = getPhoneNumbers();
        if (phoneNumbers!=null){
            msg.append("<t:PhoneNumbers>");
            for (PhoneNumber phoneNumber : phoneNumbers){
                msg.append("<t:Entry Key=\"" + phoneNumber.getType() + "\">" + phoneNumber.getNumber() + "</t:Entry>");
            }
            msg.append("</t:PhoneNumbers>");
        }
        

        if (this.getBirthDay()!=null) msg.append("<t:Birthday>" + formatDate(birthday) + "</t:Birthday>");
        if (this.getTitle()!=null) msg.append("<t:JobTitle>" + title + "</t:JobTitle>");
        if (this.getLastName()!=null) msg.append("<t:Surname>" + lastName + "</t:Surname>");

        msg.append("</t:Contact>");
        msg.append("</m:Items>");
        msg.append("</m:CreateItem>");
        msg.append("</soap:Body>");
        msg.append("</soap:Envelope>");
        


        org.w3c.dom.Document xml = conn.execute(msg.toString());
        org.w3c.dom.NodeList nodes = xml.getElementsByTagName("t:ItemId");
        if (nodes!=null && nodes.getLength()>0){
            id = javaxt.xml.DOM.getAttributeValue(nodes.item(0), "Id");
        }

        if (id==null) throw new ExchangeException("Failed to parse ItemId while saving the contact.");
    }





    public String toString(){
        return this.getFullName();
    }
}