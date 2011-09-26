package javaxt.exchange;

//******************************************************************************
//**  Exchange Contact
//******************************************************************************
/**
 *   Used to represent a contact found in the Contacts Folder.
 *   http://msdn.microsoft.com/en-us/library/aa581315%28v=EXCHG.140%29.aspx
 *
 ******************************************************************************/

public class Contact {


    /*
  //This is an ordered list of all the contact properties
    private String MimeContent
    private String ItemId
    private String ParentFolderId
    private String ItemClass
    private String Subject
    private String Sensitivity
    private String Body
    private String Attachments
    private String DateTimeReceived
    private String Size
    private String Categories
    private String Importance
    private String InReplyTo
    private String IsSubmitted
    private String IsDraft
    private String IsFromMe
    private String IsResend
    private String IsUnmodified
    private String InternetMessageHeaders
    private String DateTimeSent
    private String DateTimeCreated
    private String ResponseObjects
    private String ReminderDueBy
    private String ReminderIsSet
    private String ReminderMinutesBeforeStart
    private String DisplayCc
    private String DisplayTo
    private String HasAttachments
    private String ExtendedProperty
    private String Culture
    private String EffectiveRights
    private String LastModifiedName
    private String LastModifiedTime
    private String IsAssociated
    private String WebClientReadFormQueryString
    private String WebClientEditFormQueryString
    private String ConversationId
    private String UniqueBody

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


    private String id;
    //private String changeKey;
    private String firstName;
    private String lastName;
    private String fullName;
    private String company;
    private String title;
    private java.util.HashSet<String> categories = new java.util.HashSet<String>();
    private java.util.ArrayList<EmailAddress> emailAddresses = new java.util.ArrayList<EmailAddress>();
    private java.util.HashMap<String, PhoneNumber> phoneNumbers = new java.util.HashMap<String, PhoneNumber>();
    private java.util.HashMap<String, PhysicalAddress> physicalAddresses = new java.util.HashMap<String, PhysicalAddress>();
    private javaxt.utils.Date birthday;
    private java.util.HashMap<String, String> updates = new java.util.HashMap<String, String>();


  //**************************************************************************
  //** resetUpdates
  //**************************************************************************
  /** Calls to 'add' and 'set' methods in this class are recorded in a hashmap.
   *  The hashmap is later used when updating a contact via the save() method.
   *  Use this method to reset the list of updates. 
   */
    protected void resetUpdates(){
        updates.clear();
    }


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
        this.id = contact.id;
        this.firstName = contact.firstName;
        this.lastName = contact.lastName;
        this.fullName = contact.fullName;
        this.company = contact.company;
        this.title = contact.title;
        this.categories = contact.categories;
        this.emailAddresses = contact.emailAddresses;
        this.phoneNumbers = contact.phoneNumbers;
        this.physicalAddresses = contact.physicalAddresses;
        this.birthday = contact.birthday;
        this.updates = contact.updates;
    }


  //**************************************************************************
  //** Constructor
  //**************************************************************************
  /** Creates a new instance of this class
   */
    public Contact(String exchangeID, Connection conn) throws ExchangeException{

        org.w3c.dom.Document xml = Folder.getItem(exchangeID, conn);
        //new javaxt.io.File("/temp/exchange-getitem2.xml").write(xml);
        org.w3c.dom.NodeList nodes = xml.getElementsByTagName("t:Contact");
        for (int i=0; i<nodes.getLength(); i++){
            org.w3c.dom.Node node = nodes.item(i);
            if (node.getNodeType()==1) parseContact(node);
        }
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
        parseContact(contactNode);
    }


  //**************************************************************************
  //** parseContact
  //**************************************************************************
  /** Used to parse an xml node with contact information.
   */
    private void parseContact(org.w3c.dom.Node contactNode){
        org.w3c.dom.NodeList outerNodes = contactNode.getChildNodes();
        for (int i=0; i<outerNodes.getLength(); i++){
            org.w3c.dom.Node outerNode = outerNodes.item(i);
            if (outerNode.getNodeType()==1){
                String nodeName = outerNode.getNodeName();
                if (nodeName.contains(":")) nodeName = nodeName.substring(nodeName.indexOf(":")+1);
                if (nodeName.equalsIgnoreCase("ItemId")){
                    id = javaxt.xml.DOM.getAttributeValue(outerNode, "Id");
                    //changeKey = javaxt.xml.DOM.getAttributeValue(outerNode, "ChangeKey");
                }
                else if(nodeName.equalsIgnoreCase("CompanyName")){
                    company = javaxt.xml.DOM.getNodeValue(outerNode);
                }
                else if(nodeName.equalsIgnoreCase("JobTitle")){
                    title = javaxt.xml.DOM.getNodeValue(outerNode);
                }
                else if(nodeName.equalsIgnoreCase("Birthday")){
                    javaxt.utils.Date date = new javaxt.utils.Date(javaxt.xml.DOM.getNodeValue(outerNode));
                    if (!date.failedToParse()) birthday = date;
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
                else if (nodeName.equalsIgnoreCase("Categories")){
                    org.w3c.dom.NodeList childNodes = outerNode.getChildNodes();
                    for (int j=0; j<childNodes.getLength(); j++){
                        org.w3c.dom.Node childNode = childNodes.item(j);
                        if (childNode.getNodeType()==1){
                            categories.add(childNode.getTextContent());
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

        if (firstName!=null){
            firstName = firstName.trim();
            if (firstName.length()==0) firstName = null;
        }

        if (lastName!=null){
            lastName = lastName.trim();
            if (lastName.length()==0) lastName = null;
        }


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
        if (firstName!=null){
            firstName = firstName.trim();
            if (firstName.length()==0) firstName = null;
        }
        return firstName;
    }

    public String getLastName(){
        if (lastName!=null){
            lastName = lastName.trim();
            if (lastName.length()==0) lastName = null;
        }
        return lastName;
    }



  //**************************************************************************
  //** setFullName
  //**************************************************************************
  /** Used to set the "Display Name" attribute for this contact.
   */
    public void setFullName(String fullName){

        if (fullName!=null){
            fullName = fullName.trim();
            if (fullName.length()==0) fullName = null;
        }

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
        return arr.toArray(new PhysicalAddress[arr.size()]);
    }


  //**************************************************************************
  //** setPhysicalAddresses
  //**************************************************************************
  /** Used to add phone numbers to a contact.
   */
    public void setPhysicalAddresses(PhysicalAddress[] physicalAddresses) {

        if (physicalAddresses==null || physicalAddresses.length==0) return; //removephysicalAddresses() ???

      //See whether any updates are required
        int numMatches = 0;
        int total = 0;
        for (PhysicalAddress physicalAddress : physicalAddresses){
            if (physicalAddress!=null){
                total++;
                String type = physicalAddress.getType();
                if (this.physicalAddresses.containsKey(type)){
                    if (this.physicalAddresses.get(type).equals(physicalAddress)) numMatches++;
                }
            }
        }

      //If the input array equals the current list of physicalAddresses, do nothing...
        if (numMatches==total && numMatches==this.getPhysicalAddresses().length){
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
        return arr.toArray(new PhoneNumber[arr.size()]);
    }


  //**************************************************************************
  //** setPhoneNumbers
  //**************************************************************************
  /** Used to add phone numbers to a contact.
   */
    public void setPhoneNumbers(PhoneNumber[] phoneNumbers) {

        if (phoneNumbers==null || phoneNumbers.length==0) return; //removephoneNumbers() ???

      //See whether any updates are required
        int numMatches = 0;
        int total = 0;
        for (PhoneNumber phoneNumber : phoneNumbers){
            System.out.println(phoneNumber);
            if (phoneNumber!=null){
                total++;
                String type = phoneNumber.getType();
                if (this.phoneNumbers.containsKey(type)){
                    if (this.phoneNumbers.get(type).equals(phoneNumber)) numMatches++;
                }
            }
        }

      //If the input array equals the current list of phoneNumbers, do nothing...
        if (numMatches==total && numMatches==this.getPhoneNumbers().length){
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
        return arr.toArray(new EmailAddress[arr.size()]);
    }


  //**************************************************************************
  //** setEmailAddresses
  //**************************************************************************
  /** Used to add email Addresses to a contact.
   */
    public void setEmailAddresses(EmailAddress[] emailAddresses) {

        if (emailAddresses==null || emailAddresses.length==0) return; //removeEmailAddresses() ???

      //See whether any updates are required
        int numMatches = 0;
        int total = 0;
        for (EmailAddress emailAddress : emailAddresses){
            if (emailAddress!=null){
                total++;
                if (this.emailAddresses.contains(emailAddress)) numMatches++;
            }
        }

      //If the input array equals the current list of emailAddresses, do nothing...
        if (numMatches==total && numMatches==this.getEmailAddresses().length){
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
        if (company!=null){
            company = company.trim();
            if (company.length()==0) company = null;
        }

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
        if (company!=null){
            company = company.trim();
            if (company.length()==0) company = null;
        }
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
        if (birthday!=null){
            date = new javaxt.utils.Date(birthday);
            if (date.failedToParse()){
                if (this.birthday!=null) date = this.birthday;
                else date = null;
            }
        }
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
        if (title!=null){
            title = title.trim();
            if (title.length()==0) title = null;
        }

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
        if (title!=null){
            title = title.trim();
            if (title.length()==0) title = null;
        }
        return title;
    }



    public void setExchangeID(String id){
        if (id!=null){
            id = id.trim();
            if (id.length()<25) id = null;
        }
        this.id = id;
    }


    public String getExchangeID(){
        return id;
    }


  //**************************************************************************
  //** setCategories
  //**************************************************************************
  /** Used to add categories to a contact.
   */
    public void setCategories(String[] categories){
        
        if (categories==null || categories.length==0) return; //removeCategories() ???


      //See whether any updates are required
        int numMatches = 0;
        if (categories.length==this.categories.size()){
            for (String category : categories){
                if (category!=null){
                    if (this.categories.contains(category)) numMatches++;
                }
            }
        }

      //If the input array equals the current list of categories, do nothing...
        if (numMatches==categories.length){
            return;
        }
        else {
            this.categories.clear();
            for (String category : categories){
                addCategory(category);
            }
        }
    }


    public void setCategory(String category){
        setCategories(new String[]{category});
    }

  //**************************************************************************
  //** addCategory
  //**************************************************************************
  /** Used to add a category to a contact.
   */
    public void addCategory(String category){

        if (category!=null){
            category = category.trim();
            if (category.length()==0) category = null;
        }
        if (category==null) return;


        if (id!=null && !categories.contains(category)){

            categories.add(category);
            
            StringBuffer xml = new StringBuffer();
            java.util.Iterator<String> it = categories.iterator();
            while (it.hasNext()){
                xml.append("<t:String>" + it.next() + "</t:String>");
            }
            updates.put("Categories", xml.toString());
        }
        else{
            categories.add(category);
        }
    }

    public String[] getCategories(){
        return categories.toArray(new String[categories.size()]);
    }

   

  //**************************************************************************
  //** save
  //**************************************************************************
  /**  Used to save/update a contact. Returns the Exchange ID for the item.
   */
    public String save(Connection conn) throws ExchangeException{

        if (id==null){
            return addContact(conn);
        }
        else{
            updateContact(conn);
            return id;
        }
    }


  //**************************************************************************
  //** updateContact
  //**************************************************************************
  /** Used to update the contact.
   */
    private void updateContact(Connection conn) throws ExchangeException {

        if (updates.isEmpty()) return;


        StringBuffer msg = new StringBuffer();
        msg.append("<?xml version=\"1.0\" encoding=\"utf-8\" ?>");
        msg.append("<soap:Envelope xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\" xmlns:m=\"http://schemas.microsoft.com/exchange/services/2006/messages\">");
        msg.append("<soap:Body>");
        msg.append("<m:UpdateItem ConflictResolution=\"AutoResolve\">");
        msg.append("<m:ItemChanges>");
        msg.append("<t:ItemChange>");
        msg.append("<t:ItemId Id=\"" + id + "\" ChangeKey=\"" + getChangeKey(conn) + "\" />"); //
        msg.append("<t:Updates>");


        String namespace = "contacts";
        java.util.Iterator<String> it = updates.keySet().iterator();
        while (it.hasNext()){
            String key = it.next();
            String value = updates.get(key);
            if (key.equalsIgnoreCase("Categories")) namespace = "item";

            if (value==null){
                System.out.println("Delete " + key);
                msg.append("<t:DeleteItemField>");
                msg.append("<t:FieldURI FieldURI=\"" + namespace + ":" + key + "\"/>");
                msg.append("</t:DeleteItemField>");                
            }
            else{
                System.out.println("Update " + key);

                if (value.trim().startsWith("<t:SetItemField") || value.trim().startsWith("<t:DeleteItemField")){
                    msg.append(value);
                }
                else{
                    msg.append("<t:SetItemField>");
                    msg.append("<t:FieldURI FieldURI=\"" + namespace + ":" + key + "\" />");
                    msg.append("<t:Contact>");
                    msg.append("<t:" + key + ">" + value + "</t:" + key + ">");
                    msg.append("</t:Contact>");
                    msg.append("</t:SetItemField>");
                }
                
            }
        }

        msg.append("</t:Updates>");
        msg.append("</t:ItemChange>");
        msg.append("</m:ItemChanges>");
        msg.append("</m:UpdateItem>");
        msg.append("</soap:Body>");
        msg.append("</soap:Envelope>");

System.out.println(msg + "\r\n");

        updates.clear();
//if (true) return;

        conn.execute(msg.toString());
        //javaxt.http.Response response = conn.execute(msg.toString());

        //String txt = response.getText();
        //System.out.println(txt);

    }



  //**************************************************************************
  //** addContact
  //**************************************************************************
  /** Used to create a new contact. Returns an id for the newly created
   *  contact or null is there was an error.
   */
    private String addContact(Connection conn) throws ExchangeException {

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
        if (!categories.isEmpty()){
            msg.append("<t:Categories>");
            java.util.Iterator<String> it = categories.iterator();
            while (it.hasNext()){
                msg.append("<t:String>" + it.next() + "</t:String>");
            }
            msg.append("</t:Categories>");
        }

        msg.append("<t:FileAsMapping>LastCommaFirst</t:FileAsMapping>");
        msg.append("<t:DisplayName>" + getFullName() + "</t:DisplayName>");
        if (firstName!=null) msg.append("<t:GivenName>" + firstName + "</t:GivenName>");
        if (company!=null) msg.append("<t:CompanyName>" + company + "</t:CompanyName>");

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
        if (!physicalAddresses.isEmpty()){
            msg.append("<t:PhysicalAddresses>");
            for (PhysicalAddress address : getPhysicalAddresses()){
                msg.append(address.toXML("t", true));
            }
            msg.append("</t:PhysicalAddresses>");
        }
        

      //Add Phone Numbers
        if (!phoneNumbers.isEmpty()){
            msg.append("<t:PhoneNumbers>");
            for (PhoneNumber phoneNumber : getPhoneNumbers()){
                msg.append("<t:Entry Key=\"" + phoneNumber.getType() + "\">" + phoneNumber.getNumber() + "</t:Entry>");
            }
            msg.append("</t:PhoneNumbers>");
        }
        

        if (birthday!=null) msg.append("<t:Birthday>" + formatDate(birthday) + "</t:Birthday>");
        if (title!=null) msg.append("<t:JobTitle>" + title + "</t:JobTitle>");
        if (lastName!=null) msg.append("<t:Surname>" + lastName + "</t:Surname>");        

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
        
        return id;
    }


  //**************************************************************************
  //** delete
  //**************************************************************************
  /** Used to delete a contact.
   */
    public void delete(Connection conn) throws ExchangeException {
        String msg =
        "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
        + "<soap:Envelope xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\" xmlns:m=\"http://schemas.microsoft.com/exchange/services/2006/messages\">"
        //+ "<soap:Header><t:RequestServerVersion Version=\"Exchange2007_SP1\"/></soap:Header>"
        + "<soap:Body>"
        + "<m:DeleteItem DeleteType=\"MoveToDeletedItems\">"
        + "<m:ItemIds>"
        + "<t:ItemId Id=\"" + id + "\"/></m:ItemIds>" //ChangeKey=\"EQAAABYAAAA9IPsqEJarRJBDywM9WmXKAAV+D0Dq\"
        + "</m:DeleteItem>"
        + "</soap:Body>"
        + "</soap:Envelope>";
        conn.execute(msg);
    }





  //**************************************************************************
  //** formatDate
  //**************************************************************************
  /** Used to format a date into a string that Exchange Web Services can
   *  understand.
   */
    private String formatDate(javaxt.utils.Date date){
        if (date==null) return null;
        String d = date.toString("yyyy-MM-dd HH:mm:ssZ").replace(" ", "T");
        return d.substring(0, d.length()-2) + ":" + d.substring(d.length()-2);
    }



  //**************************************************************************
  //** getChangeKey
  //**************************************************************************
  /** Used to retrieve the latest ChangeKey for this contact. The ChangeKey is
   *  required to update an existing contact.
   */
    private String getChangeKey(Connection conn){

        try{
            org.w3c.dom.Document xml = Folder.getItem(id, conn);
            org.w3c.dom.NodeList nodes = xml.getElementsByTagName("t:Contact");
            for (int i=0; i<nodes.getLength(); i++){
                org.w3c.dom.Node node = nodes.item(i);
                if (node.getNodeType()==1){

                    org.w3c.dom.NodeList outerNodes = node.getChildNodes();
                    for (int j=0; j<outerNodes.getLength(); j++){
                        org.w3c.dom.Node outerNode = outerNodes.item(j);
                        if (outerNode.getNodeType()==1){
                            String nodeName = outerNode.getNodeName();
                            if (nodeName.contains(":")) nodeName = nodeName.substring(nodeName.indexOf(":")+1);
                            if (nodeName.equalsIgnoreCase("ItemId")){
                                return javaxt.xml.DOM.getAttributeValue(outerNode, "ChangeKey");
                            }
                        }
                    }

                }

            }
        }
        catch(ExchangeException e){

        }

        return null;
    }




    public String toString(){
        return this.getFullName();
    }



}