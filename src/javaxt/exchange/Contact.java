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
    private java.util.HashSet<String> emailAddresses = new java.util.HashSet<String>();
    private java.util.HashMap<String, String> phoneNumbers = new java.util.HashMap<String, String>();
    private java.util.ArrayList<PhysicalAddress> physicalAddresses = new java.util.ArrayList<PhysicalAddress>();
    private javaxt.utils.Date birthday;
    //private String folderID;
    private java.util.HashMap<String, String> updates = new java.util.HashMap<String, String>();


    protected Contact(){}


  //**************************************************************************
  //** Constructor
  //**************************************************************************
  /** Creates a new instance of this class
   */
    public Contact(String exchangeID, Connection conn) {

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
        this.firstName = firstName;
        this.lastName = lastName;
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
                                //System.out.println(email);
                                emailAddresses.add(email.toLowerCase());
                            }
                        }
                    }
                }
                else if (nodeName.equalsIgnoreCase("PhoneNumbers")){
                    org.w3c.dom.NodeList childNodes = outerNode.getChildNodes();
                    for (int j=0; j<childNodes.getLength(); j++){
                        org.w3c.dom.Node childNode = childNodes.item(j);
                        if (childNode.getNodeType()==1){
                            String childNodeName = childNode.getNodeName();
                            if (childNodeName.contains(":")) childNodeName = childNodeName.substring(childNodeName.indexOf(":")+1);
                            if (childNodeName.equalsIgnoreCase("Entry")){
                                String phoneNumber = javaxt.xml.DOM.getNodeValue(childNode);
                                String type = javaxt.xml.DOM.getAttributeValue(childNode, "Key");
                                phoneNumbers.put(phoneNumber, type);
                            }
                        }
                    }
                }
                else if (nodeName.equalsIgnoreCase("PhysicalAddresses")){
                    org.w3c.dom.NodeList childNodes = outerNode.getChildNodes();
                    for (int j=0; j<childNodes.getLength(); j++){
                        org.w3c.dom.Node childNode = childNodes.item(j);
                        if (childNode.getNodeType()==1){
                            physicalAddresses.add(new PhysicalAddress(childNode));
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

    protected void setName(String firstName, String lastName) {
        this.firstName = firstName;
        this.lastName = lastName;
    }

    public String getFirstName(){
        return firstName;
    }

    public String getLastName(){
        return lastName;
    }



  //**************************************************************************
  //** setFullName
  //**************************************************************************
  /** Used to set the "Display Name" attribute for this contact.
   */
    public void setFullName(String fullName){
        if (id!=null) {
            if (fullName==null) updates.put("DisplayName", null);
            else if (!fullName.equals(this.fullName)) updates.put("DisplayName", fullName);
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
  //** getPhoneNumbers
  //**************************************************************************
  /** Returns a hashmap with all known phone numbers associated with this
   *  contact where the key is the number and the value is the type.
   */
    public java.util.HashMap<String, String> getPhoneNumbers(){
        return this.phoneNumbers;
    }


    public void setPhoneNumber(String phoneNumber, String type){
        //phoneNumbers.clear();
        addPhoneNumber(phoneNumber, type);
    }
    

  //**************************************************************************
  //** addPhoneNumber
  //**************************************************************************
  /** Used to associate a phone number with this contact.
   * @param phoneNumber The phone number (e.g. 555-555-5555)
   * @param type Type of phone number. Options include:
   * <ul>
   * <li>AssistantPhone</li><li>BusinessFax</li><li>BusinessPhone</li>
   * <li>BusinessPhone2</li><li>Callback</li><li>CarPhone</li>
   * <li>CompanyMainPhone</li><li>HomeFax</li><li>HomePhone</li>
   * <li>HomePhone2</li><li>Isdn</li><li>MobilePhone</li><li>OtherFax</li>
   * <li>OtherTelephone</li><li>Pager</li><li>PrimaryPhone</li>
   * <li>RadioPhone</li><li>Telex</li><li>TtyTddPhone</li>
   * </ul>
   */
    public void addPhoneNumber(String phoneNumber, String type){

        if (phoneNumber==null) return;
        else phoneNumber = phoneNumber.trim();

      //Make sure there are enough number in the string to represent an actual phone number
        String[] numbers = new String[]{"0","1","2","3","4","5","6","7","8","9"};
        int numCount = 0;
        for (int i=0; i<phoneNumber.length(); i++){
            for (String num : numbers){
                if (phoneNumber.substring(i, i+1).equals(num)) numCount++;
            }
        }
        if (numCount<7) return;

        //System.out.println(phoneNumber + " (" + type + ")");

        type = type.toUpperCase();
        if (type.contains("ASSISTANT")) type = "AssistantPhone";
        else if (type.contains("BUSINESSFAX")) type = "BusinessFax";
        else if (type.contains("BUSINESS") || type.contains("WORK") || type.contains("OFFICE")) type = "BusinessPhone";
        //else if (type.contains("BUSINESS2")) type = "BusinessPhone2";
        else if (type.contains("CALLBACK")) type = "Callback";
        else if (type.contains("CAR")) type = "CarPhone";
        //else if (type.contains("COMPANYMAIN")) type = "CompanyMainPhone";
        else if (type.equals("HOMEFAX")) type = "HomeFax";
        else if (type.contains("HOME")) type = "HomePhone";
        //else if (type.contains("HOME2")) type = "HomePhone2";
        else if (type.contains("ISDN")) type = "Isdn";
        else if (type.contains("MOBILE")) type = "MobilePhone";
        else if (type.equals("OTHERFAX")) type = "OtherFax";
        else if (type.contains("PAGER")) type = "Pager";
        else if (type.contains("PRIMARY")) type = "PrimaryPhone";
        else if (type.contains("RADIO")) type = "RadioPhone";
        else if (type.contains("TELEX")) type = "Telex";
        else if (type.contains("TTYTDD")) type = "TtyTddPhone";
        else type = "OtherTelephone";

        //System.out.println(phoneNumber + " (" + type + ")");

        if (id!=null && !phoneNumbers.containsKey(phoneNumber)){
            StringBuffer xml = new StringBuffer();
            java.util.Iterator<String> phone = phoneNumbers.keySet().iterator();
            while (phone.hasNext()){
                String number = phone.next();
                type = phoneNumbers.get(number);

                //xml.append("<t:Entry Key=\"" + type + "\">" + number + "</t:Entry>");
                xml.append("<t:IndexedFieldURI FieldURI=\"contacts:PhoneNumber\" FieldIndex=\"" + type + "\"/>");
                xml.append("<t:Contact><t:PhoneNumbers><t:Entry Key=\"" + type + "\">" + number + "</t:Entry></t:PhoneNumbers></t:Contact>");

            }
            updates.put("PhoneNumbers", xml.toString());
        }
        

        phoneNumbers.put(phoneNumber, type);
    }




    public String[] getEmailAddresses(){
        if (emailAddresses.size()==0) return null;
        return emailAddresses.toArray(new String[emailAddresses.size()]);
    }


    public void setEmailAddress(String emailAddress){
        emailAddresses.clear();
        addEmailAddress(emailAddress);
    }

    public void addEmailAddress(String emailAddress){
        if (emailAddress!=null) {
            emailAddress = emailAddress.trim();
            if (emailAddress.contains("@")){
                emailAddresses.add(emailAddress.toLowerCase());
            }
        }
    }

    public void removeEmailAddress(String emailAddress){
        emailAddresses.remove(emailAddress.toLowerCase());
    }


  //**************************************************************************
  //** setCompanyName
  //**************************************************************************
  /** Used to associate the contact with a company or organization.
   */
    public void setCompanyName(String company){
        if (id!=null) {
            if (company==null) updates.put("CompanyName", null);
            else if (!company.equals(this.company)) updates.put("CompanyName", company);
        }
        this.company = company;
    }

    public String getCompanyName(){
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
            if (birthday==null) updates.put("Birthday", null);
            else if (!formatDate(birthday).equals(formatDate(this.birthday))){
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
        if (id!=null) {
            if (title==null) updates.put("JobTitle", null);
            else if (!title.equals(this.title)) updates.put("JobTitle", title);
        }
        this.title = title;
    }

    public String getTitle(){
        return title;
    }



    public void setAddress(PhysicalAddress address){
        this.physicalAddresses.clear();
        this.physicalAddresses.add(address);
    }


    public void setExchangeID(String id){
        this.id = id;
    }


    public String getExchangeID(){
        return id;
    }



  //**************************************************************************
  //** addCategory
  //**************************************************************************
  /** Used to add a category to a contact.
   */
    public void addCategory(String category){
        
        if (id!=null && !categories.contains(category)){
            StringBuffer xml = new StringBuffer();
            java.util.Iterator<String> it = categories.iterator();
            while (it.hasNext()){
                xml.append("<t:String>" + it.next() + "</t:String>");
            }
            updates.put("Categories", xml.toString());
        }
        categories.add(category);
    }

    public String[] getCategories(){
        return categories.toArray(new String[categories.size()]);
    }

   

  //**************************************************************************
  //** save
  //**************************************************************************
  /**  Used to save/update a contact. Returns the Exchange ID for the item.
   */
    public String save(Connection conn){

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
    private void updateContact(Connection conn){

        if (updates.isEmpty()) return;

        String changeKey = getChangeKey(conn);

        StringBuffer msg = new StringBuffer();
        msg.append("<?xml version=\"1.0\" encoding=\"utf-8\" ?>");
        msg.append("<soap:Envelope xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\" xmlns:m=\"http://schemas.microsoft.com/exchange/services/2006/messages\">");
        msg.append("<soap:Body>");
        msg.append("<m:UpdateItem ConflictResolution=\"AutoResolve\">");
        msg.append("<m:ItemChanges>");
        msg.append("<t:ItemChange>");
        msg.append("<t:ItemId Id=\"" + id + "\" ChangeKey=\"" + changeKey + "\" />"); //
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
                //System.out.println(value);
                
                
                if (value.trim().startsWith("<t:IndexedFieldURI")){
                    for (String field : value.split("<t:IndexedFieldURI")){
                        if (field.trim().length()>0){
                            msg.append("<t:SetItemField>");
                            msg.append("<t:IndexedFieldURI");
                            msg.append(field);
                            msg.append("</t:SetItemField>");
                        }
                    }
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
System.out.println(msg);

        updates.clear();
//if (true) return;

        javaxt.http.Response response = conn.execute(msg.toString());

        int status = response.getStatus();
        if (status>=400 && status<600){
            try{
                java.io.InputStream errorStream = response.getErrorStream();
                java.io.BufferedReader buf = new java.io.BufferedReader (
                                    new java.io.InputStreamReader(errorStream));

                try {
                    String line;
                    while  ((line = buf.readLine())!= null) {
                        System.out.println(line);
                    }
                }
                catch (java.io.IOException e) {}

                errorStream.close();
                buf.close();
            }
            catch(Exception e){
                e.printStackTrace();
            }
        }
        else{


            String txt = response.getText();
            System.out.println(txt);
            //javaxt.io.File file = new javaxt.io.File("/temp/exchange-newcontact.xml");
            //file.write(txt);

        }


    }



  //**************************************************************************
  //** addContact
  //**************************************************************************
  /** Used to create a new contact. Returns an id for the newly created
   *  contact or null is there was an error.
   */
    private String addContact(Connection conn){

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


        //WARNING -- ORDER IS VERY IMPORTANT!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!


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
        String[] emails = getEmailAddresses();
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
            java.util.Iterator<PhysicalAddress> it = physicalAddresses.iterator();
            while (it.hasNext()){
                msg.append(it.next().toXML("t"));
            }
            msg.append("</t:PhysicalAddresses>");
        }



      //Add Phone Numbers
        if (!phoneNumbers.isEmpty()){
            msg.append("<t:PhoneNumbers>");
            java.util.Iterator<String> phone = phoneNumbers.keySet().iterator();
            while (phone.hasNext()){
                String number = phone.next();
                String type = phoneNumbers.get(number);
                msg.append("<t:Entry Key=\"" + type + "\">" + number + "</t:Entry>");
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
        


        javaxt.http.Response response = conn.execute(msg.toString());

        int status = response.getStatus();
        if (status>=400 && status<600){
            try{
                java.io.InputStream errorStream = response.getErrorStream();
                java.io.BufferedReader buf = new java.io.BufferedReader (
                                    new java.io.InputStreamReader(errorStream));

                try {
                    String line;
                    while  ((line = buf.readLine())!= null) {
                        System.out.println(line);
                    }
                }
                catch (java.io.IOException e) {}

                errorStream.close();
                buf.close();
            }
            catch(Exception e){
                e.printStackTrace();
            }
        }
        else{

            /*
            String txt = response.getText();
            System.out.println(txt);
            javaxt.io.File file = new javaxt.io.File("/temp/exchange-newcontact.xml");
            file.write(txt);
            */
            org.w3c.dom.Document xml = response.getXML();
            org.w3c.dom.NodeList nodes = xml.getElementsByTagName("t:ItemId");
            if (nodes!=null && nodes.getLength()>0){
                String id = javaxt.xml.DOM.getAttributeValue(nodes.item(0), "Id");
                //System.out.println(id);
                return id;
            }

        }
        



        //ItemId


        return null;

    }


  //**************************************************************************
  //** delete
  //**************************************************************************
  /** Used to delete a contact.
   */
    public void delete(Connection conn){
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

        return null;
    }




    public String toString(){
        return this.getFullName();
    }



}