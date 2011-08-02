package javaxt.exchange;

//******************************************************************************
//**  Exchange Contact
//******************************************************************************
/**
 *   Used to represent a contact found in the Contacts Folder.
 *
 ******************************************************************************/

public class Contact {

    private String id;
    private String firstName;
    private String lastName;
    private String fullName;
    private String company;
    private String title;
    private java.util.HashSet<String> emailAddresses = new java.util.HashSet<String>();
    private java.util.HashMap<String, String> phoneNumbers = new java.util.HashMap<String, String>();
    private java.util.ArrayList<PhysicalAddress> physicalAddresses = new java.util.ArrayList<PhysicalAddress>();
    private java.util.Date birthday;
    //private String folderID;

    protected Contact(){}


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
        org.w3c.dom.NodeList outerNodes = contactNode.getChildNodes();
        for (int i=0; i<outerNodes.getLength(); i++){
            org.w3c.dom.Node outerNode = outerNodes.item(i);
            if (outerNode.getNodeType()==1){
                String nodeName = outerNode.getNodeName();
                if (nodeName.contains(":")) nodeName = nodeName.substring(nodeName.indexOf(":")+1);
                if (nodeName.equalsIgnoreCase("ItemId")){
                    id = javaxt.xml.DOM.getAttributeValue(outerNode, "Id");
                }
                else if(nodeName.equalsIgnoreCase("CompanyName")){
                    company = javaxt.xml.DOM.getNodeValue(outerNode);
                }
                else if(nodeName.equalsIgnoreCase("JobTitle")){
                    title = javaxt.xml.DOM.getNodeValue(outerNode);
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
                                //System.out.println(phoneNumber + " (" + type + ")");
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

    public void setFullName(String fullName){
        this.fullName = fullName;
    }

    public String getFullName(){
        if (fullName==null){
            String firstName = this.getFirstName();
            if (firstName==null) firstName = "";
            String lastName = this.getLastName();
            if (lastName!=null) lastName="";
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
        phoneNumbers.clear();
        addPhoneNumber(phoneNumber, type);
    }
    

    public void addPhoneNumber(String phoneNumber, String type){
        //type = "MobilePhone";
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


    public void setCompanyName(String company){
        this.company = company;
    }

    public String getCompanyName(){
        return company;
    }



    public void setBirthDay(java.util.Date birthday){
        this.birthday = birthday;
    }

    public java.util.Date getBirthDay(){
        return birthday;
    }
    

    public String getTitle(){
        return title;
    }

    public void setTitle(String title){
        this.title = title;
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




    public String save(Connection conn){

        if (id==null){
            return addContact(conn);
        }
        else{
            //update
        }

        return null;
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


        //msg.append("<t:Subject>John ZZmith2</t:Subject>");
        msg.append("<t:FileAsMapping>LastCommaFirst</t:FileAsMapping>");
        //msg.append("<t:DisplayName>John ZZmith</t:DisplayName>");
        //msg.append("<t:GivenName>John</t:GivenName><t:Surname>ZZmith</t:Surname>");
        
      //Set General Information (order is important!
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

    public void delete(Connection conn){
        String msg =
        "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
        + "<soap:Envelope xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\" xmlns:m=\"http://schemas.microsoft.com/exchange/services/2006/messages\">"
        //+ "<soap:Header><t:RequestServerVersion Version=\"Exchange2007_SP1\"/></soap:Header>"
        + "<soap:Body>"
        + "<m:DeleteItem DeleteType=\"MoveToDeletedItems\">"
        + "<m:ItemIds>"
        + "<t:ItemId Id=\"" + id + "\" ChangeKey=\"EQAAABYAAAA9IPsqEJarRJBDywM9WmXKAAV+D0Dq\"/></m:ItemIds>"
        + "</m:DeleteItem>"
        + "</soap:Body>"
        + "</soap:Envelope>";

        //conn.execute(msg);

    }




    public String toString(){
        if (firstName!=null && lastName!=null){
            return lastName + ", " + firstName;
        }
        else{
            if (firstName!=null) return firstName;
            else return lastName;
        }
    }



}