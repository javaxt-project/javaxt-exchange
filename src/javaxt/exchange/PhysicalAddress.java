package javaxt.exchange;

//******************************************************************************
//**  Address Class
//******************************************************************************
/**
 *   Used to represent a physical address for a given contact.
 *
 ******************************************************************************/

public class PhysicalAddress {

    /*
    public static final String HOME_ADDRESS = "Home";
    public static final String BUSINESS_ADDRESS = "Business";
    public static final String OTHER_ADDRESS = "Other";
    */

    private String type;
    private java.util.ArrayList<String> street = new java.util.ArrayList<String>();
    private String city;
    private String state;
    private String country;
    private String postalCode;


  //**************************************************************************
  //** Constructor
  //**************************************************************************
  /** Creates a new instance of Address. 
   */
    public PhysicalAddress(String type){
        setType(type);
    }
    
    public PhysicalAddress(){
        this("Home");
    }


  //**************************************************************************
  //** Constructor
  //**************************************************************************
  /** Creates a new instance of Address using a node from a 
   *  FindItemResponseMessage.
   */
    protected PhysicalAddress(org.w3c.dom.Node addressNode) {

        type = javaxt.xml.DOM.getAttributeValue(addressNode, "Key");

        org.w3c.dom.NodeList childNodes = addressNode.getChildNodes();
        for (int j=0; j<childNodes.getLength(); j++){
            org.w3c.dom.Node childNode = childNodes.item(j);
            if (childNode.getNodeType()==1){
                String childNodeName = childNode.getNodeName();
                if (childNodeName.contains(":")) childNodeName = childNodeName.substring(childNodeName.indexOf(":")+1);
                String value = javaxt.xml.DOM.getNodeValue(childNode);
                if (value.length()>0){
                    if (childNodeName.equalsIgnoreCase("Street")){
                        street.add(value);
                    }
                    else if (childNodeName.equalsIgnoreCase("City")){
                        city = value;
                    }
                    else if (childNodeName.equalsIgnoreCase("State")){
                        state = value;
                    }
                    else if (childNodeName.equalsIgnoreCase("CountryOrRegion")){
                        country = value;
                    }
                    else if (childNodeName.equalsIgnoreCase("PostalCode")){
                        postalCode = value;
                    }
                }
            }
        }        
    }


  //**************************************************************************
  //** setType
  //**************************************************************************
  /** Used to set the type or category of address. Note that this field is
   *  required.
   * @param type Options include "Home", "Business", and "Other"
   */
    public void setType(String type) {
        if (type==null) type = "";
        if (type.toUpperCase().contains("HOME")) type = "Home";
        else if(type.toUpperCase().contains("BUSINESS") ||
                type.toUpperCase().contains("COMPANY")) type = "Business";
        else type = "Other";
        this.type = type;
    }

    public String getType(){
        return type;
    }


    public String[] getStreet(){
        if (street.size()==0) return null;
        else return street.toArray(new String[street.size()]);
    }

    public void addStreet(String street){
        for (String s : street.trim().split("\n")){
            s = s.trim();
            if (s.length()>0) this.street.add(s);
        }
    }
    
    public void setStreet(String street){
        this.street.clear();
        this.addStreet(street);
    }
    
    
    public String getCity(){
        return city;        
    }
    
    public void setCity(String city){
        this.city = city;
    }
    
    public String getState(){
        return state;        
    }
    
    public void setState(String state){
        this.state = state;
    }

    public String getCountry(){
        return country;
    }

    public void setCountry(String country){
        this.country = country;
    }

    public String getPostalCode(){
        return postalCode;
    }

    public void setPostalCode(String postalCode){
        this.postalCode = postalCode;
    }



  //**************************************************************************
  //** toXML
  //**************************************************************************
  /** Creates an xml fragment that can be used to save/update a contact via
   *  exchange web services (ews). <br/>
   *  http://msdn.microsoft.com/en-us/library/aa563318%28v=exchg.140%29.aspx
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
        java.util.Iterator<String> it = street.iterator();
        str.append("<" + namespace + "Entry Key=\"" + type + "\">");
        while (it.hasNext()){
            str.append("<" + namespace + "Street>" + it.next() + "</" + namespace + "Street>");
        }
        if (city!=null) str.append("<" + namespace + "City>" + city + "</" + namespace + "City>");
        //else str.append("<City/>");
        
        if (state!=null) str.append("<" + namespace + "State>" + state + "</" + namespace + "State>");
        //else str.append("<State/>");     
        
        if (country!=null) str.append("<" + namespace + "Country>" + country + "</" + namespace + "Country>");
        //else str.append("<Country/>");
        
        if (postalCode!=null) str.append("<" + namespace + "PostalCode>" + postalCode + "</" + namespace + "PostalCode>");
        //else str.append("<PostalCode/>");
        
        str.append("</" + namespace + "Entry>");
        
        return str.toString().trim();
    }



  //**************************************************************************
  //** toString
  //**************************************************************************
  /** Returns a human readable representation of this address. */

    public String toString(){
        StringBuffer str = new StringBuffer();
        java.util.Iterator<String> it = street.iterator();
        while (it.hasNext()){
            str.append(it.next() + "\r\n");
        }
        if (city!=null) str.append(city);
        if (state!=null) str.append(", " + state);
        if (postalCode!=null) str.append(" " + postalCode);
        return str.toString().trim();
    }
}