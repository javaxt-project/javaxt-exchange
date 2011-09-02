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
   *  @param type Options include "Home", "Business", and "Other". Accepts
   *  minor varients including "Home Address", "Work", etc.
   */
    public void setType(String type) {
        if (type==null) type = "";
        type = type.toUpperCase();

        if (type.contains("HOME") || type.contains("PERSONAL")) type = "Home";
        else if(type.contains("BUSINESS") || type.contains("COMPANY") ||
                type.contains("WORK")) type = "Business";
        else type = "Other";

        this.type = type;
    }


  //**************************************************************************
  //** getType
  //**************************************************************************
  /** Returns the type or category of address. Options include "Home",
   *  "Business", and "Other".
   */
    public String getType(){
        return type;
    }


  //**************************************************************************
  //** getStreet
  //**************************************************************************
  /** Returns the street address as an array.
   */
    public String[] getStreet(){
        if (street.size()==0) return null;
        else return street.toArray(new String[street.size()]);
    }


  //**************************************************************************
  //** addStreet
  //**************************************************************************
  /** Used to add a street address.
   */
    public void addStreet(String street){
        if (street==null) return;
        for (String s : street.trim().split("\n")){
            s = s.trim();
            if (s.length()>0) this.street.add(s);
        }
    }
    
    
  //**************************************************************************
  //** setStreet
  //**************************************************************************
  /** Used to set the street address. Multiple lines can be specified using a
   *  "\n" or "\r\n" delimitor.
   */
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
  /** Returns an xml fragment used to save or update a contact via Exchange
   *  Web Services (EWS):<br/>
   *  http://msdn.microsoft.com/en-us/library/aa563318%28v=exchg.140%29.aspx
   *
   *  @param namespace The namespace assigned to the "type". Typically this is
   *  "t" which corresponds to
   *  "http://schemas.microsoft.com/exchange/services/2006/types".
   *  Use a null value is you do not wish to append a namespace.
   *
   *  @param insert Boolean used to specify whether to return an xml
   *  formatted for inserts or updates. If true, the xml will be formatted for
   *  inserts. If false, xml will be formatted for updates.
   */
    protected String toXML(String namespace, boolean insert){

      //Update namespace prefix
        if (namespace!=null){
            if (!namespace.endsWith(":")) namespace+=":";
        }
        else{
            namespace = "";
        }

        
        StringBuffer str = new StringBuffer();

        if (insert){
            str.append("<" + namespace + "Entry Key=\"" + type + "\">");
            
            String street = getStreets();
            if (street!=null) str.append("<" + namespace + "Street>" + street + "</" + namespace + "Street>");

            if (city!=null) str.append("<" + namespace + "City>" + city + "</" + namespace + "City>");

            if (state!=null) str.append("<" + namespace + "State>" + state + "</" + namespace + "State>");

            if (country!=null) str.append("<" + namespace + "Country>" + country + "</" + namespace + "Country>");

            if (postalCode!=null) str.append("<" + namespace + "PostalCode>" + postalCode + "</" + namespace + "PostalCode>");

            str.append("</" + namespace + "Entry>");
        }
        else{

            String streets = getStreets();
            if (streets!=null) str.append(getUpdateXML("Street", streets, namespace));
            else str.append(getDeleteXML("Street", namespace));

            if (city!=null) str.append(getUpdateXML("City", city, namespace));
            else str.append(getDeleteXML("City", namespace));

            if (state!=null) str.append(getUpdateXML("State", state, namespace));
            else str.append(getDeleteXML("State", namespace));

            if (country!=null) str.append(getUpdateXML("CountryOrRegion", country, namespace));
            else str.append(getDeleteXML("CountryOrRegion", namespace));

            if (postalCode!=null) str.append(getUpdateXML("PostalCode", postalCode, namespace));
            else str.append(getDeleteXML("PostalCode", namespace));

        }
        
        return str.toString().trim();
    }



    private String getUpdateXML(String key, String value, String namespace){
        StringBuffer str = new StringBuffer();
        str.append("<" + namespace + "SetItemField>");
        str.append("<" + namespace + "IndexedFieldURI FieldURI=\"contacts:PhysicalAddress:" + key + "\" FieldIndex=\"" + type + "\"/>");
        str.append("<" + namespace + "Contact>");
        str.append("<" + namespace + "PhysicalAddresses>");
        str.append("<" + namespace + "Entry Key=\"" + type + "\">");
        str.append("<" + namespace + key + ">" + value + "</" + namespace + key + ">");
        str.append("</" + namespace + "Entry>");
        str.append("</" + namespace + "PhysicalAddresses>");
        str.append("</" + namespace + "Contact>");
        str.append("</" + namespace + "SetItemField>");
        return str.toString();
    }

    private String getDeleteXML(String key, String namespace){
        StringBuffer str = new StringBuffer();
        str.append("<" + namespace + "DeleteItemField>");
        str.append("<" + namespace + "IndexedFieldURI FieldURI=\"contacts:PhysicalAddress:" + key + "\" FieldIndex=\"" + type + "\"/>");
        str.append("</" + namespace + "DeleteItemField>");
        return str.toString();
    }


  //**************************************************************************
  //** getStreets
  //**************************************************************************
  /** Returns a properly formatted street address for an XML/SOAP message. */

    private String getStreets(){

        String streets = "";
        java.util.Iterator<String> it = street.iterator();
        while (it.hasNext()){
            String street = it.next();
            if (street!=null) street = street.trim();
            if (street!=null && street.length()>0){
                streets += street + "\r\n";
            }
        }
        streets = streets.trim();
        if (streets.length()>0){

          //Here's how to preserve line breaks:
          //http://blogs.msdn.com/b/pcreehan/archive/2009/07/22/line-breaks-in-managed-web-service-proxy-classes.aspx
            streets = streets.replace("\r", "&#x0d;");
            streets = streets.replace("\n", "&#x0a;");
            return streets;
        }
        else{
            return null;
        }

    }


  //**************************************************************************
  //** toString
  //**************************************************************************
  /** Returns a human readable representation of this address. */

    public String toString(){
        StringBuffer str = new StringBuffer();
        java.util.Iterator<String> it = street.iterator();
        while (it.hasNext()){
            String street = it.next();
            if (street!=null) street = street.trim();
            if (street!=null && street.length()>0){
                str.append(street + "\r\n");
            }
        }
        if (city!=null) str.append(city);
        if (state!=null) str.append(", " + state);
        if (postalCode!=null) str.append(" " + postalCode);
        return str.toString().trim();
    }



    public boolean equals(Object obj){

        if (obj!=null){
            String str = obj.toString().trim().replace("\r", "").replace("\n", "");
            String str2 = this.toString().replace("\r", "").replace("\n", "");
            /*
            for (int i=0; i<str.length(); i++){
                System.out.print(str.charAt(i) + " - " + str2.charAt(i));
                if (str.charAt(i) != str2.charAt(i)) System.out.print(" <---");
                System.out.print("\r\n");
            }
            */
            return str.equalsIgnoreCase(str2);
        }
        else return false;
    }
}