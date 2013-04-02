package javaxt.exchange;

//******************************************************************************
//**  PhoneNumber Class
//******************************************************************************
/**
 *   Used to represent a phone number.
 *
 ******************************************************************************/

public class PhoneNumber {

    private String number;
    private String type;


  /** Static variable to support the getPhoneNumber method. */
    private final static String[] numbers = new String[]{"0","1","2","3","4","5","6","7","8","9"};

    private final static String[] types = new String[]{

      //Personal Phone
        "HomePhone","MobilePhone","HomeFax","HomePhone2",

      //Work Phone
        "BusinessPhone","BusinessPhone2","BusinessFax",
        "CompanyMainPhone","AssistantPhone",

      //Primary/Callback
        "PrimaryPhone","Callback",

      //Other
        "OtherFax","OtherTelephone",

      //Legacy
        "CarPhone","Pager","Isdn","RadioPhone","Telex","TtyTddPhone"
    };


  //**************************************************************************
  //** Constructor
  //**************************************************************************
  /** Creates a new instance of PhoneNumber. Throws an exception if the number
   *  is invalid.
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
    public PhoneNumber(String phoneNumber, String type) throws ExchangeException {
        this.number = getNumber(phoneNumber);
        this.type = getType(type);

        if (number==null) throw new ExchangeException("Invalid Phone Number:  " + phoneNumber);
    }

    
    protected PhoneNumber(org.w3c.dom.Node childNode) throws ExchangeException {
        if (childNode.getNodeType()==1){
            String childNodeName = childNode.getNodeName();
            if (childNodeName.contains(":")) childNodeName = childNodeName.substring(childNodeName.indexOf(":")+1);
            if (childNodeName.equalsIgnoreCase("Entry")){
                this.number = getNumber(javaxt.xml.DOM.getNodeValue(childNode));
                this.type = javaxt.xml.DOM.getAttributeValue(childNode, "Key");


            }
        }

        if (number==null) {
            throw new ExchangeException("Invalid Phone Number:  "
                + javaxt.xml.DOM.getNodeValue(childNode));
        }
    }


    public String getNumber(){
        return number;
    }


  //**************************************************************************
  //** getType
  //**************************************************************************
  /** Returns the type of phone number. Options include:
   * <ul>
   * <li>AssistantPhone</li><li>BusinessFax</li><li>BusinessPhone</li>
   * <li>BusinessPhone2</li><li>Callback</li><li>CarPhone</li>
   * <li>CompanyMainPhone</li><li>HomeFax</li><li>HomePhone</li>
   * <li>HomePhone2</li><li>Isdn</li><li>MobilePhone</li><li>OtherFax</li>
   * <li>OtherTelephone</li><li>Pager</li><li>PrimaryPhone</li>
   * <li>RadioPhone</li><li>Telex</li><li>TtyTddPhone</li>
   * </ul>
   */
    public String getType(){
        this.hashCode();
        return type;
    }


  //**************************************************************************
  //** getTypes
  //**************************************************************************
  /** Returns a static list of phone number types:
   * <ul>
   * <li>AssistantPhone</li><li>BusinessFax</li><li>BusinessPhone</li>
   * <li>BusinessPhone2</li><li>Callback</li><li>CarPhone</li>
   * <li>CompanyMainPhone</li><li>HomeFax</li><li>HomePhone</li>
   * <li>HomePhone2</li><li>Isdn</li><li>MobilePhone</li><li>OtherFax</li>
   * <li>OtherTelephone</li><li>Pager</li><li>PrimaryPhone</li>
   * <li>RadioPhone</li><li>Telex</li><li>TtyTddPhone</li>
   * </ul>
   */
    public static String[] getTypes(){
        return types;
    }

    public String toString(){
        return type + ":  " + number;
    }

    public int hashCode(){
        return toString().hashCode();
    }

    public boolean equals(Object obj){
        if (obj instanceof PhoneNumber){
            if (obj!=null) return obj.toString().equalsIgnoreCase(this.toString());
        }
        return false;
    }

    
  //**************************************************************************
  //** getNumber
  //**************************************************************************
  /** Used to normalize a phone number into a common format. For example, if
   *  you pass in "(555) 555-5555" or "555.555.5555" or "5555555555", the
   *  method will return "555-555-5555". Returns null if there are less than
   *  10 digits found in the phone number. Removes "1" country code prefix.
   *  Note that this method is designed to work with US phone numbers.
   *  International phone numbers require a different solution.
   */
    private String getNumber(String phoneNumber){
        if (phoneNumber==null) return null;
        StringBuffer str = new StringBuffer();
        for (int i=0; i<phoneNumber.length(); i++){
            for (String number : numbers){
                if (phoneNumber.substring(i, i+1).equals(number)){
                    str.append(number);
                    break;
                }
            }
        }

        if (str.length()<10) return null;
        else{

            String prefix = "";
            if (str.length()>10){
                prefix = str.substring(0, str.length()-10);
                if (prefix.equals("1")) prefix = "";
                else prefix += "-";
            }

            return prefix + str.substring(str.length()-10, str.length()-7) +
                    "-" + str.substring(str.length()-7, str.length()-4) + "-" +
                    str.substring(str.length()-4);
        }
    }


  //**************************************************************************
  //** getType
  //**************************************************************************
  /** Returns a phone number type that most closely matches the given string.
   */
    private String getType(String type){
        type = type.toUpperCase().replace(" ", "");
        for (String t : types){
            if (t.toUpperCase().equals(type)) return t;
        }


        if (type.contains("HOME")) return "HomePhone";
        if (type.contains("MOBILE")) return "MobilePhone";
        
        if (type.equals("FAX") || type.equals("COMPANYFAX") || type.equals("WORKFAX")) return "BusinessFax";
        if (type.equals("PERSONALFAX")) return "HomeFax";
        if (type.contains("FAX")) return "OtherFax";

        if (type.contains("BUSINESS") || type.contains("WORK") || type.contains("OFFICE")) return "BusinessPhone";
        if (type.contains("ASSISTANT")) return "AssistantPhone";


        if (type.contains("PRIMARY")) return "PrimaryPhone";
        if (type.contains("CALLBACK")) return "Callback";
        

        if (type.contains("CAR")) return "CarPhone";
        if (type.contains("RADIO")) return "RadioPhone";
        if (type.contains("TTY") || type.contains("TDD")) return "TtyTddPhone";

        return "OtherTelephone";
    }
}