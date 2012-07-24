package javaxt.exchange;

//******************************************************************************
//**  ExtendedProperty Class
//******************************************************************************
/**
 *   Used to represent a custom, extended property associated with a folder
 *   item.
 *
 ******************************************************************************/

public class ExtendedProperty {

    private String id;
    private String name;
    private String type = "String";
    private String value;


  //**************************************************************************
  //** Constructor
  //**************************************************************************
  /** Creates a new instance of this class.
   *  @param id A unique id in the form of a Microsoft GUID
   */
    public ExtendedProperty(String id, String name, String value){
        this.id = id;
        this.name = name;
        this.value = value;
    }


    
  //**************************************************************************
  //** Constructor
  //**************************************************************************
  /** Creates a new instance of this class.
   *  @param node "ExtendedProperty" node
   */
    public ExtendedProperty(org.w3c.dom.Node node){

        org.w3c.dom.NodeList childNodes = node.getChildNodes();
        for (int j=0; j<childNodes.getLength(); j++){
            org.w3c.dom.Node childNode = childNodes.item(j);
            if (childNode.getNodeType()==1){

                String childNodeName = childNode.getNodeName();
                if (childNodeName.contains(":")) childNodeName = childNodeName.substring(childNodeName.indexOf(":")+1);

                if (childNodeName.equalsIgnoreCase("ExtendedFieldURI")){
                    name = javaxt.xml.DOM.getAttributeValue(childNode, "PropertyName");
                    type = javaxt.xml.DOM.getAttributeValue(childNode, "PropertyType");
                    id = javaxt.xml.DOM.getAttributeValue(childNode, "PropertySetId");

                    String PropertyTag = javaxt.xml.DOM.getAttributeValue(childNode, "PropertyTag");
                    if (PropertyTag.length()>0) name = PropertyTag;
                }
                else if (childNodeName.equalsIgnoreCase("Value")){
                    value = javaxt.xml.DOM.getNodeValue(childNode);
                }
            }
        }
    }

  /** Returns the guid associated with this property. */
    public String getID(){
        return id;
    }

    public String getName(){
        return name;
    }

    public String getType(){
        return type;
    }

    public String getValue(){
        return value;
    }

    public void setValue(String value){
        this.value = value;
    }


  //**************************************************************************
  //** toXML
  //**************************************************************************
  /** Returns an xml fragment used to save or update an ExtendedProperty via
   *  Exchange Web Services (EWS):<br/>
   *  http://msdn.microsoft.com/en-us/library/exchange/dd633654%28v=exchg.80%29
   *
   *  @param namespace The namespace assigned to the "type". Typically this is
   *  "t" which corresponds to
   *  "http://schemas.microsoft.com/exchange/services/2006/types".
   *  Use a null value is you do not wish to append a namespace.
   *
   *  @param operation String used to specify whether to return an xml
   *  formatted for inserts, updates or deletes. Valid
   */
    protected String toXML(String namespace, String operation){

      //Update namespace prefix
        if (namespace!=null){
            if (!namespace.endsWith(":")) namespace+=":";
        }
        else{
            namespace = "";
        }

        StringBuffer xml = new StringBuffer();
        if (operation.equalsIgnoreCase("create")){            
            xml.append("<" + namespace + "ExtendedProperty>");
            xml.append("<" + namespace + "ExtendedFieldURI PropertySetId=\"" + id + "\" PropertyName=\"" + name + "\" PropertyType=\"" + type + "\" />");
            xml.append("<" + namespace + "Value>" + value + "</" + namespace + "Value>");
            xml.append("</" + namespace + "ExtendedProperty>");
        }
        else if(operation.equalsIgnoreCase("update")){
            xml.append("<" + namespace + "SetItemField>");
            xml.append("<" + namespace + "ExtendedFieldURI PropertySetId=\"" + id + "\" PropertyName=\"" + name + "\" PropertyType=\"" + type + "\" />");
            xml.append("<" + namespace + "Message>");
            xml.append("<" + namespace + "ExtendedProperty>");
            xml.append("<" + namespace + "ExtendedFieldURI PropertySetId=\"" + id + "\" PropertyName=\"" + name + "\" PropertyType=\"" + type + "\" />");
            xml.append("<" + namespace + "Value>" + value + "</" + namespace + "Value>");
            xml.append("</" + namespace + "ExtendedProperty>");
            xml.append("</" + namespace + "Message>");
            xml.append("</" + namespace + "SetItemField>");
        }
        else if(operation.equalsIgnoreCase("delete")){
            xml.append("<" + namespace + "DeleteItemField>");
            xml.append("<" + namespace + "ExtendedFieldURI PropertySetId=\"" + id + "\" PropertyName=\"" + name + "\" PropertyType=\"" + type + "\" />");
            xml.append("</" + namespace + "DeleteItemField>");
        }

        return xml.toString();
    }

    
    public String toString(){
        return name + ": " + value;
    }

    public int hashCode(){
        return name.toUpperCase().hashCode();
    }

    public boolean equals(Object obj){
        if (obj instanceof ExtendedProperty){
            ExtendedProperty property = (ExtendedProperty) obj;
            return (property.id.equalsIgnoreCase(this.id) && property.value.equals(this.value));
        }
        return false;
    }
}