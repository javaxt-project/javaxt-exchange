package javaxt.exchange;

//******************************************************************************
//**  ExtendedFieldURI
//******************************************************************************
/**
 *   Used to represent an ExtendedFieldURI.
 *
 ******************************************************************************/

public class ExtendedFieldURI extends FieldURI {

    private String id;
    private String name;
    private String type = "String";

  //**************************************************************************
  //** Constructor
  //**************************************************************************
  /** Creates a new instance of this class.
   *  @param id A unique id in the form of a Microsoft GUID.
   *  @param name Property name.
   *  @param name Property type (e.g. "String", "Integer", "SystemTime", etc).
   */
    public ExtendedFieldURI(String id, String name, String type){
        this.id = id;
        this.name = name;
        this.type = type;
    }


  //**************************************************************************
  //** Constructor
  //**************************************************************************
  /** Creates a new instance of this class.
   *  @param node "ExtendedProperty" node
   */
    protected static Object[] parse(org.w3c.dom.Node node){

        String id = null;
        String name = null;
        String type = null;
        String value = null;

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

      //TODO: Convert the value to a proper type before instantiating the Value class
        javaxt.utils.Value val = new javaxt.utils.Value(value);
        

        return new Object[]{new ExtendedFieldURI(id, name, type), val};
    }


  //**************************************************************************
  //** getID
  //**************************************************************************
  /** Returns the guid associated with this property.
   */
    public String getID(){
        return id;
    }


  //**************************************************************************
  //** getName
  //**************************************************************************
  /** Returns the name of this property.
   */
    public String getName(){
        return name;
    }


  //**************************************************************************
  //** getType
  //**************************************************************************
  /** Returns the type of property (e.g. "String").
   */
    public String getType(){
        return type;
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
   *  formatted for inserts, updates or deletes. Valid options include
   *  "create", "update", "delete".
   */
    protected String toXML(String namespace, String operation, javaxt.utils.Value value){

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

    
    public String toXML(String namespace){
        String idAttr = (id==null ? "" : "PropertySetId=\"" + id + "\"");
        String nameAttr = (name.startsWith("0x") ? "PropertyTag=\"" + name + "\"" : "PropertyName=\"" + name + "\"");
        String typeAttr = "PropertyType=\"" + type + "\"";
        return "<" + namespace + ":ExtendedFieldURI " + nameAttr + " " + idAttr + " " + typeAttr + "/>";
    }

    
    public String toString(){
        return name;
    }

    public int hashCode(){ 
        return name.startsWith("0x") ? name.hashCode() : id.toUpperCase().hashCode();
    }

    public boolean equals(Object obj){
        if (obj!=null){
            if (obj instanceof ExtendedFieldURI){
                ExtendedFieldURI property = (ExtendedFieldURI) obj;
                return (property.id.equalsIgnoreCase(this.id));
            }
        }
        return false;
    }
}