package javaxt.exchange;

//******************************************************************************
//**  FieldOrder Class
//******************************************************************************
/**
 *   This class is used by the Folder.getItems() method to sort results.
 *
 ******************************************************************************/

public class FieldOrder {
    
    private FieldURI field;
    private boolean descending;


  //**************************************************************************
  //** Constructor
  //**************************************************************************
  /** Creates a new instance of this class using a FieldURI or 
   *  ExtendedFieldURI.
   *
   *  @param field A FieldURI or ExtendedFieldURI representing a field to sort
   *  by.
   *
   *  @param descending Indicates the sort direction. True indicates descending
   *  order. False indicates ascending order.
   */
    public FieldOrder(FieldURI field, boolean descending){
        this.field = field;
        this.descending = descending;
    }


  //**************************************************************************
  //** getField
  //**************************************************************************
  /** Returns the FieldURI or ExtendedFieldURI by which to sort.
   */
    public FieldURI getField(){
        return field;
    }


  //**************************************************************************
  //** getOrder
  //**************************************************************************
  /** Returns the direction of the sort (e.g. "Ascending" or "Descending").
   */
    public String getOrder(){
        return descending? "Descending" : "Ascending";
    }


  //**************************************************************************
  //** toXML
  //**************************************************************************
  /** Returns an XML fragment representing the field and order. This method is
   *  used by the
   *  http://msdn.microsoft.com/en-us/library/aa564968%28v=exchg.140%29.aspx
   */
    protected String toXML(){
        return
        "<t:FieldOrder Order=\"" + getOrder() + "\">" + this.field.toXML("t") +
        "</t:FieldOrder>";
    }

    public String toString(){
        return field.toString() + " " + getOrder();
    }


  //**************************************************************************
  //** hashCode
  //**************************************************************************
  /** Returns the field's hash code.
   */
    public int hashCode(){
        return field.hashCode();
    }


  //**************************************************************************
  //** equals
  //**************************************************************************
  /** Used to compare this FieldOrder to another. Returns true if the field
   *  and sort direction match.
   */
    public boolean equals(Object obj){
        if (obj!=null){
            if (obj instanceof FieldOrder){
                FieldOrder field = (FieldOrder) obj;
                return (field.hashCode()==this.hashCode() &&
                        field.descending==this.descending);
            }
        }
        return false;
    }
}