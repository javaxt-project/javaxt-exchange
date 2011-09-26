package javaxt.exchange;

//******************************************************************************
//**  EmailAddress Class
//******************************************************************************
/**
 *   Enter class description here
 *
 ******************************************************************************/

public class EmailAddress {

    private String emailAddress;


  //**************************************************************************
  //** Constructor
  //**************************************************************************
  /** Creates a new instance of EmailAddress. */

    public EmailAddress(String email) throws ExchangeException {
        String emailAddress = email;
        if (emailAddress!=null){
            emailAddress = emailAddress.toLowerCase().trim();
            if (!emailAddress.contains("@") || emailAddress.length()==0) emailAddress = null; //Need a stronger validation than this...
        }
        if (emailAddress==null) throw new ExchangeException("Invalid Email Address:  " + email);
        this.emailAddress = emailAddress;
    }

    public String toString(){
        return emailAddress;
    }

    public int hashCode(){
        return emailAddress.hashCode();
    }

    public boolean equals(Object obj){
        if (obj!=null){
            if (obj instanceof EmailAddress){
                return obj.toString().equalsIgnoreCase(emailAddress);
            }
            else if (obj instanceof String){
                try{
                    return new EmailAddress((String) obj).toString().equalsIgnoreCase(emailAddress);
                }
                catch(Exception e){
                }
            }
        }
        return false;
    }

}