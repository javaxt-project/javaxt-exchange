package javaxt.exchange;
import java.util.regex.Pattern;

//******************************************************************************
//**  EmailAddress Class
//******************************************************************************
/**
 *   Used to represent an email address. 
 *
 ******************************************************************************/

public class EmailAddress {


  /** Email validation from mkyong:
   *  http://www.mkyong.com/regular-expressions/how-to-validate-email-address-with-regular-expression/
   */
    private static final Pattern pattern = Pattern.compile(
  //"^[_A-Za-z0-9-]+(\\.[_A-Za-z0-9-]+)*@[A-Za-z0-9]+(\\.[A-Za-z0-9]+)*(\\.[A-Za-z]{2,})$"
    "^[_A-Za-z0-9-+]+(\\.[_A-Za-z0-9-+]+)*@[A-Za-z0-9-]+(\\.[A-Za-z0-9-]+)*(\\.[A-Za-z]{2,})$"
    );

    private String emailAddress;

  //**************************************************************************
  //** Constructor
  //**************************************************************************
  /** Creates a new instance of EmailAddress. */

    public EmailAddress(String email) throws ExchangeException {
        String emailAddress = email;
        if (emailAddress!=null){
            emailAddress = emailAddress.toLowerCase().trim();
            if (!pattern.matcher(emailAddress).matches()) emailAddress = null;
        }
        
        if (emailAddress==null){
            throw new ExchangeException("Invalid Email Address:  " + email); //<--DO NOT CHANGE THIS ERROR MESSAGE! Check code for usage...
        }
        
        this.emailAddress = emailAddress;
    }


  //**************************************************************************
  //** toString
  //**************************************************************************
  /** Returns the email address used to instantiate this class. Note that the
   *  email address is trimmed and converted to lower case in the constructor.
   */
    public String toString(){
        return emailAddress;
    }

    
    public int hashCode(){
        return emailAddress.hashCode();
    }

  //**************************************************************************
  //** equals
  //**************************************************************************
  /** Performs case insensitive string comparison between two email addresses.
   *  @param obj String or EmailAddress
   */
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