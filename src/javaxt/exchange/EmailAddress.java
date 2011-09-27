package javaxt.exchange;
import java.util.regex.Pattern;

//******************************************************************************
//**  EmailAddress Class
//******************************************************************************
/**
 *   Used to represent an email address. Includes email validation from mkyong:
 *   http://www.mkyong.com/regular-expressions/how-to-validate-email-address-with-regular-expression/
 *
 ******************************************************************************/

public class EmailAddress {

    private String emailAddress;
    private static final Pattern pattern = Pattern.compile(
  //"^[_A-Za-z0-9-]+(\\.[_A-Za-z0-9-]+)*@[A-Za-z0-9]+(\\.[A-Za-z0-9]+)*(\\.[A-Za-z]{2,})$"
    "^[_A-Za-z0-9-+]+(\\.[_A-Za-z0-9-+]+)*@[A-Za-z0-9-]+(\\.[A-Za-z0-9-]+)*(\\.[A-Za-z]{2,})$"
    );

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
            throw new ExchangeException("Invalid Email Address:  " + email);
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