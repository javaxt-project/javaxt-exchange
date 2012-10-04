package javaxt.exchange;

//******************************************************************************
//**  GUID Class
//******************************************************************************
/**
 *   Used to generate a random sequence of characters used to resemble a
 *   Microsoft GUID.
 *
 *   "[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}"
 *
 ******************************************************************************/

public class Guid {

    private final static String chars = "0123456789ABCDEF"; //GHIJKLMNOPQRSTUVWXYZ
    private String id;

    public Guid(){
       id =
       getRandomChars(8) + "-" +
       getRandomChars(4) + "-" +
       getRandomChars(4) + "-" +
       getRandomChars(4) + "-" +
       getRandomChars(12);

       //"c11ff724-aa03-4555-9952-8fa248a11c3e"
    }


    public String toString(){
        return id;
    }

    public int hashCode(){
        return id.hashCode();
    }

    public boolean equals(Object obj){
        if (obj!=null){
            if (obj instanceof String || obj instanceof Guid){
                return obj.toString().equalsIgnoreCase(id);
            }
        }
        return false;
    }

    private String getRandomChars(int numChars){
        StringBuffer str = new StringBuffer();
        for (int i = 1; i<=numChars; i++){
             int x = new java.util.Random().nextInt(chars.length());
             str.append(chars.substring(x,x+1));
        }
        return str.toString();
    }
}