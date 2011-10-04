package javaxt.exchange;

//******************************************************************************
//**  ServerVersion Class
//******************************************************************************
/**
 *   Used to parse ServerVersionInfo from the SOAP header. Source:
 *   http://support.microsoft.com/kb/158530
 *
 ******************************************************************************/

public class ServerVersion {
    
    private static java.util.HashMap<String, String> versions = getVersions();
    private String versionNumber;
    private String versionName;

  //**************************************************************************
  //** Constructor
  //**************************************************************************
  /** Creates a new instance of ServerVersion. */

    public ServerVersion(org.w3c.dom.Document xml) {
        org.w3c.dom.Node[] nodes = javaxt.xml.DOM.getElementsByTagName("ServerVersionInfo", xml);
        if (nodes.length>0){
            org.w3c.dom.NamedNodeMap attr = nodes[0].getAttributes();
            String MajorVersion = attr.getNamedItem("MajorVersion").getTextContent();
            String MinorVersion = attr.getNamedItem("MinorVersion").getTextContent();

            java.util.Iterator<String> it = versions.keySet().iterator();
            while (it.hasNext()){
                String versionNumber = it.next();
                String versionName = versions.get(versionNumber);
                if (versionNumber.startsWith(MajorVersion + "." + MinorVersion + ".")){
                    this.versionNumber = versionNumber;
                    this.versionName = versionName;
                    return;
                }
            }
        }
    }

    public String getBuildNumber(){
        return versionNumber;
    }

    public String getName(){
        return versionName;
    }

    public String toString(){
        return versionName + " (" + versionNumber + ")";
    }

    private static java.util.HashMap<String, String> getVersions(){
        java.util.HashMap<String, String> versions = new java.util.HashMap<String, String>();
        versions.put("8.1.0240.006","Microsoft Exchange Server 2007 SP1");
        versions.put("8.2.0176.002","Microsoft Exchange Server 2007 SP2");
        versions.put("8.3.0083.006","Microsoft Exchange Server 2007 SP3");
        versions.put("14.00.0639.021","Microsoft Exchange Server 2010");
        versions.put("14.01.0218.015","Microsoft Exchange Server 2010 SP1");
        return versions;
    }

}