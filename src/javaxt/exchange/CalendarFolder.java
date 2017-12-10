package javaxt.exchange;

//******************************************************************************
//**  CalendarFolder Class
//******************************************************************************
/**
 *   Enter class description here
 *
 ******************************************************************************/

public class CalendarFolder extends Folder {

    private java.util.HashSet<FieldURI> props = new java.util.HashSet<FieldURI>();
    private Connection conn;

  //**************************************************************************
  //** Constructor
  //**************************************************************************
  /** Creates a new instance of CalendarFolder. */

    public CalendarFolder(Connection conn) throws ExchangeException {
        super("calendar", conn);
        props.add(new FieldURI("calendar:TimeZone"));
        this.conn = conn;
    }

    /*
    public CalendarEvent[] getEvents() throws ExchangeException {

        java.util.ArrayList<CalendarEvent> events = new java.util.ArrayList<CalendarEvent>();

        int offset = 0;
        int maxRecords = 25;
        while(true){

            CalendarEvent[] arr = getEvents(offset, maxRecords);
            for (int i=0; i<arr.length; i++){
                events.add(arr[i]);
            }
            if (arr.length<maxRecords) break;
            else offset+=maxRecords;
        }

        return events.toArray(new CalendarEvent[events.size()]);
    }
    */


  //**************************************************************************
  //** getEvents
  //**************************************************************************
  /** Returns an array of CalendarEvents for a given range.
   *  @param limit Maximum number of items to return.
   *  @param offset Item offset. 0 implies no offset.
   */
    public CalendarEvent[] getEvents(int offset, int limit) throws ExchangeException {
        java.util.ArrayList<CalendarEvent> events = new java.util.ArrayList<CalendarEvent>();
        org.w3c.dom.Document xml = getItems(offset, limit, props, null, null);
        org.w3c.dom.Node[] nodes = javaxt.xml.DOM.getElementsByTagName("CalendarItem", xml);

        for (org.w3c.dom.Node node : nodes){
            events.add(new CalendarEvent(node, conn));
        }
        return events.toArray(new CalendarEvent[events.size()]);
    }


  //**************************************************************************
  //** getEvents
  //**************************************************************************
  /** Returns an array of CalendarEvents for a given day.
   */
    public CalendarEvent[] getEvents(javaxt.utils.Date date) throws ExchangeException {
        return getEvents(date.getDate());
    }


  //**************************************************************************
  //** getEvents
  //**************************************************************************
  /** Returns an array of CalendarEvents for a given day.
   */
    public CalendarEvent[] getEvents(java.util.Date date) throws ExchangeException {
        try{
            javaxt.utils.Date startDate = new javaxt.utils.Date(date);
            startDate = new javaxt.utils.Date(startDate.toString("yyyy-MM-dd"));

            javaxt.utils.Date nextDay = new javaxt.utils.Date(date);
            nextDay.add(1, "day");
            nextDay = new javaxt.utils.Date(nextDay.toString("yyyy-MM-dd"));


            return getEvents(startDate, nextDay);
        }
        catch(java.text.ParseException e){
          //Should never happen!!
            return null;
        }        
    }


  //**************************************************************************
  //** getEvents
  //**************************************************************************
  /** Returns an array of CalendarEvents for a given date range.
   */
    public CalendarEvent[] getEvents(java.util.Date start, java.util.Date end) throws ExchangeException {
        return getEvents(new javaxt.utils.Date(start), new javaxt.utils.Date(end));
    }


  //**************************************************************************
  //** getEvents
  //**************************************************************************
  /** Returns an array of CalendarEvents for a given date range.
   */
    public CalendarEvent[] getEvents(javaxt.utils.Date start, javaxt.utils.Date end) throws ExchangeException {

        String StartDate = FolderItem.formatDate(start);
        String EndDate = FolderItem.formatDate(end);

        java.util.ArrayList<CalendarEvent> events = new java.util.ArrayList<CalendarEvent>();
        org.w3c.dom.Document xml = getItems("<m:CalendarView StartDate=\"" + StartDate + "\" EndDate=\"" + EndDate + "\"/>", props, null, null);
        org.w3c.dom.Node[] nodes = javaxt.xml.DOM.getElementsByTagName("CalendarItem", xml);

        for (org.w3c.dom.Node node : nodes){
            CalendarEvent event = new CalendarEvent(node, conn);
            if (event.isAllDayEvent() && event.getStartTime().compareTo(start, "days")==-1){
                
            }
            else{
                events.add(new CalendarEvent(node, conn));
            }
        }



        return events.toArray(new CalendarEvent[events.size()]);
    }
    

    
/*
    public void sendSharingRequest(Mailbox recipient, Mailbox sender, Connection conn) throws ExchangeException {
        
        if (true) throw new ExchangeException("not implemented");
        
      
        javaxt.exchange.Email email = new javaxt.exchange.Email();
        email.setSubject("I'd like to share my calendar with you");
        email.setBody("<div style=\"color:blue;\"><span>Do you see this in blue?</div>", "HTML");
        email.addRecipient("To", recipient);
        email.setMessageClass("IPM.Sharing"); //"PidTagMessageClass", "IPM.Sharing"
        //email.setMessageType("Sharing"); //"PidNameContentClass", "Sharing"
        
        
      //Create attachment
        StringBuffer str = new StringBuffer();
        str.append("<?xml version=\"1.0\"?>");
        str.append("<SharingMessage xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" ");
        str.append("xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" ");
        str.append("xmlns=\"http://schemas.microsoft.com/sharing/2008\">");
        str.append("<DataType>calendar</DataType>");
        str.append("<Initiator>");
        str.append("<Name>" + sender.getName() + "</Name>");
        str.append("<SmtpAddress>" + sender.getEmailAddress() + "</SmtpAddress>");
//str.append("<EntryId>" + sender.getEntryID() + "</EntryId>");
        str.append("</Initiator>");
        str.append("<Invitation>");
        str.append("<Providers>");
        str.append("<Provider Type=\"ms-exchange-internal\" TargetRecipients=\"" + recipient.getEmailAddress() + "\">");
        str.append("<FolderId xmlns=\"http://schemas.microsoft.com/exchange/sharing/2008\">");
        str.append(this.getID());
        str.append("</FolderId>");
        str.append("<MailboxId xmlns=\"http://schemas.microsoft.com/exchange/sharing/2008\">");
        str.append(sender.getID());
        str.append("</MailboxId>");
        str.append("</Provider>");
        str.append("</Providers>");
        str.append("</Invitation>");
        str.append("</SharingMessage>");
        //email.addAttachment(str.toString());
        
        
        
        java.util.HashMap<String, String> headers = new java.util.HashMap<String, String>();

        

//        X-MS-Exchange-Organization-AuthAs: Internal 
//        X-MS-Exchange-Organization-AuthMechanism: 04 
//        X-MS-Exchange-Organization-AuthSource: exchange2010.example.com 
//        X-MS-Has-Attach: yes 
//        X-MS-Exchange-Organization-SCL: -1 
//        X-MS-TNEF-Correlator: 
//        x-sharing-capabilities: 402B0 
//        x-sharing-flavor: 310 
//        x-sharing-provider-guid: AEF0060000000000C000000000000046 
//        x-sharing-provider-name: Microsoft Exchange 
//        x-sharing-provider-url: http://www.microsoft.com/exchange/ 
//        x-sharing-remote-path: 
//        x-sharing-remote-name: =?iso-8859-1?Q?Jane=B4s_her_share_calendar?= 
//        x-sharing-remote-uid: 00000000BC5AB69D4370454FB480A15DB2A7E93C0100DCB29F06B087BB439B542470E7A5D6E300002DAFEAEA0000 
//        x-sharing-remote-store-uid: 0000000038A1BB1005E5101AA1BB08002B2A56C20000454D534D44422E444C4C00000000000000001B55FA20AA6611CD9BC800AA002FC45A0C00000045584348414E474532303130002F6F3D4669727374204F7267616E697A6174696F6E2F6F753D45786368616E67652041646D696E6973747261746976652047726F7570202846594449424F484632335350444C54292F636E3D526563697069656E74732F636E3D4A616E6520446F6534633900 
//        x-sharing-remote-type: IPF.Appointment        
            
        

        // Section 2.2.2
        headers.put("x-sharing-capabilities", "40290"); //Test representation of SharingCapabilities value 
        headers.put("x-sharing-flavor", "20310"); //Text representation of SharingFlavor value [MS-OXSHARE] 2.2.2.6 
        
        

        headers.put("x-sharing-local-type", "IPF.Appointment"); //MUST be set to same value as PidLidSharingLocalType 
        headers.put("x-sharing-provider-guid", "AEF0060000000000C000000000000046"); //Constant Required Value [MS-OXSHARE] 2.2.2.13 
        headers.put("x-sharing-provider-name", "Microsoft Exchange"); //Constant Required Value [MS-OXSHARE] 2.2.2.15] 
        headers.put("x-sharing-provider-url", "HTTP://www.microsoft.com/exchange"); //Constant Required Value [MS-OXSHARE] 2.2.2.17 

        // Section 2.2.3
        headers.put("x-sharing-remote-name", "Calendar"); //MUST be set to same value as PidLidSharingRemoteName 
        headers.put("x-sharing-remote-store-uid", sender.getID()); //MUST be set to same value as PidLidSharingRemoteStoreUid 
        headers.put("x-sharing-remote-type", "IPF.Appointment"); //Constant Required Value [MS-OXSHARE] 2.2.3.6 
        headers.put("x-sharing-remote-uid", this.getID()); //Must be set to same value as PidLidSharingRemoteUid 



       
       
       String action = "SendOnly";
       
      //Send the message and create a copy in the "sentitems" folder
        StringBuffer msg = new StringBuffer();
        msg.append("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
        msg.append("<soap:Envelope xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\" xmlns:m=\"http://schemas.microsoft.com/exchange/services/2006/messages\">");
        msg.append("<soap:Body>");
        msg.append("<m:CreateItem MessageDisposition=\"" + action + "\">");

        
        
        msg.append("<m:SavedItemFolderId>");

            String folderName = "drafts";
            if (action.equals("SendOnly") || action.equals("SendAndSaveCopy")) folderName = "sentitems";
            msg.append("<t:DistinguishedFolderId Id=\"" + folderName + "\" />");
        
        msg.append("</m:SavedItemFolderId>");
        msg.append("<m:Items>");
        
        


        String messageType = "Message";
        
        msg.append("<t:" + messageType + ">");
        msg.append("<t:ItemClass>" + email.getMessageClass() + "</t:ItemClass>"); //<--New for sharing requests
        msg.append("<t:Subject>" + email.getSubject() + "</t:Subject>");
        msg.append("<t:Sensitivity>" + email.getSensitivity() + "</t:Sensitivity>");

        
      //Set body
        if (email.getBody()!=null){
            msg.append("<t:Body BodyType=\"" + email.getBodyType() + "\">");
            //msg.append(wrap(body));
            msg.append("</t:Body>");
        };


        msg.append("<t:Importance>" + email.getImportance() + "</t:Importance>");


        msg.append("<t:ToRecipients>");
        msg.append(recipient.toXML("t"));
        msg.append("</t:ToRecipients>");


        msg.append("</t:" + messageType + ">");


        
        msg.append("</m:Items>");
        msg.append("</m:CreateItem>");
        msg.append("</soap:Body>");
        msg.append("</soap:Envelope>");        
        
        conn.execute(msg.toString(), headers);
        
    }
*/
}