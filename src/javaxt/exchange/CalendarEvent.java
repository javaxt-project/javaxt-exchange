package javaxt.exchange;

//******************************************************************************
//**  CalendarEvent Class
//******************************************************************************
/**
 *   Used to represent a calendar entry.
 *
 ******************************************************************************/

public class CalendarEvent extends FolderItem {

    private String location;
    private javaxt.utils.Date startTime;
    private javaxt.utils.Date endTime;
    private Mailbox organizer;

  //**************************************************************************
  //** Constructor
  //**************************************************************************
  /** Creates a new instance of this class
   */
    public CalendarEvent(String exchangeID, Connection conn) throws ExchangeException{
        super(exchangeID, conn);
        parseCalendarItem();
    }

  //**************************************************************************
  //** Constructor
  //**************************************************************************
  /** Creates a new instance of CalendarEvent. */

    public CalendarEvent(org.w3c.dom.Node calendarItemNode) {
        super(calendarItemNode);
        parseCalendarItem();
    }


  //**************************************************************************
  //** parseContact
  //**************************************************************************
  /** Used to parse an xml node with event information.
   */
    private void parseCalendarItem(){
        org.w3c.dom.NodeList outerNodes = this.getChildNodes();
        for (int i=0; i<outerNodes.getLength(); i++){
            org.w3c.dom.Node outerNode = outerNodes.item(i);
            if (outerNode.getNodeType()==1){
                String nodeName = outerNode.getNodeName();
                if (nodeName.contains(":")) nodeName = nodeName.substring(nodeName.indexOf(":")+1);
                if (nodeName.equalsIgnoreCase("Location")){
                    location = javaxt.xml.DOM.getNodeValue(outerNode);
                }
                else if(nodeName.equalsIgnoreCase("Start")){
                    javaxt.utils.Date date = new javaxt.utils.Date(javaxt.xml.DOM.getNodeValue(outerNode));
                    if (!date.failedToParse()) startTime = date;
                }
                else if(nodeName.equalsIgnoreCase("End")){
                    javaxt.utils.Date date = new javaxt.utils.Date(javaxt.xml.DOM.getNodeValue(outerNode));
                    if (!date.failedToParse()) endTime = date;
                }
                else if(nodeName.equalsIgnoreCase("Organizer")){
                    org.w3c.dom.Node[] mailbox = javaxt.xml.DOM.getElementsByTagName("Mailbox", outerNode);
                    if (mailbox.length>0) organizer = new Mailbox(mailbox[0]);
                }
            }
        }
    }


    public String getSubject(){
        return super.getSubject();
    }

    public void setSubject(String subject){
        super.setSubject(subject);
    }

    public String getBody(){
        return super.getBody();
    }

    public void setBody(String body){
        super.setBody(body);
    }

    public Mailbox getOrganizer(){
        return organizer;
    }

    public void setOrganizer(Mailbox organizer){
        //this.organizer = organizer;
    }

    public javaxt.utils.Date getStartTime(){
        return startTime;
    }

    public javaxt.utils.Date getEndTime(){
        return endTime;
    }


  //**************************************************************************
  //** getDuration
  //**************************************************************************
  /** Returns the duration of the meeting/event
   *  @param units Units of measure (e.g. hours, minutes, seconds, weeks,
   *  months, years, etc.)
   */
    public long getDuration(String units){
        return endTime.compareTo(startTime, units);
    }


    public String getLocation(){
        location = getValue(location);
        return location;
    }

    public void setLocation(String location){
        location = getValue(location);

        if (id!=null) {
            if (location==null && this.location!=null) updates.put("Location", null);
            if (location!=null && !location.equals(this.location)) updates.put("Location", location);
        }

        this.location = location;
    }

    
    private String create(Connection conn) throws ExchangeException {

        StringBuffer msg = new StringBuffer();
        msg.append("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
        msg.append("<soap:Envelope xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\" xmlns:m=\"http://schemas.microsoft.com/exchange/services/2006/messages\">");
        //"<soap:Header>"<t:RequestServerVersion Version=\"Exchange2007_SP1\"/></soap:Header>"
        msg.append("<soap:Body><m:CreateItem>");
        msg.append("<m:SavedItemFolderId>");
        msg.append("<t:DistinguishedFolderId Id=\"calendar\" />"); //msg.append("<t:FolderId Id=\"" + folderID + "\"/>");
        msg.append("</m:SavedItemFolderId>");
        msg.append("<m:Items>");
        msg.append("<t:CalendarItem>");


      //Add categories
        if (!categories.isEmpty()){
            msg.append("<t:Categories>");
            java.util.Iterator<String> it = categories.iterator();
            while (it.hasNext()){
                msg.append("<t:String>" + it.next() + "</t:String>");
            }
            msg.append("</t:Categories>");
        }


        msg.append("</t:CalendarItem>");
        msg.append("</m:Items>");
        msg.append("</m:CreateItem>");
        msg.append("</soap:Body>");
        msg.append("</soap:Envelope>");



        org.w3c.dom.Document xml = conn.execute(msg.toString());
        org.w3c.dom.NodeList nodes = xml.getElementsByTagName("t:ItemId");
        if (nodes!=null && nodes.getLength()>0){
            id = javaxt.xml.DOM.getAttributeValue(nodes.item(0), "Id");
        }

        return id;
    }

    
  //**************************************************************************
  //** updateInvite
  //**************************************************************************
  /** To remove attendees from a list, you make an UpdateItem call on the
   *  CalendarItem, and use the SetItemField to update the attendee list with
   *  all the current attendees you want on the list.
   */
    private void updateInvite(){

        /*
        <m:UpdateItem ConflictResolution="AutoResolve" SendMeetingInvitationsOrCancellations="SendOnlyToChanged" xmlns:m="http://.../messages" xmlns:t="http://.../types">
         <m:ItemChanges>
            <t:ItemChange>
              <t:ItemId Id="AAApAHByaW5ja..." ChangeKey="DwAAA..." />
              <t:Updates>
                <!-- When using SetItemField, we must specify all Required Attendees -->
                <t:SetItemField>
                  <t:FieldURI FieldURI="calendar:RequiredAttendees" />
                  <t:CalendarItem>
                    <t:RequiredAttendees>
                      <t:Attendee>
                        <t:Mailbox>
                          <t:Name>Andy</t:Name>
                          <t:EmailAddress>andy@contoso.com</t:EmailAddress>
                        </t:Mailbox>
                      </t:Attendee>
                      <t:Attendee>
                        <t:Mailbox>
                          <t:Name>Carol</t:Name>
                          <t:EmailAddress>carol@contoso.com</t:EmailAddress>
                        </t:Mailbox>
                      </t:Attendee>
                    </t:RequiredAttendees>
                  </t:CalendarItem>
                </t:SetItemField>
              </t:Updates>
            </t:ItemChange>
          </m:ItemChanges>
        </m:UpdateItem>
         */
    }
}