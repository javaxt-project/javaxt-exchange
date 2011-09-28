package javaxt.exchange;

//******************************************************************************
//**  CalendarEvent Class
//******************************************************************************
/**
 *   Used to represent a calendar entry.
 *
 ******************************************************************************/

public class CalendarEvent {

    private String id;
    private String subject;
    private javaxt.utils.Date startTime;
    private javaxt.utils.Date endTime;


  //**************************************************************************
  //** Constructor
  //**************************************************************************
  /** Creates a new instance of CalendarEvent. */

    public CalendarEvent(org.w3c.dom.Node calendarItemNode) {
        parseContact(calendarItemNode);
    }


  //**************************************************************************
  //** parseContact
  //**************************************************************************
  /** Used to parse an xml node with event information.
   */
    private void parseContact(org.w3c.dom.Node calendarItemNode){
        org.w3c.dom.NodeList outerNodes = calendarItemNode.getChildNodes();
        for (int i=0; i<outerNodes.getLength(); i++){
            org.w3c.dom.Node outerNode = outerNodes.item(i);
            if (outerNode.getNodeType()==1){
                String nodeName = outerNode.getNodeName();
                if (nodeName.contains(":")) nodeName = nodeName.substring(nodeName.indexOf(":")+1);
                if (nodeName.equalsIgnoreCase("ItemId")){
                    id = javaxt.xml.DOM.getAttributeValue(outerNode, "Id");
                    //changeKey = javaxt.xml.DOM.getAttributeValue(outerNode, "ChangeKey");
                }
                else if(nodeName.equalsIgnoreCase("Subject")){
                    subject = javaxt.xml.DOM.getNodeValue(outerNode);
                }
                else if(nodeName.equalsIgnoreCase("Start")){
                    javaxt.utils.Date date = new javaxt.utils.Date(javaxt.xml.DOM.getNodeValue(outerNode));
                    if (!date.failedToParse()) startTime = date;
                }
                else if(nodeName.equalsIgnoreCase("End")){
                    javaxt.utils.Date date = new javaxt.utils.Date(javaxt.xml.DOM.getNodeValue(outerNode));
                    if (!date.failedToParse()) endTime = date;
                }
            }
        }
    }


    public String getID(){
        return id;
    }

    public String getSubject(){
        return subject;
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