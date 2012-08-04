package javaxt.exchange;

//******************************************************************************
//**  CalendarEvent Class
//******************************************************************************
/**
 *   Used to represent a calendar entry.
 *
 *   TODOs:
 *   1. Meeting reminders
 *   2. Recurring events
 *
 ******************************************************************************/

public class CalendarEvent extends FolderItem {

    private String location;
    private javaxt.utils.Date startTime;
    private javaxt.utils.Date endTime;
    private Mailbox organizer;
    private java.util.HashMap<Mailbox, Boolean> attendees = new java.util.HashMap<Mailbox, Boolean>();
    private String freeBusyStatus;
    private Integer reminder;


  //**************************************************************************
  //** Constructor
  //**************************************************************************
  /** This constructor is provided for application developers who wish to
   *  extend this class.
   */
    protected CalendarEvent(){}


  //**************************************************************************
  //** Constructor
  //**************************************************************************
  /** Creates a new instance of this class using an existing Calendar Event,
   *  effectively creating a clone.
   */
    public CalendarEvent(javaxt.exchange.CalendarEvent event){

      //General information
        this.id = event.id;
        this.subject = event.subject;
        this.body = event.body;
        this.bodyType = event.bodyType;
        this.categories = event.categories;
        this.updates = event.updates;
        this.lastModified = event.lastModified;
        this.extendedProperties = event.extendedProperties;

      //Calendar specific information
        this.location = event.location;
        this.startTime = event.startTime;
        this.endTime = event.endTime;
        this.organizer = event.organizer;
        this.attendees = event.attendees;
        this.freeBusyStatus = event.freeBusyStatus;
        this.reminder = event.reminder;
    }


  //**************************************************************************
  //** Constructor
  //**************************************************************************
  /** Creates a new instance of this class with some basic event information.
   */
    public CalendarEvent(String title, String description, String location, Mailbox organizer, javaxt.utils.Date startTime, int duration){
        this.setSubject(title);
        this.setBody(description, "Text");
        this.setLocation(location);
        this.setOrganizer(organizer);
        this.setStartTime(startTime.getDate());
        this.setEndTime(startTime.add(duration, "minutes"));
    }


  //**************************************************************************
  //** Constructor
  //**************************************************************************
  /** Creates a new instance of this class with an id
   */
    public CalendarEvent(String exchangeID, Connection conn, ExtendedProperty[] AdditionalProperties) throws ExchangeException{
        super(exchangeID, conn, AdditionalProperties);
        parseCalendarItem();
    }


  //**************************************************************************
  //** Constructor
  //**************************************************************************
  /** Creates a new instance of this class with an id
   */
    public CalendarEvent(String exchangeID, Connection conn) throws ExchangeException{
        this(exchangeID, conn, null);
    }


  //**************************************************************************
  //** Constructor
  //**************************************************************************
  /** Creates a new instance of this class using event information found in a
   *  "*.ics" file (iCalendar Event).
   *  http://en.wikipedia.org/wiki/ICalendar
   */
    public CalendarEvent(String iCalendar) throws Exception{
        org.w3c.dom.Document xml = iCalToXML(iCalendar);
        throw new Exception("Not Implemented.");
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

        String timezone = null;

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
                    try{
                        startTime = new javaxt.utils.Date(javaxt.xml.DOM.getNodeValue(outerNode));
                    }
                    catch(java.text.ParseException e){}
                }
                else if (nodeName.equalsIgnoreCase("End")){
                    try{
                        endTime = new javaxt.utils.Date(javaxt.xml.DOM.getNodeValue(outerNode));
                    }
                    catch(java.text.ParseException e){}
                }
                else if (nodeName.equalsIgnoreCase("TimeZone")){                    
                    timezone = javaxt.xml.DOM.getNodeValue(outerNode);
                }
                else if(nodeName.equalsIgnoreCase("IsAllDayEvent")){
                    String b = javaxt.xml.DOM.getNodeValue(outerNode).toLowerCase();
                    //this (b.startsWith("t"))
                }
                else if(nodeName.equalsIgnoreCase("LegacyFreeBusyStatus")){
                    freeBusyStatus = javaxt.xml.DOM.getNodeValue(outerNode);
                }
                else if(nodeName.equalsIgnoreCase("Organizer")){
                    org.w3c.dom.Node[] mailbox = javaxt.xml.DOM.getElementsByTagName("Mailbox", outerNode);
                    if (mailbox.length>0) organizer = new Mailbox(mailbox[0]);
                }
                else if(nodeName.equalsIgnoreCase("RequiredAttendees")){
                    for (org.w3c.dom.Node node : javaxt.xml.DOM.getElementsByTagName("Mailbox", outerNode)){
                        attendees.put(new Mailbox(node), true);
                    }
                }
                else if(nodeName.equalsIgnoreCase("OptionalAttendees")){
                    for (org.w3c.dom.Node node : javaxt.xml.DOM.getElementsByTagName("Mailbox", outerNode)){
                        attendees.put(new Mailbox(node), false);
                    }
                }
            }
        }
        setTimeZone(timezone);
    }


  //**************************************************************************
  //** setTimeZone
  //**************************************************************************
  /** Used to set the time zone for the event. This is important for
   *  calculating all day events.
   */
    public void setTimeZone(String timezone){
        if (timezone==null) return;
        if (timezone.trim().length()==0){
            setTimeZone(java.util.TimeZone.getDefault()); //TODO: Get default timezone from the server?
        }
        else{
            setTimeZone(javaxt.utils.Date.getTimeZone(timezone));
        }
    }


  //**************************************************************************
  //** setTimeZone
  //**************************************************************************
  /** Used to set the time zone for the event. This is important for
   *  calculating all day events.
   */
    public void setTimeZone(java.util.TimeZone timezone){
        startTime.setTimeZone(timezone);
        endTime.setTimeZone(timezone);
    }


  //**************************************************************************
  //** getSubject
  //**************************************************************************
  /** Returns a descriptive title for this event.
   */
    public String getSubject(){
        return super.getSubject();
    }


  //**************************************************************************
  //** setTitle
  //**************************************************************************
  /** Used to set a descriptive title for this event.
   */
    public void setSubject(String title){
        super.setSubject(title);
    }


  //**************************************************************************
  //** getBody
  //**************************************************************************
  /** Used to get the content of the meeting invite.
   */
    public String getBody(){
        return super.getBody();
    }


  //**************************************************************************
  //** setBody
  //**************************************************************************
  /** Used to set the content of the meeting invite (e.g. purpose, agenda,
   *  address, notes, etc.)
   *  @param format Text format ("Best", "HTML", or "Text").
   */
    public void setBody(String description, String format){
        super.setBody(description, format);
    }


  //**************************************************************************
  //** getAvailability
  //**************************************************************************
  /** Returns the organizer's availability during this event
   * (LegacyFreeBusyStatus). Values include "Free", "Busy", "Tentative",
   *  "OOF" (out of office), and "NoData".
   */
    public String getAvailability(){
        return this.freeBusyStatus;
    }


  //**************************************************************************
  //** setAvailability
  //**************************************************************************
  /** Used to set the organizer's availability during this event.
   *  @param freeBusyStatus Values include "Free", "Busy", "Tentative",
   *  "OOF" (out of office), and "NoData".
   */
    public void setAvailability(String freeBusyStatus){
        freeBusyStatus = getValue(freeBusyStatus);

        if (id!=null) {
            if (freeBusyStatus==null && this.freeBusyStatus!=null) updates.put("LegacyFreeBusyStatus", null);
            if (freeBusyStatus!=null && !freeBusyStatus.equals(this.freeBusyStatus)) updates.put("LegacyFreeBusyStatus", freeBusyStatus);
        }

        this.freeBusyStatus = freeBusyStatus;
    }


  //**************************************************************************
  //** getAttendees
  //**************************************************************************
  /** Returns a HashMap with a list of attendees associated with this event.
   *  The key in the HashMap represents a Mailbox item and the value is a
   *  boolean used to indicate whether the attendee is required.
   */
    public java.util.HashMap<Mailbox, Boolean> getAttendees(){
        if (attendees.isEmpty()) return null;
        else return attendees;
    }


  //**************************************************************************
  //** setAttendees
  //**************************************************************************
  /** Used to set attendees associated with this event.
   */
    public void setAttendees(java.util.HashMap<Mailbox, Boolean> attendees) {

        if (attendees==null || attendees.isEmpty()){
            removeAttendees();
            return;
        }

      //See whether any updates are required
        int numMatches = 0;
        int total = 0;
        java.util.Iterator<Mailbox> it = attendees.keySet().iterator();
        while (it.hasNext()){
            total++;
            Mailbox attendee = it.next();
            boolean isRequired = attendees.get(attendee);
            if (this.attendees.containsKey(attendee)){
                if (this.attendees.get(attendee)==isRequired) numMatches++;
            }
        }


        int numAttendees = 0;
        if (this.getAttendees()!=null) numAttendees = this.getAttendees().size();

      //If the input array equals the current list of phoneNumbers, do nothing...
        if (numMatches==total && numMatches==numAttendees){
            return;
        }
        else {
            this.attendees.clear();
            it = attendees.keySet().iterator();
            while (it.hasNext()){
                Mailbox attendee = it.next();
                boolean isRequired = attendees.get(attendee);
                addAttendee(attendee, isRequired);
            }
        }
    }


  //**************************************************************************
  //** addAttendee
  //**************************************************************************
  /**  Used to add an attendee to this event.
   */
    public void addAttendee(Mailbox attendee, boolean isRequired){

        boolean update = false;
        if (attendees.containsKey(attendee)){
            update = (attendees.get(attendee)!=isRequired);
        }
        else{
            update = true;
        }

        if (update){
            attendees.put(attendee, isRequired);
            if (id!=null) updates.put("Attendees", getAttendeeUpdates());
        }
    }


  //**************************************************************************
  //** removeAttendee
  //**************************************************************************
  /**  Used to remove an attendee from this event.
   */
    public void removeAttendee(Mailbox attendee){
        System.out.println(attendees.containsKey(attendee));
        if (attendees.containsKey(attendee)){
            attendees.remove(attendee);
            if (id!=null) updates.put("Attendees", getAttendeeUpdates());
        }
    }


  //**************************************************************************
  //** removeAttendees
  //**************************************************************************
  /** Removes all attendees from this event.
   */
    public void removeAttendees(){
        if (!attendees.isEmpty()){
            this.attendees.clear();
            if (id!=null) updates.put("Attendees", getAttendeeUpdates());
        }
    }


  //**************************************************************************
  //** getAttendeeUpdates
  //**************************************************************************
  /** Returns an xml fragment used to update the list of attendees.
   *
   *  Note that unlike phone numbers, emails, etc. we cannot call
   *  DeleteItemField to remove individual attendees. Instead, we simply call
   *  SetItemField to update the attendee list.
   *
   *  To remove attendees from a list, you make an UpdateItem call on the
   *  CalendarItem, and use the SetItemField to update the attendee list with
   *  all the current attendees you want on the list.
   */
    private String getAttendeeUpdates(){

        if (attendees.isEmpty() && id==null) return "";
        else{
            StringBuffer xml = new StringBuffer();

            if (attendees.values().contains(true)){
                xml.append("<t:SetItemField>");
                xml.append("<t:FieldURI FieldURI=\"calendar:RequiredAttendees\" />");
                xml.append("<t:CalendarItem>");
                xml.append("<t:RequiredAttendees>");
                java.util.Iterator<Mailbox> it = attendees.keySet().iterator();
                while (it.hasNext()){
                    Mailbox attendee = it.next();
                    boolean isRequired = attendees.get(attendee);
                    if (isRequired){
                        xml.append("<t:Attendee>");
                        xml.append(attendee.toXML("t"));
                        xml.append("</t:Attendee>");
                    }
                }
                xml.append("</t:RequiredAttendees>");
                xml.append("</t:CalendarItem>");
                xml.append("</t:SetItemField>");
            }
            else{
                if (id!=null){
                    xml.append("<t:DeleteItemField>");
                    xml.append("<t:FieldURI FieldURI=\"calendar:RequiredAttendees\" /> ");
                    xml.append("</t:DeleteItemField>");
                }
            }

            if (attendees.values().contains(false)){
                xml.append("<t:SetItemField>");
                xml.append("<t:FieldURI FieldURI=\"calendar:OptionalAttendees\" />");
                xml.append("<t:CalendarItem>");
                xml.append("<t:OptionalAttendees>");
                java.util.Iterator<Mailbox> it = attendees.keySet().iterator();
                while (it.hasNext()){
                    Mailbox attendee = it.next();
                    boolean isRequired = attendees.get(attendee);
                    if (!isRequired){
                        xml.append("<t:Attendee>");
                        xml.append(attendee.toXML("t"));
                        xml.append("</t:Attendee>");
                    }
                }
                xml.append("</t:OptionalAttendees>");
                xml.append("</t:CalendarItem>");
                xml.append("</t:SetItemField>");
            }
            else{
                if (id!=null){
                    xml.append("<t:DeleteItemField>");
                    xml.append("<t:FieldURI FieldURI=\"calendar:OptionalAttendees\" /> ");
                    xml.append("</t:DeleteItemField>");
                }
            }

            return xml.toString();
        }
    }


  //**************************************************************************
  //** setStartTime
  //**************************************************************************
  /** Used to set the start time of the event using a java.util.Date.
   */
    public void setStartTime(java.util.Date startTime){
        setStartTime(new javaxt.utils.Date(startTime));
    }


  //**************************************************************************
  //** setStartTime
  //**************************************************************************
  /** Used to set the start time of the event using a javaxt.utils.Date.
   */
    public void setStartTime(javaxt.utils.Date startTime){

        if (id!=null){
            if (startTime==null && this.startTime!=null) updates.put("Start", null);
            if (startTime!=null && !formatDate(startTime).equals(formatDate(this.startTime))){
                updates.put("Start", formatDate(startTime));
                updates.put("IsAllDayEvent", this.isAllDayEvent());
            }
        }
        this.startTime = startTime;
    }


  //**************************************************************************
  //** setStartTime
  //**************************************************************************
  /** Used to set the start time using a String. If the method fails to parse
   *  the string, the new value will be ignored.
   */
    public void setStartTime(String startTime){
        javaxt.utils.Date date = null;
        try{
            date = new javaxt.utils.Date(startTime);
        }
        catch(java.text.ParseException e){}
        setStartTime(date);
    }


  //**************************************************************************
  //** getStartTime
  //**************************************************************************
  /** Returns the date of birth.
   */
    public javaxt.utils.Date getStartTime(){
        return startTime;
    }


  //**************************************************************************
  //** setEndTime
  //**************************************************************************
  /** Used to set the end time of the event using a java.util.Date.
   */
    public void setEndTime(java.util.Date endTime){
        setEndTime(new javaxt.utils.Date(endTime));
    }


  //**************************************************************************
  //** setEndTime
  //**************************************************************************
  /** Used to set the end time of the event using a javaxt.utils.Date.
   */
    public void setEndTime(javaxt.utils.Date endTime){

        if (id!=null){
            if (endTime==null && this.endTime!=null) updates.put("End", null);
            if (endTime!=null && !formatDate(endTime).equals(formatDate(this.endTime))){
                updates.put("End", formatDate(endTime));
                updates.put("IsAllDayEvent", this.isAllDayEvent());
            }
        }
        this.endTime = endTime;
    }


  //**************************************************************************
  //** setEndTime
  //**************************************************************************
  /** Used to set the end time using a String. If the method fails to parse
   *  the string, the new value will be ignored.
   */
    public void setEndTime(String endTime){
        javaxt.utils.Date date = null;
        try{
            date = new javaxt.utils.Date(endTime);
        }
        catch(java.text.ParseException e){}
        setEndTime(date);
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
  //** isAllDayEvent
  //**************************************************************************
  /** Returns a boolean used to indicate whether this is an all day event.
   *  Note that the logic only works if the timezone has been set.
   */
    public boolean isAllDayEvent(){
        return (startTime.getHour()==0 && startTime.getMinute()==0 &&
                startTime.getSecond()==0 && this.getDuration("hours")==24);
    }


    public void setAllDayEvent(javaxt.utils.Date date){
        try{
            String d1 = date.toString("MM/dd/yyyy");
            java.util.Date d2 = new javaxt.utils.Date(d1).add(1, "day");
            this.setStartTime(d1);
            this.setEndTime(d2);
            //IsAllDayEvent
            if (id!=null) {
                updates.put("IsAllDayEvent", true);
            }
        }
        catch(Exception e){}
    }

    public void setAllDayEvent(String date){
        try{
            setAllDayEvent(new javaxt.utils.Date(date));
        }
        catch(java.text.ParseException e){}
    }


  //**************************************************************************
  //** getLocation
  //**************************************************************************
  /** Returns the location of this event.
   */
    public String getLocation(){
        location = getValue(location);
        return location;
    }


  //**************************************************************************
  //** setLocation
  //**************************************************************************
  /** Used to set the location of this event.
   */
    public void setLocation(String location){
        location = getValue(location);

        if (id!=null) {
            if (location==null && this.location!=null) updates.put("Location", null);
            if (location!=null && !location.equals(this.location)) updates.put("Location", location);
        }

        this.location = location;
    }

    public Mailbox getOrganizer(){
        return organizer;
    }

    public void setOrganizer(Mailbox organizer){
        this.organizer = organizer;
    }

    public Integer getReminder(){
        return reminder;
    }

    public void setReminder(Integer reminder){

        if (id!=null) {
            if (reminder==null && this.reminder!=null) updates.put("ReminderMinutesBeforeStart", null);
            if (reminder!=null && !reminder.equals(this.reminder)) updates.put("ReminderMinutesBeforeStart", reminder);
        }

        this.reminder = reminder;
    }



    
    
    private String create(Connection conn) throws ExchangeException {

        StringBuffer msg = new StringBuffer();
        msg.append("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
        msg.append("<soap:Envelope xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\" xmlns:m=\"http://schemas.microsoft.com/exchange/services/2006/messages\">");
        msg.append("<soap:Body><m:CreateItem SendMeetingInvitations=\"SendToAllAndSaveCopy\">");
        msg.append("<m:SavedItemFolderId>");
        msg.append("<t:DistinguishedFolderId Id=\"calendar\" />"); //msg.append("<t:FolderId Id=\"" + folderID + "\"/>");
        msg.append("</m:SavedItemFolderId>");
        msg.append("<m:Items>");
        msg.append("<t:CalendarItem>");

      /*
        Here's is an ordered list of all the contact properties. The first set
        is pretty generic and applies to other Exchange "Items". The second set
        is specific to contacts. The third set is also generic and seems to
        apply to all items. WARNING -- ORDER IS VERY IMPORTANT!!! If you
        mess up the order of the properties, the save will fail - at least it
        did on my Exchange Server 2007 SP3 (8.3)


        MimeContent, ItemId, ParentFolderId, ItemClass, Subject, Sensitivity,
        Body, Attachments, DateTimeReceived, Size, Categories, Importance,
        InReplyTo, IsSubmitted, IsDraft, IsFromMe, IsResend, IsUnmodified,
        InternetMessageHeaders, DateTimeSent, DateTimeCreated, ResponseObjects,
        ReminderDueBy, ReminderIsSet, ReminderMinutesBeforeStart, DisplayCc,
        DisplayTo, HasAttachments, ExtendedProperty, Culture,

        Start, End, OriginalStart, IsAllDayEvent, LegacyFreeBusyStatus,
        Location, When, IsMeeting, IsCancelled, IsRecurring, MeetingRequestWasSent,
        IsResponseRequested, CalendarItemType, MyResponseType, Organizer,
        RequiredAttendees, OptionalAttendees, Resources, ConflictingMeetingCount,
        AdjacentMeetingCount, ConflictingMeetings, AdjacentMeetings, Duration,
        TimeZone, AppointmentReplyTime, AppointmentSequenceNumber, AppointmentState,
        Recurrence, FirstOccurrence, LastOccurrence, ModifiedOccurrences,
        DeletedOccurrences, MeetingTimeZone, StartTimeZone, EndTimeZone,
        ConferenceType, AllowNewTimeProposal, IsOnlineMeeting, MeetingWorkspaceUrl,
        NetShowUrl,

        EffectiveRights, LastModifiedName, LastModifiedTime, IsAssociated,
        WebClientReadFormQueryString, WebClientEditFormQueryString,
        ConversationId, UniqueBody
      */

        if (getSubject()!=null) msg.append("<t:Subject>" + getSubject() + "</t:Subject>");
        if (getBody()!=null) msg.append("<t:Body BodyType=\"" + bodyType + "\">" + getBody() + "</t:Body>");

      //Add categories
        if (!categories.isEmpty()){
            msg.append("<t:Categories>");
            java.util.Iterator<String> it = categories.iterator();
            while (it.hasNext()){
                msg.append("<t:String>" + it.next() + "</t:String>");
            }
            msg.append("</t:Categories>");
        }

        if (this.getReminder()!=null){
            msg.append("<t:ReminderIsSet>true</t:ReminderIsSet>");
            msg.append("<t:ReminderMinutesBeforeStart>" + this.getReminder() + "</t:ReminderMinutesBeforeStart>");
        }


      //Add extended properties
        ExtendedProperty[] properties = this.getExtendedProperties();
        if (properties!=null){
            for (ExtendedProperty property : properties){
                msg.append(property.toXML("t", "create"));
            }
        }


        msg.append("<t:Start>" + formatDate(getStartTime()) + "</t:Start>");
        msg.append("<t:End>" + formatDate(getEndTime()) + "</t:End>");

        msg.append("<t:IsAllDayEvent>" + this.isAllDayEvent() + "</t:IsAllDayEvent>");
        if (getLocation()!=null) msg.append("<t:Location>" + getLocation() + "</t:Location>");
        //msg.append("<t:IsCancelled>" + false + "</t:IsCancelled>");
        //msg.append("<t:CalendarItemType>" + "Single" + "</t:CalendarItemType>");
        //msg.append("<t:Organizer>" + getOrganizer().toXML("t") + "</t:Organizer>");


        if (attendees.values().contains(true)){
            msg.append("<t:RequiredAttendees>");
            java.util.Iterator<Mailbox> it = attendees.keySet().iterator();
            while (it.hasNext()){
                Mailbox attendee = it.next();
                boolean isRequired = attendees.get(attendee);
                if (isRequired){
                    msg.append("<t:Attendee>");
                    msg.append(attendee.toXML("t"));
                    msg.append("</t:Attendee>");
                }
            }
            msg.append("</t:RequiredAttendees>");
        }
        

        if (attendees.values().contains(false)){
            msg.append("<t:OptionalAttendees>");
            java.util.Iterator<Mailbox> it = attendees.keySet().iterator();
            while (it.hasNext()){
                Mailbox attendee = it.next();
                boolean isRequired = attendees.get(attendee);
                if (!isRequired){
                    msg.append("<t:Attendee>");
                    msg.append(attendee.toXML("t"));
                    msg.append("</t:Attendee>");
                }
            }
            msg.append("</t:OptionalAttendees>");
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
  //** save
  //**************************************************************************
  /**  Used to save/update an event. Returns the Exchange ID for the item.
   */
    public String save(Connection conn) throws ExchangeException {

        java.util.HashMap<String, String> options = new java.util.HashMap<String, String>();
        options.put("ConflictResolution", "AutoResolve");
        options.put("SendMeetingInvitationsOrCancellations", "SendOnlyToChanged");

        if (id==null) create(conn);
        else update("CalendarItem", "calendar", options, conn);

      //Update last modified date
        this.lastModified = new CalendarEvent(id, conn).getLastModifiedTime();
        
        return id;
    }

    
  //**************************************************************************
  //** delete
  //**************************************************************************
  /**  Used to delete an event. Provides an option to notify invitees.
   */
    public void delete(Connection conn, boolean notify) throws ExchangeException {
        java.util.HashMap<String, String> options = new java.util.HashMap<String, String>();
        options.put("DeleteType", "MoveToDeletedItems");
        if (notify) options.put("SendMeetingCancellations", "SendToAllAndSaveCopy");
        else options.put("SendMeetingCancellations", "SendToNone");
        super.delete(options, conn);
    }


  //**************************************************************************
  //** toString
  //**************************************************************************
  /**  Returns a string representation of this item.
   */
    public String toString(){
        return this.getSubject();
    }


  //**************************************************************************
  //** iCalToXML
  //**************************************************************************
  /** Used to convert an iCalendar ICS File into an XML Document.
   */
    private org.w3c.dom.Document iCalToXML(String iCalendar){
        StringBuffer str = new StringBuffer();
        java.util.ArrayList<String> nodes = new java.util.ArrayList<String>();
        String[] rows = iCalendar.split("\r\n");
        for (int i=0; i<rows.length; i++){
            String row = rows[i].trim();

          //Special case for wrapped lines
            while (true){
                if (i+1>rows.length-1) break;
                String nextRow = rows[i+1];
                if (nextRow.trim().length()>0 && nextRow.startsWith(" ")){
                    i++;
                    row += nextRow.trim();
                }
                else{
                    break;
                }
            }

            if (row.contains(":")){
                String key = row.substring(0, row.indexOf(":")).trim();
                String value = row.substring(row.indexOf(":")+1).trim();
                String[] attr = new String[0];

                if (key.contains(";")){
                    attr = key.substring(key.indexOf(";")+1).split(";");
                    key = key.substring(0, key.indexOf(";")).trim();
                }
                key = key.toLowerCase();

                if (key.equals("begin")){
                    nodes.add(key);
                    str.append("<" + value.toLowerCase() + ">\r\n");
                }
                else if (key.equals("end")){
                    nodes.remove(nodes.size()-1);
                    str.append("</" + value.toLowerCase() + ">\r\n");
                }
                else{
                    str.append("<" + key + ">");
                    str.append(value);
                    str.append("</" + key + ">\r\n");
                }


            }


        }

        System.out.println(str);
        return null;
    }
}