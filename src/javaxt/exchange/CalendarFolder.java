package javaxt.exchange;

//******************************************************************************
//**  CalendarFolder Class
//******************************************************************************
/**
 *   Enter class description here
 *
 ******************************************************************************/

public class CalendarFolder extends Folder {

    private java.util.ArrayList<String> props = new java.util.ArrayList<String>();

  //**************************************************************************
  //** Constructor
  //**************************************************************************
  /** Creates a new instance of CalendarFolder. */

    public CalendarFolder(Connection conn) throws ExchangeException {
        super("calendar", conn);
        props.add("calendar:TimeZone");
    }

    /*
    public CalendarEvent[] getEvents() throws ExchangeException {

        java.util.ArrayList<CalendarEvent> events = new java.util.ArrayList<CalendarEvent>();

        int offset = 0;
        int maxRecords = 25;
        while(true){

            CalendarEvent[] arr = getEvents(maxRecords, offset);
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
   */
    private CalendarEvent[] getEvents(int numEntries, int offset) throws ExchangeException {

        java.util.ArrayList<CalendarEvent> events = new java.util.ArrayList<CalendarEvent>();
        org.w3c.dom.Document xml = getItems(numEntries, offset, props, null);
        org.w3c.dom.Node[] nodes = javaxt.xml.DOM.getElementsByTagName("CalendarItem", xml);

        for (org.w3c.dom.Node node : nodes){
            events.add(new CalendarEvent(node));
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
        org.w3c.dom.Document xml = getItems("<m:CalendarView StartDate=\"" + StartDate + "\" EndDate=\"" + EndDate + "\"/>", props, null);
        org.w3c.dom.Node[] nodes = javaxt.xml.DOM.getElementsByTagName("CalendarItem", xml);

        for (org.w3c.dom.Node node : nodes){
            CalendarEvent event = new CalendarEvent(node);
            if (event.isAllDayEvent() && event.getStartTime().compareTo(start, "days")==-1){
                
            }
            else{
                events.add(new CalendarEvent(node));
            }
        }



        return events.toArray(new CalendarEvent[events.size()]);
    }
}