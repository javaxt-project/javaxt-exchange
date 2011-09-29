package javaxt.exchange;

//******************************************************************************
//**  CalendarFolder Class
//******************************************************************************
/**
 *   Enter class description here
 *
 ******************************************************************************/

public class CalendarFolder extends Folder {

    private Connection conn;

  //**************************************************************************
  //** Constructor
  //**************************************************************************
  /** Creates a new instance of CalendarFolder. */

    public CalendarFolder(Connection conn) throws ExchangeException {
        super("calendar", conn);
        this.conn = conn;
    }


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


    public CalendarEvent[] getEvents(int numEntries, int offset) throws ExchangeException {

        java.util.ArrayList<CalendarEvent> events = new java.util.ArrayList<CalendarEvent>();
        org.w3c.dom.Document xml = new javaxt.io.File("/temp/exchange-calendaritems.xml").getXML(); //getItems(conn, numEntries, offset);
        org.w3c.dom.Node[] nodes = javaxt.xml.DOM.getElementsByTagName("CalendarItem", xml);

        for (org.w3c.dom.Node node : nodes){
            events.add(new CalendarEvent(node));
        }
        return events.toArray(new CalendarEvent[events.size()]);
    }
}