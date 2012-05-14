package javaxt.exchange;

//******************************************************************************
//**  CalendarInvite Class
//******************************************************************************
/**
 *   Used to represent an invitation to a calendar event.
 *
 ******************************************************************************/

public class CalendarInvite {

    public CalendarInvite(CalendarEvent event){
        
    }

    public void accept(boolean tentative){
        if (tentative) sendNotification("TentativelyAcceptItem");
        else sendNotification("AcceptItem");
    }

    public void decline(){
        sendNotification("DeclineItem");
    }

    private void sendNotification(String action){
        /*
        <?xml version="1.0" encoding="utf-8"?>
        <soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
                       xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
          <soap:Body>
            <CreateItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages"
                        MessageDisposition="SendAndSaveCopy">
              <Items>
                <AcceptItem xmlns="http://schemas.microsoft.com/exchange/services/2006/types">
                  <ReferenceItemId Id="AAAlAFVzZ"
                                   ChangeKey="CwAAABYAA"/>
                </AcceptItem>
              </Items>
            </CreateItem>
          </soap:Body>
        </soap:Envelope>
        */
    }
}
