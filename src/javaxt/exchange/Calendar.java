package javaxt.exchange;

//******************************************************************************
//**  Calendar Class
//******************************************************************************
/**
 *   Enter class description here
 *
 ******************************************************************************/

public class Calendar {


  //**************************************************************************
  //** Constructor
  //**************************************************************************
  /** Creates a new instance of Calendar. */

    public Calendar() {

    }
    

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