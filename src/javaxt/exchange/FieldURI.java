package javaxt.exchange;

public class FieldURI {

    private static final String[] fieldURIs = "calendar:AdjacentMeetingCount,calendar:AdjacentMeetings,calendar:AllowNewTimeProposal,calendar:AppointmentReplyTime,calendar:AppointmentSequenceNumber,calendar:AppointmentState,calendar:CalendarItemType,calendar:ConferenceType,calendar:ConflictingMeetingCount,calendar:ConflictingMeetings,calendar:DateTimeStamp,calendar:DeletedOccurrences,calendar:Duration,calendar:End,calendar:EndTimeZone,calendar:FirstOccurrence,calendar:IsAllDayEvent,calendar:IsCancelled,calendar:IsMeeting,calendar:IsOnlineMeeting,calendar:IsRecurring,calendar:IsResponseRequested,calendar:LastOccurrence,calendar:LegacyFreeBusyStatus,calendar:Location,calendar:MeetingRequestWasSent,calendar:MeetingTimeZone,calendar:MeetingWorkspaceUrl,calendar:ModifiedOccurrences,calendar:MyResponseType,calendar:NetShowUrl,calendar:OptionalAttendees,calendar:Organizer,calendar:OriginalStart,calendar:Recurrence,calendar:RecurrenceId,calendar:RequiredAttendees,calendar:Resources,calendar:Start,calendar:StartTimeZone,calendar:TimeZone,calendar:UID,calendar:When,contacts:Alias,contacts:AssistantName,contacts:Birthday,contacts:BusinessHomePage,contacts:Children,contacts:Companies,contacts:CompanyName,contacts:CompleteName,contacts:ContactSource,contacts:Culture,contacts:Department,contacts:DirectoryId,contacts:DirectReports,contacts:DisplayName,contacts:EmailAddresses,contacts:FileAs,contacts:FileAsMapping,contacts:Generation,contacts:GivenName,contacts:HasPicture,contacts:HasPicture,contacts:ImAddresses,contacts:Initials,contacts:JobTitle,contacts:Manager,contacts:ManagerMailbox,contacts:MiddleName,contacts:Mileage,contacts:MSExchangeCertificate,contacts:Nickname,contacts:Notes,contacts:OfficeLocation,contacts:PhoneNumbers,contacts:PhoneticFirstName,contacts:PhoneticFullName,contacts:PhoneticLastName,contacts:Photo,contacts:PhysicalAddresses,contacts:PostalAddressIndex,contacts:Profession,contacts:SpouseName,contacts:Surname,contacts:UserSMIMECertificate,contacts:WeddingAnniversary,conversation:Categories,conversation:ConversationId,conversation:ConversationTopic,conversation:FlagStatus,conversation:GlobalCategories,conversation:GlobalFlagStatus,conversation:GlobalHasAttachments,conversation:GlobalImportance,conversation:GlobalItemClasses,conversation:GlobalItemIds,conversation:GlobalLastDeliveryTime,conversation:GlobalMessageCount,conversation:GlobalSize,conversation:GlobalUniqueRecipients,conversation:GlobalUniqueSenders,conversation:GlobalUniqueUnreadSenders,conversation:GlobalUnreadCount,conversation:HasAttachments,conversation:Importance,conversation:ItemClasses,conversation:ItemIds,conversation:LastDeliveryTime,conversation:MessageCount,conversation:Size,conversation:UniqueRecipients,conversation:UniqueSenders,conversation:UniqueUnreadSenders,conversation:UnreadCount,distributionlist:Members,folder:ChildFolderCount,folder:DisplayName,folder:EffectiveRights,folder:FolderClass,folder:FolderId,folder:ManagedFolderInformation,folder:ParentFolderId,folder:PermissionSet,folder:SearchParameters,folder:SharingEffectiveRights,folder:TotalCount,folder:UnreadCount,item:Attachments,item:Body,item:Categories,item:ConversationId,item:Culture,item:DateTimeCreated,item:DateTimeReceived,item:DateTimeSent,item:DisplayCc,item:DisplayTo,item:EffectiveRights,item:HasAttachments,item:Importance,item:InReplyTo,item:InternetMessageHeaders,item:IsAssociated,item:IsDraft,item:IsFromMe,item:IsResend,item:IsSubmitted,item:IsUnmodified,item:ItemClass,item:ItemId,item:LastModifiedName,item:LastModifiedTime,item:MimeContent,item:ParentFolderId,item:ReminderDueBy,item:ReminderIsSet,item:ReminderMinutesBeforeStart,item:ResponseObjects,item:Sensitivity,item:Size,item:Subject,item:UniqueBody,item:WebClientEditFormQueryString,item:WebClientReadFormQueryString,meeting:AssociatedCalendarItemId,meeting:HasBeenProcessed,meeting:IsDelegated,meeting:IsOutOfDate,meeting:ResponseType,meetingRequest:IntendedFreeBusyStatus,meetingRequest:MeetingRequestType,message:BccRecipients,message:CcRecipients,message:ConversationIndex,message:ConversationTopic,message:From,message:InternetMessageId,message:IsDeliveryReceiptRequested,message:IsRead,message:IsReadReceiptRequested,message:IsResponseRequested,message:References,message:ReplyTo,message:Sender,message:ToRecipients,postitem:PostedTime,task:ActualWork,task:AssignedTime,task:BillingInformation,task:ChangeCount,task:Companies,task:CompleteDate,task:Contacts,task:DelegationState,task:Delegator,task:DueDate,task:IsAssignmentEditable,task:IsComplete,task:IsRecurring,task:IsTeamTask,task:Mileage,task:Owner,task:PercentComplete,task:Recurrence,task:StartDate,task:Status,task:StatusDescription,task:TotalWork".split(",");


    private String fieldURI;

    public FieldURI(String fieldURI){
        this.fieldURI = fieldURI;
    }

    protected FieldURI(){
    }

    public String toString(){
        return fieldURI;
    }
        
    public int hashCode(){
        return fieldURI.hashCode();
    }
    
    public boolean equals(Object obj){
        if (obj!=null){
            if (obj instanceof FieldURI){
                return ((FieldURI) obj).hashCode()==this.hashCode();
            }
        }
        return false;
    }
    
    public String toXML(String namespace){
        return "<" + namespace + ":FieldURI FieldURI=\"" + this.toString() + "\"/>";
    }


  /** Returns a list of all known Field URIs*/
    public static String[] getFieldURIs(){
        return fieldURIs;
    }
}
