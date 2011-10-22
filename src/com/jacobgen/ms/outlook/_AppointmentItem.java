/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class _AppointmentItem extends Dispatch {

	public static final String componentName = "Outlook._AppointmentItem";

	public _AppointmentItem() {
		super(componentName);
	}

	/**
	* This constructor is used instead of a case operation to
	* turn a Dispatch object into a wider object - it must exist
	* in every wrapper class whose instances may be returned from
	* method calls wrapped in VT_DISPATCH Variants.
	*/
	public _AppointmentItem(Dispatch d) {
		// take over the IDispatch pointer
		m_pDispatch = d.m_pDispatch;
		// null out the input's pointer
		d.m_pDispatch = 0;
	}

	public _AppointmentItem(String compName) {
		super(compName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type _Application
	 */
	public _Application getApplication() {
		return new _Application(Dispatch.get(this, "Application").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getClass1() {
		return Dispatch.get(this, "Class").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type _NameSpace
	 */
	public _NameSpace getSession() {
		return new _NameSpace(Dispatch.get(this, "Session").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Object
	 */
	public Object getParent() {
		return Dispatch.get(this, "Parent");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Actions
	 */
	public Actions getActions() {
		return new Actions(Dispatch.get(this, "Actions").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Attachments
	 */
	public Attachments getAttachments() {
		return new Attachments(Dispatch.get(this, "Attachments").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getBillingInformation() {
		return Dispatch.get(this, "BillingInformation").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param billingInformation an input-parameter of type String
	 */
	public void setBillingInformation(String billingInformation) {
		Dispatch.put(this, "BillingInformation", billingInformation);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getBody() {
		return Dispatch.get(this, "Body").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param body an input-parameter of type String
	 */
	public void setBody(String body) {
		Dispatch.put(this, "Body", body);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getCategories() {
		return Dispatch.get(this, "Categories").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param categories an input-parameter of type String
	 */
	public void setCategories(String categories) {
		Dispatch.put(this, "Categories", categories);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getCompanies() {
		return Dispatch.get(this, "Companies").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param companies an input-parameter of type String
	 */
	public void setCompanies(String companies) {
		Dispatch.put(this, "Companies", companies);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getConversationIndex() {
		return Dispatch.get(this, "ConversationIndex").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getConversationTopic() {
		return Dispatch.get(this, "ConversationTopic").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type java.util.Date
	 */
	public java.util.Date getCreationTime() {
		return Dispatch.get(this, "CreationTime").getJavaDate();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getEntryID() {
		return Dispatch.get(this, "EntryID").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type FormDescription
	 */
	public FormDescription getFormDescription() {
		return new FormDescription(Dispatch.get(this, "FormDescription").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type _Inspector
	 */
	public _Inspector getGetInspector() {
		return new _Inspector(Dispatch.get(this, "GetInspector").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getImportance() {
		return Dispatch.get(this, "Importance").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param importance an input-parameter of type int
	 */
	public void setImportance(int importance) {
		Dispatch.put(this, "Importance", new Variant(importance));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type java.util.Date
	 */
	public java.util.Date getLastModificationTime() {
		return Dispatch.get(this, "LastModificationTime").getJavaDate();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Variant
	 */
	public Variant getMAPIOBJECT() {
		return Dispatch.get(this, "MAPIOBJECT");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getMessageClass() {
		return Dispatch.get(this, "MessageClass").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param messageClass an input-parameter of type String
	 */
	public void setMessageClass(String messageClass) {
		Dispatch.put(this, "MessageClass", messageClass);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getMileage() {
		return Dispatch.get(this, "Mileage").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param mileage an input-parameter of type String
	 */
	public void setMileage(String mileage) {
		Dispatch.put(this, "Mileage", mileage);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getNoAging() {
		return Dispatch.get(this, "NoAging").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param noAging an input-parameter of type boolean
	 */
	public void setNoAging(boolean noAging) {
		Dispatch.put(this, "NoAging", new Variant(noAging));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getOutlookInternalVersion() {
		return Dispatch.get(this, "OutlookInternalVersion").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getOutlookVersion() {
		return Dispatch.get(this, "OutlookVersion").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getSaved() {
		return Dispatch.get(this, "Saved").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getSensitivity() {
		return Dispatch.get(this, "Sensitivity").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param sensitivity an input-parameter of type int
	 */
	public void setSensitivity(int sensitivity) {
		Dispatch.put(this, "Sensitivity", new Variant(sensitivity));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getSize() {
		return Dispatch.get(this, "Size").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getSubject() {
		return Dispatch.get(this, "Subject").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param subject an input-parameter of type String
	 */
	public void setSubject(String subject) {
		Dispatch.put(this, "Subject", subject);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getUnRead() {
		return Dispatch.get(this, "UnRead").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param unRead an input-parameter of type boolean
	 */
	public void setUnRead(boolean unRead) {
		Dispatch.put(this, "UnRead", new Variant(unRead));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type UserProperties
	 */
	public UserProperties getUserProperties() {
		return new UserProperties(Dispatch.get(this, "UserProperties").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param saveMode an input-parameter of type int
	 */
	public void close(int saveMode) {
		Dispatch.call(this, "Close", new Variant(saveMode));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Object
	 */
	public Object copy() {
		return Dispatch.call(this, "Copy");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void delete() {
		Dispatch.call(this, "Delete");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param modal an input-parameter of type Variant
	 */
	public void display(Variant modal) {
		Dispatch.call(this, "Display", modal);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void display() {
		Dispatch.call(this, "Display");
	}

//	/**
//	 * Wrapper for calling the ActiveX-Method and receiving the output-parameter(s).
//	 * @param modal an input-parameter of type Variant
//	 */
//	public void display(Variant modal) {
//		Dispatch.call(this, "Display", modal);
//
//	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param destFldr an input-parameter of type MAPIFolder
	 * @return the result is of type Object
	 */
	public Object move(MAPIFolder destFldr) {
		return Dispatch.call(this, "Move", destFldr);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void printOut() {
		Dispatch.call(this, "PrintOut");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void save() {
		Dispatch.call(this, "Save");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param path an input-parameter of type String
	 * @param type an input-parameter of type Variant
	 */
	public void saveAs(String path, Variant type) {
		Dispatch.call(this, "SaveAs", path, type);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param path an input-parameter of type String
	 */
	public void saveAs(String path) {
		Dispatch.call(this, "SaveAs", path);
	}

//	/**
//	 * Wrapper for calling the ActiveX-Method and receiving the output-parameter(s).
//	 * @param path an input-parameter of type String
//	 * @param type an input-parameter of type Variant
//	 */
//	public void saveAs(String path, Variant type) {
//		Dispatch.call(this, "SaveAs", path, type);
//
//	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getAllDayEvent() {
		return Dispatch.get(this, "AllDayEvent").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param allDayEvent an input-parameter of type boolean
	 */
	public void setAllDayEvent(boolean allDayEvent) {
		Dispatch.put(this, "AllDayEvent", new Variant(allDayEvent));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getBusyStatus() {
		return Dispatch.get(this, "BusyStatus").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param busyStatus an input-parameter of type int
	 */
	public void setBusyStatus(int busyStatus) {
		Dispatch.put(this, "BusyStatus", new Variant(busyStatus));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getDuration() {
		return Dispatch.get(this, "Duration").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param duration an input-parameter of type int
	 */
	public void setDuration(int duration) {
		Dispatch.put(this, "Duration", new Variant(duration));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type java.util.Date
	 */
	public java.util.Date getEnd() {
		return Dispatch.get(this, "End").getJavaDate();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param end an input-parameter of type java.util.Date
	 */
	public void setEnd(java.util.Date end) {
		Dispatch.put(this, "End", new Variant(end));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getIsOnlineMeeting() {
		return Dispatch.get(this, "IsOnlineMeeting").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param isOnlineMeeting an input-parameter of type boolean
	 */
	public void setIsOnlineMeeting(boolean isOnlineMeeting) {
		Dispatch.put(this, "IsOnlineMeeting", new Variant(isOnlineMeeting));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getIsRecurring() {
		return Dispatch.get(this, "IsRecurring").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getLocation() {
		return Dispatch.get(this, "Location").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param location an input-parameter of type String
	 */
	public void setLocation(String location) {
		Dispatch.put(this, "Location", location);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getMeetingStatus() {
		return Dispatch.get(this, "MeetingStatus").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param meetingStatus an input-parameter of type int
	 */
	public void setMeetingStatus(int meetingStatus) {
		Dispatch.put(this, "MeetingStatus", new Variant(meetingStatus));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getNetMeetingAutoStart() {
		return Dispatch.get(this, "NetMeetingAutoStart").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param netMeetingAutoStart an input-parameter of type boolean
	 */
	public void setNetMeetingAutoStart(boolean netMeetingAutoStart) {
		Dispatch.put(this, "NetMeetingAutoStart", new Variant(netMeetingAutoStart));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getNetMeetingOrganizerAlias() {
		return Dispatch.get(this, "NetMeetingOrganizerAlias").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param netMeetingOrganizerAlias an input-parameter of type String
	 */
	public void setNetMeetingOrganizerAlias(String netMeetingOrganizerAlias) {
		Dispatch.put(this, "NetMeetingOrganizerAlias", netMeetingOrganizerAlias);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getNetMeetingServer() {
		return Dispatch.get(this, "NetMeetingServer").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param netMeetingServer an input-parameter of type String
	 */
	public void setNetMeetingServer(String netMeetingServer) {
		Dispatch.put(this, "NetMeetingServer", netMeetingServer);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getNetMeetingType() {
		return Dispatch.get(this, "NetMeetingType").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param netMeetingType an input-parameter of type int
	 */
	public void setNetMeetingType(int netMeetingType) {
		Dispatch.put(this, "NetMeetingType", new Variant(netMeetingType));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getOptionalAttendees() {
		return Dispatch.get(this, "OptionalAttendees").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param optionalAttendees an input-parameter of type String
	 */
	public void setOptionalAttendees(String optionalAttendees) {
		Dispatch.put(this, "OptionalAttendees", optionalAttendees);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getOrganizer() {
		return Dispatch.get(this, "Organizer").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Recipients
	 */
	public Recipients getRecipients() {
		return new Recipients(Dispatch.get(this, "Recipients").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getRecurrenceState() {
		return Dispatch.get(this, "RecurrenceState").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getReminderMinutesBeforeStart() {
		return Dispatch.get(this, "ReminderMinutesBeforeStart").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param reminderMinutesBeforeStart an input-parameter of type int
	 */
	public void setReminderMinutesBeforeStart(int reminderMinutesBeforeStart) {
		Dispatch.put(this, "ReminderMinutesBeforeStart", new Variant(reminderMinutesBeforeStart));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getReminderOverrideDefault() {
		return Dispatch.get(this, "ReminderOverrideDefault").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param reminderOverrideDefault an input-parameter of type boolean
	 */
	public void setReminderOverrideDefault(boolean reminderOverrideDefault) {
		Dispatch.put(this, "ReminderOverrideDefault", new Variant(reminderOverrideDefault));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getReminderPlaySound() {
		return Dispatch.get(this, "ReminderPlaySound").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param reminderPlaySound an input-parameter of type boolean
	 */
	public void setReminderPlaySound(boolean reminderPlaySound) {
		Dispatch.put(this, "ReminderPlaySound", new Variant(reminderPlaySound));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getReminderSet() {
		return Dispatch.get(this, "ReminderSet").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param reminderSet an input-parameter of type boolean
	 */
	public void setReminderSet(boolean reminderSet) {
		Dispatch.put(this, "ReminderSet", new Variant(reminderSet));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getReminderSoundFile() {
		return Dispatch.get(this, "ReminderSoundFile").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param reminderSoundFile an input-parameter of type String
	 */
	public void setReminderSoundFile(String reminderSoundFile) {
		Dispatch.put(this, "ReminderSoundFile", reminderSoundFile);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type java.util.Date
	 */
	public java.util.Date getReplyTime() {
		return Dispatch.get(this, "ReplyTime").getJavaDate();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param replyTime an input-parameter of type java.util.Date
	 */
	public void setReplyTime(java.util.Date replyTime) {
		Dispatch.put(this, "ReplyTime", new Variant(replyTime));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getRequiredAttendees() {
		return Dispatch.get(this, "RequiredAttendees").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param requiredAttendees an input-parameter of type String
	 */
	public void setRequiredAttendees(String requiredAttendees) {
		Dispatch.put(this, "RequiredAttendees", requiredAttendees);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getResources() {
		return Dispatch.get(this, "Resources").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param resources an input-parameter of type String
	 */
	public void setResources(String resources) {
		Dispatch.put(this, "Resources", resources);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getResponseRequested() {
		return Dispatch.get(this, "ResponseRequested").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param responseRequested an input-parameter of type boolean
	 */
	public void setResponseRequested(boolean responseRequested) {
		Dispatch.put(this, "ResponseRequested", new Variant(responseRequested));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getResponseStatus() {
		return Dispatch.get(this, "ResponseStatus").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type java.util.Date
	 */
	public java.util.Date getStart() {
		return Dispatch.get(this, "Start").getJavaDate();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param start an input-parameter of type java.util.Date
	 */
	public void setStart(java.util.Date start) {
		Dispatch.put(this, "Start", new Variant(start));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void clearRecurrencePattern() {
		Dispatch.call(this, "ClearRecurrencePattern");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type MailItem
	 */
	public MailItem forwardAsVcal() {
		return new MailItem(Dispatch.call(this, "ForwardAsVcal").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type RecurrencePattern
	 */
	public RecurrencePattern getRecurrencePattern() {
		return new RecurrencePattern(Dispatch.call(this, "GetRecurrencePattern").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param response an input-parameter of type int
	 * @param fNoUI an input-parameter of type Variant
	 * @param fAdditionalTextDialog an input-parameter of type Variant
	 * @return the result is of type MeetingItem
	 */
	public MeetingItem respond(int response, Variant fNoUI, Variant fAdditionalTextDialog) {
		return new MeetingItem(Dispatch.call(this, "Respond", new Variant(response), fNoUI, fAdditionalTextDialog).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param response an input-parameter of type int
	 * @param fNoUI an input-parameter of type Variant
	 * @return the result is of type MeetingItem
	 */
	public MeetingItem respond(int response, Variant fNoUI) {
		return new MeetingItem(Dispatch.call(this, "Respond", new Variant(response), fNoUI).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param response an input-parameter of type int
	 * @return the result is of type MeetingItem
	 */
	public MeetingItem respond(int response) {
		return new MeetingItem(Dispatch.call(this, "Respond", new Variant(response)).toDispatch());
	}

//	/**
//	 * Wrapper for calling the ActiveX-Method and receiving the output-parameter(s).
//	 * @param response an input-parameter of type int
//	 * @param fNoUI an input-parameter of type Variant
//	 * @param fAdditionalTextDialog an input-parameter of type Variant
//	 * @return the result is of type MeetingItem
//	 */
//	public MeetingItem respond(int response, Variant fNoUI, Variant fAdditionalTextDialog) {
//		MeetingItem result_of_Respond = new MeetingItem(Dispatch.call(this, "Respond", new Variant(response), fNoUI, fAdditionalTextDialog).toDispatch());
//
//
//		return result_of_Respond;
//	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void send() {
		Dispatch.call(this, "Send");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getNetMeetingDocPathName() {
		return Dispatch.get(this, "NetMeetingDocPathName").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param netMeetingDocPathName an input-parameter of type String
	 */
	public void setNetMeetingDocPathName(String netMeetingDocPathName) {
		Dispatch.put(this, "NetMeetingDocPathName", netMeetingDocPathName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getNetShowURL() {
		return Dispatch.get(this, "NetShowURL").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param netShowURL an input-parameter of type String
	 */
	public void setNetShowURL(String netShowURL) {
		Dispatch.put(this, "NetShowURL", netShowURL);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Links
	 */
	public Links getLinks() {
		return new Links(Dispatch.get(this, "Links").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getConferenceServerAllowExternal() {
		return Dispatch.get(this, "ConferenceServerAllowExternal").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param conferenceServerAllowExternal an input-parameter of type boolean
	 */
	public void setConferenceServerAllowExternal(boolean conferenceServerAllowExternal) {
		Dispatch.put(this, "ConferenceServerAllowExternal", new Variant(conferenceServerAllowExternal));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getConferenceServerPassword() {
		return Dispatch.get(this, "ConferenceServerPassword").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param conferenceServerPassword an input-parameter of type String
	 */
	public void setConferenceServerPassword(String conferenceServerPassword) {
		Dispatch.put(this, "ConferenceServerPassword", conferenceServerPassword);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type ItemProperties
	 */
	public ItemProperties getItemProperties() {
		return new ItemProperties(Dispatch.get(this, "ItemProperties").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getDownloadState() {
		return Dispatch.get(this, "DownloadState").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void showCategoriesDialog() {
		Dispatch.call(this, "ShowCategoriesDialog");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getInternetCodepage() {
		return Dispatch.get(this, "InternetCodepage").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param internetCodepage an input-parameter of type int
	 */
	public void setInternetCodepage(int internetCodepage) {
		Dispatch.put(this, "InternetCodepage", new Variant(internetCodepage));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getMarkForDownload() {
		return Dispatch.get(this, "MarkForDownload").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param markForDownload an input-parameter of type int
	 */
	public void setMarkForDownload(int markForDownload) {
		Dispatch.put(this, "MarkForDownload", new Variant(markForDownload));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getIsConflict() {
		return Dispatch.get(this, "IsConflict").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getMeetingWorkspaceURL() {
		return Dispatch.get(this, "MeetingWorkspaceURL").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getAutoResolvedWinner() {
		return Dispatch.get(this, "AutoResolvedWinner").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Conflicts
	 */
	public Conflicts getConflicts() {
		return new Conflicts(Dispatch.get(this, "Conflicts").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type PropertyAccessor
	 */
	public PropertyAccessor getPropertyAccessor() {
		return new PropertyAccessor(Dispatch.get(this, "PropertyAccessor").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Account
	 */
	public Account getSendUsingAccount() {
		return new Account(Dispatch.get(this, "SendUsingAccount").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param sendUsingAccount an input-parameter of type Account
	 */
	public void setSendUsingAccount(Account sendUsingAccount) {
		Dispatch.put(this, "SendUsingAccount", sendUsingAccount);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getGlobalAppointmentID() {
		return Dispatch.get(this, "GlobalAppointmentID").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getForceUpdateToAllAttendees() {
		return Dispatch.get(this, "ForceUpdateToAllAttendees").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param forceUpdateToAllAttendees an input-parameter of type boolean
	 */
	public void setForceUpdateToAllAttendees(boolean forceUpdateToAllAttendees) {
		Dispatch.put(this, "ForceUpdateToAllAttendees", new Variant(forceUpdateToAllAttendees));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type java.util.Date
	 */
	public java.util.Date getStartUTC() {
		return Dispatch.get(this, "StartUTC").getJavaDate();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param startUTC an input-parameter of type java.util.Date
	 */
	public void setStartUTC(java.util.Date startUTC) {
		Dispatch.put(this, "StartUTC", new Variant(startUTC));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type java.util.Date
	 */
	public java.util.Date getEndUTC() {
		return Dispatch.get(this, "EndUTC").getJavaDate();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param endUTC an input-parameter of type java.util.Date
	 */
	public void setEndUTC(java.util.Date endUTC) {
		Dispatch.put(this, "EndUTC", new Variant(endUTC));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type java.util.Date
	 */
	public java.util.Date getStartInStartTimeZone() {
		return Dispatch.get(this, "StartInStartTimeZone").getJavaDate();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param startInStartTimeZone an input-parameter of type java.util.Date
	 */
	public void setStartInStartTimeZone(java.util.Date startInStartTimeZone) {
		Dispatch.put(this, "StartInStartTimeZone", new Variant(startInStartTimeZone));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type java.util.Date
	 */
	public java.util.Date getEndInEndTimeZone() {
		return Dispatch.get(this, "EndInEndTimeZone").getJavaDate();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param endInEndTimeZone an input-parameter of type java.util.Date
	 */
	public void setEndInEndTimeZone(java.util.Date endInEndTimeZone) {
		Dispatch.put(this, "EndInEndTimeZone", new Variant(endInEndTimeZone));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type _TimeZone
	 */
	public _TimeZone getStartTimeZone() {
		return new _TimeZone(Dispatch.get(this, "StartTimeZone").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param startTimeZone an input-parameter of type _TimeZone
	 */
	public void setStartTimeZone(_TimeZone startTimeZone) {
		Dispatch.put(this, "StartTimeZone", startTimeZone);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type _TimeZone
	 */
	public _TimeZone getEndTimeZone() {
		return new _TimeZone(Dispatch.get(this, "EndTimeZone").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param endTimeZone an input-parameter of type _TimeZone
	 */
	public void setEndTimeZone(_TimeZone endTimeZone) {
		Dispatch.put(this, "EndTimeZone", endTimeZone);
	}

}
