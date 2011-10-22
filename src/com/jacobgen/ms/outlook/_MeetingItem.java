/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class _MeetingItem extends Dispatch {

	public static final String componentName = "Outlook._MeetingItem";

	public _MeetingItem() {
		super(componentName);
	}

	/**
	* This constructor is used instead of a case operation to
	* turn a Dispatch object into a wider object - it must exist
	* in every wrapper class whose instances may be returned from
	* method calls wrapped in VT_DISPATCH Variants.
	*/
	public _MeetingItem(Dispatch d) {
		// take over the IDispatch pointer
		m_pDispatch = d.m_pDispatch;
		// null out the input's pointer
		d.m_pDispatch = 0;
	}

	public _MeetingItem(String compName) {
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

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getAutoForwarded() {
		return Dispatch.get(this, "AutoForwarded").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param autoForwarded an input-parameter of type boolean
	 */
	public void setAutoForwarded(boolean autoForwarded) {
		Dispatch.put(this, "AutoForwarded", new Variant(autoForwarded));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type java.util.Date
	 */
	public java.util.Date getDeferredDeliveryTime() {
		return Dispatch.get(this, "DeferredDeliveryTime").getJavaDate();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param deferredDeliveryTime an input-parameter of type java.util.Date
	 */
	public void setDeferredDeliveryTime(java.util.Date deferredDeliveryTime) {
		Dispatch.put(this, "DeferredDeliveryTime", new Variant(deferredDeliveryTime));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getDeleteAfterSubmit() {
		return Dispatch.get(this, "DeleteAfterSubmit").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param deleteAfterSubmit an input-parameter of type boolean
	 */
	public void setDeleteAfterSubmit(boolean deleteAfterSubmit) {
		Dispatch.put(this, "DeleteAfterSubmit", new Variant(deleteAfterSubmit));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type java.util.Date
	 */
	public java.util.Date getExpiryTime() {
		return Dispatch.get(this, "ExpiryTime").getJavaDate();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param expiryTime an input-parameter of type java.util.Date
	 */
	public void setExpiryTime(java.util.Date expiryTime) {
		Dispatch.put(this, "ExpiryTime", new Variant(expiryTime));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type java.util.Date
	 */
	public java.util.Date getFlagDueBy() {
		return Dispatch.get(this, "FlagDueBy").getJavaDate();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param flagDueBy an input-parameter of type java.util.Date
	 */
	public void setFlagDueBy(java.util.Date flagDueBy) {
		Dispatch.put(this, "FlagDueBy", new Variant(flagDueBy));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getFlagRequest() {
		return Dispatch.get(this, "FlagRequest").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param flagRequest an input-parameter of type String
	 */
	public void setFlagRequest(String flagRequest) {
		Dispatch.put(this, "FlagRequest", flagRequest);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getFlagStatus() {
		return Dispatch.get(this, "FlagStatus").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param flagStatus an input-parameter of type int
	 */
	public void setFlagStatus(int flagStatus) {
		Dispatch.put(this, "FlagStatus", new Variant(flagStatus));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getOriginatorDeliveryReportRequested() {
		return Dispatch.get(this, "OriginatorDeliveryReportRequested").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param originatorDeliveryReportRequested an input-parameter of type boolean
	 */
	public void setOriginatorDeliveryReportRequested(boolean originatorDeliveryReportRequested) {
		Dispatch.put(this, "OriginatorDeliveryReportRequested", new Variant(originatorDeliveryReportRequested));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type java.util.Date
	 */
	public java.util.Date getReceivedTime() {
		return Dispatch.get(this, "ReceivedTime").getJavaDate();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param receivedTime an input-parameter of type java.util.Date
	 */
	public void setReceivedTime(java.util.Date receivedTime) {
		Dispatch.put(this, "ReceivedTime", new Variant(receivedTime));
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
	 * @return the result is of type java.util.Date
	 */
	public java.util.Date getReminderTime() {
		return Dispatch.get(this, "ReminderTime").getJavaDate();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param reminderTime an input-parameter of type java.util.Date
	 */
	public void setReminderTime(java.util.Date reminderTime) {
		Dispatch.put(this, "ReminderTime", new Variant(reminderTime));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Recipients
	 */
	public Recipients getReplyRecipients() {
		return new Recipients(Dispatch.get(this, "ReplyRecipients").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type MAPIFolder
	 */
	public MAPIFolder getSaveSentMessageFolder() {
		return new MAPIFolder(Dispatch.get(this, "SaveSentMessageFolder").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param saveSentMessageFolder an input-parameter of type MAPIFolder
	 */
	public void setSaveSentMessageFolder(MAPIFolder saveSentMessageFolder) {
		Dispatch.put(this, "SaveSentMessageFolder", saveSentMessageFolder);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getSenderName() {
		return Dispatch.get(this, "SenderName").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getSent() {
		return Dispatch.get(this, "Sent").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type java.util.Date
	 */
	public java.util.Date getSentOn() {
		return Dispatch.get(this, "SentOn").getJavaDate();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getSubmitted() {
		return Dispatch.get(this, "Submitted").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type MeetingItem
	 */
	public MeetingItem forward() {
		return new MeetingItem(Dispatch.call(this, "Forward").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param addToCalendar an input-parameter of type boolean
	 * @return the result is of type AppointmentItem
	 */
	public AppointmentItem getAssociatedAppointment(boolean addToCalendar) {
		return new AppointmentItem(Dispatch.call(this, "GetAssociatedAppointment", new Variant(addToCalendar)).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type MailItem
	 */
	public MailItem reply() {
		return new MailItem(Dispatch.call(this, "Reply").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type MailItem
	 */
	public MailItem replyAll() {
		return new MailItem(Dispatch.call(this, "ReplyAll").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void send() {
		Dispatch.call(this, "Send");
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
	 * @return the result is of type ItemProperties
	 */
	public ItemProperties getItemProperties() {
		return new ItemProperties(Dispatch.get(this, "ItemProperties").toDispatch());
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
	 * @return the result is of type int
	 */
	public int getFlagIcon() {
		return Dispatch.get(this, "FlagIcon").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param flagIcon an input-parameter of type int
	 */
	public void setFlagIcon(int flagIcon) {
		Dispatch.put(this, "FlagIcon", new Variant(flagIcon));
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
	 * @return the result is of type String
	 */
	public String getSenderEmailAddress() {
		return Dispatch.get(this, "SenderEmailAddress").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getSenderEmailType() {
		return Dispatch.get(this, "SenderEmailType").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type PropertyAccessor
	 */
	public PropertyAccessor getPropertyAccessor() {
		return new PropertyAccessor(Dispatch.get(this, "PropertyAccessor").toDispatch());
	}

}
