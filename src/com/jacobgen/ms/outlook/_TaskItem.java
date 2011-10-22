/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class _TaskItem extends Dispatch {

	public static final String componentName = "Outlook._TaskItem";

	public _TaskItem() {
		super(componentName);
	}

	/**
	* This constructor is used instead of a case operation to
	* turn a Dispatch object into a wider object - it must exist
	* in every wrapper class whose instances may be returned from
	* method calls wrapped in VT_DISPATCH Variants.
	*/
	public _TaskItem(Dispatch d) {
		// take over the IDispatch pointer
		m_pDispatch = d.m_pDispatch;
		// null out the input's pointer
		d.m_pDispatch = 0;
	}

	public _TaskItem(String compName) {
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
	 * @return the result is of type int
	 */
	public int getActualWork() {
		return Dispatch.get(this, "ActualWork").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param actualWork an input-parameter of type int
	 */
	public void setActualWork(int actualWork) {
		Dispatch.put(this, "ActualWork", new Variant(actualWork));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getCardData() {
		return Dispatch.get(this, "CardData").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param cardData an input-parameter of type String
	 */
	public void setCardData(String cardData) {
		Dispatch.put(this, "CardData", cardData);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getComplete() {
		return Dispatch.get(this, "Complete").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param complete an input-parameter of type boolean
	 */
	public void setComplete(boolean complete) {
		Dispatch.put(this, "Complete", new Variant(complete));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getContacts() {
		return Dispatch.get(this, "Contacts").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param contacts an input-parameter of type String
	 */
	public void setContacts(String contacts) {
		Dispatch.put(this, "Contacts", contacts);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getContactNames() {
		return Dispatch.get(this, "ContactNames").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param contactNames an input-parameter of type String
	 */
	public void setContactNames(String contactNames) {
		Dispatch.put(this, "ContactNames", contactNames);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type java.util.Date
	 */
	public java.util.Date getDateCompleted() {
		return Dispatch.get(this, "DateCompleted").getJavaDate();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param dateCompleted an input-parameter of type java.util.Date
	 */
	public void setDateCompleted(java.util.Date dateCompleted) {
		Dispatch.put(this, "DateCompleted", new Variant(dateCompleted));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getDelegationState() {
		return Dispatch.get(this, "DelegationState").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getDelegator() {
		return Dispatch.get(this, "Delegator").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type java.util.Date
	 */
	public java.util.Date getDueDate() {
		return Dispatch.get(this, "DueDate").getJavaDate();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param dueDate an input-parameter of type java.util.Date
	 */
	public void setDueDate(java.util.Date dueDate) {
		Dispatch.put(this, "DueDate", new Variant(dueDate));
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
	 * @return the result is of type int
	 */
	public int getOrdinal() {
		return Dispatch.get(this, "Ordinal").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param ordinal an input-parameter of type int
	 */
	public void setOrdinal(int ordinal) {
		Dispatch.put(this, "Ordinal", new Variant(ordinal));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getOwner() {
		return Dispatch.get(this, "Owner").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param owner an input-parameter of type String
	 */
	public void setOwner(String owner) {
		Dispatch.put(this, "Owner", owner);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getOwnership() {
		return Dispatch.get(this, "Ownership").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getPercentComplete() {
		return Dispatch.get(this, "PercentComplete").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param percentComplete an input-parameter of type int
	 */
	public void setPercentComplete(int percentComplete) {
		Dispatch.put(this, "PercentComplete", new Variant(percentComplete));
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
	 * @return the result is of type int
	 */
	public int getResponseState() {
		return Dispatch.get(this, "ResponseState").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getRole() {
		return Dispatch.get(this, "Role").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param role an input-parameter of type String
	 */
	public void setRole(String role) {
		Dispatch.put(this, "Role", role);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getSchedulePlusPriority() {
		return Dispatch.get(this, "SchedulePlusPriority").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param schedulePlusPriority an input-parameter of type String
	 */
	public void setSchedulePlusPriority(String schedulePlusPriority) {
		Dispatch.put(this, "SchedulePlusPriority", schedulePlusPriority);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type java.util.Date
	 */
	public java.util.Date getStartDate() {
		return Dispatch.get(this, "StartDate").getJavaDate();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param startDate an input-parameter of type java.util.Date
	 */
	public void setStartDate(java.util.Date startDate) {
		Dispatch.put(this, "StartDate", new Variant(startDate));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getStatus() {
		return Dispatch.get(this, "Status").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param status an input-parameter of type int
	 */
	public void setStatus(int status) {
		Dispatch.put(this, "Status", new Variant(status));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getStatusOnCompletionRecipients() {
		return Dispatch.get(this, "StatusOnCompletionRecipients").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param statusOnCompletionRecipients an input-parameter of type String
	 */
	public void setStatusOnCompletionRecipients(String statusOnCompletionRecipients) {
		Dispatch.put(this, "StatusOnCompletionRecipients", statusOnCompletionRecipients);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getStatusUpdateRecipients() {
		return Dispatch.get(this, "StatusUpdateRecipients").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param statusUpdateRecipients an input-parameter of type String
	 */
	public void setStatusUpdateRecipients(String statusUpdateRecipients) {
		Dispatch.put(this, "StatusUpdateRecipients", statusUpdateRecipients);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getTeamTask() {
		return Dispatch.get(this, "TeamTask").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param teamTask an input-parameter of type boolean
	 */
	public void setTeamTask(boolean teamTask) {
		Dispatch.put(this, "TeamTask", new Variant(teamTask));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getTotalWork() {
		return Dispatch.get(this, "TotalWork").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param totalWork an input-parameter of type int
	 */
	public void setTotalWork(int totalWork) {
		Dispatch.put(this, "TotalWork", new Variant(totalWork));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type TaskItem
	 */
	public TaskItem assign() {
		return new TaskItem(Dispatch.call(this, "Assign").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void cancelResponseState() {
		Dispatch.call(this, "CancelResponseState");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void clearRecurrencePattern() {
		Dispatch.call(this, "ClearRecurrencePattern");
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
	 */
	public void markComplete() {
		Dispatch.call(this, "MarkComplete");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param response an input-parameter of type int
	 * @param fNoUI an input-parameter of type Variant
	 * @param fAdditionalTextDialog an input-parameter of type Variant
	 * @return the result is of type TaskItem
	 */
	public TaskItem respond(int response, Variant fNoUI, Variant fAdditionalTextDialog) {
		return new TaskItem(Dispatch.call(this, "Respond", new Variant(response), fNoUI, fAdditionalTextDialog).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void send() {
		Dispatch.call(this, "Send");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean skipRecurrence() {
		return Dispatch.call(this, "SkipRecurrence").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Object
	 */
	public Object statusReport() {
		return Dispatch.call(this, "StatusReport");
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
	 * @return the result is of type java.util.Date
	 */
	public java.util.Date getToDoTaskOrdinal() {
		return Dispatch.get(this, "ToDoTaskOrdinal").getJavaDate();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param toDoTaskOrdinal an input-parameter of type java.util.Date
	 */
	public void setToDoTaskOrdinal(java.util.Date toDoTaskOrdinal) {
		Dispatch.put(this, "ToDoTaskOrdinal", new Variant(toDoTaskOrdinal));
	}

}
