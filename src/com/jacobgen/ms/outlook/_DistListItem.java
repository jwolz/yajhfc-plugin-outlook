/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class _DistListItem extends Dispatch {

	public static final String componentName = "Outlook._DistListItem";

	public _DistListItem() {
		super(componentName);
	}

	/**
	* This constructor is used instead of a case operation to
	* turn a Dispatch object into a wider object - it must exist
	* in every wrapper class whose instances may be returned from
	* method calls wrapped in VT_DISPATCH Variants.
	*/
	public _DistListItem(Dispatch d) {
		// take over the IDispatch pointer
		m_pDispatch = d.m_pDispatch;
		// null out the input's pointer
		d.m_pDispatch = 0;
	}

	public _DistListItem(String compName) {
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


	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getDLName() {
		return Dispatch.get(this, "DLName").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param dLName an input-parameter of type String
	 */
	public void setDLName(String dLName) {
		Dispatch.put(this, "DLName", dLName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getMemberCount() {
		return Dispatch.get(this, "MemberCount").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getCheckSum() {
		return Dispatch.get(this, "CheckSum").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Variant
	 */
	public Variant getMembers() {
		return Dispatch.get(this, "Members");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param members an input-parameter of type Variant
	 */
	public void setMembers(Variant members) {
		Dispatch.put(this, "Members", members);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Variant
	 */
	public Variant getOneOffMembers() {
		return Dispatch.get(this, "OneOffMembers");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param oneOffMembers an input-parameter of type Variant
	 */
	public void setOneOffMembers(Variant oneOffMembers) {
		Dispatch.put(this, "OneOffMembers", oneOffMembers);
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
	 * @param recipients an input-parameter of type Recipients
	 */
	public void addMembers(Recipients recipients) {
		Dispatch.call(this, "AddMembers", recipients);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param recipients an input-parameter of type Recipients
	 */
	public void removeMembers(Recipients recipients) {
		Dispatch.call(this, "RemoveMembers", recipients);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param index an input-parameter of type int
	 * @return the result is of type Recipient
	 */
	public Recipient getMember(int index) {
		return new Recipient(Dispatch.call(this, "GetMember", new Variant(index)).toDispatch());
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
	 * @param recipient an input-parameter of type Recipient
	 */
	public void addMember(Recipient recipient) {
		Dispatch.call(this, "AddMember", recipient);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param recipient an input-parameter of type Recipient
	 */
	public void removeMember(Recipient recipient) {
		Dispatch.call(this, "RemoveMember", recipient);
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
	 * @return the result is of type String
	 */
	public String getTaskSubject() {
		return Dispatch.get(this, "TaskSubject").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param taskSubject an input-parameter of type String
	 */
	public void setTaskSubject(String taskSubject) {
		Dispatch.put(this, "TaskSubject", taskSubject);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type java.util.Date
	 */
	public java.util.Date getTaskDueDate() {
		return Dispatch.get(this, "TaskDueDate").getJavaDate();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param taskDueDate an input-parameter of type java.util.Date
	 */
	public void setTaskDueDate(java.util.Date taskDueDate) {
		Dispatch.put(this, "TaskDueDate", new Variant(taskDueDate));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type java.util.Date
	 */
	public java.util.Date getTaskStartDate() {
		return Dispatch.get(this, "TaskStartDate").getJavaDate();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param taskStartDate an input-parameter of type java.util.Date
	 */
	public void setTaskStartDate(java.util.Date taskStartDate) {
		Dispatch.put(this, "TaskStartDate", new Variant(taskStartDate));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type java.util.Date
	 */
	public java.util.Date getTaskCompletedDate() {
		return Dispatch.get(this, "TaskCompletedDate").getJavaDate();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param taskCompletedDate an input-parameter of type java.util.Date
	 */
	public void setTaskCompletedDate(java.util.Date taskCompletedDate) {
		Dispatch.put(this, "TaskCompletedDate", new Variant(taskCompletedDate));
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
	 * @param markInterval an input-parameter of type int
	 */
	public void markAsTask(int markInterval) {
		Dispatch.call(this, "MarkAsTask", new Variant(markInterval));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void clearTaskFlag() {
		Dispatch.call(this, "ClearTaskFlag");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getIsMarkedAsTask() {
		return Dispatch.get(this, "IsMarkedAsTask").changeType(Variant.VariantBoolean).getBoolean();
	}

}
