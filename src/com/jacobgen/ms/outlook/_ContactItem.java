/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class _ContactItem extends Dispatch {

	public static final String componentName = "Outlook._ContactItem";

	public _ContactItem() {
		super(componentName);
	}

	/**
	* This constructor is used instead of a case operation to
	* turn a Dispatch object into a wider object - it must exist
	* in every wrapper class whose instances may be returned from
	* method calls wrapped in VT_DISPATCH Variants.
	*/
	public _ContactItem(Dispatch d) {
		// take over the IDispatch pointer
		m_pDispatch = d.m_pDispatch;
		// null out the input's pointer
		d.m_pDispatch = 0;
	}

	public _ContactItem(String compName) {
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
	 * @return the result is of type String
	 */
	public String getAccount() {
		return Dispatch.get(this, "Account").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param account an input-parameter of type String
	 */
	public void setAccount(String account) {
		Dispatch.put(this, "Account", account);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type java.util.Date
	 */
	public java.util.Date getAnniversary() {
		return Dispatch.get(this, "Anniversary").getJavaDate();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param anniversary an input-parameter of type java.util.Date
	 */
	public void setAnniversary(java.util.Date anniversary) {
		Dispatch.put(this, "Anniversary", new Variant(anniversary));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getAssistantName() {
		return Dispatch.get(this, "AssistantName").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param assistantName an input-parameter of type String
	 */
	public void setAssistantName(String assistantName) {
		Dispatch.put(this, "AssistantName", assistantName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getAssistantTelephoneNumber() {
		return Dispatch.get(this, "AssistantTelephoneNumber").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param assistantTelephoneNumber an input-parameter of type String
	 */
	public void setAssistantTelephoneNumber(String assistantTelephoneNumber) {
		Dispatch.put(this, "AssistantTelephoneNumber", assistantTelephoneNumber);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type java.util.Date
	 */
	public java.util.Date getBirthday() {
		return Dispatch.get(this, "Birthday").getJavaDate();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param birthday an input-parameter of type java.util.Date
	 */
	public void setBirthday(java.util.Date birthday) {
		Dispatch.put(this, "Birthday", new Variant(birthday));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getBusiness2TelephoneNumber() {
		return Dispatch.get(this, "Business2TelephoneNumber").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param business2TelephoneNumber an input-parameter of type String
	 */
	public void setBusiness2TelephoneNumber(String business2TelephoneNumber) {
		Dispatch.put(this, "Business2TelephoneNumber", business2TelephoneNumber);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getBusinessAddress() {
		return Dispatch.get(this, "BusinessAddress").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param businessAddress an input-parameter of type String
	 */
	public void setBusinessAddress(String businessAddress) {
		Dispatch.put(this, "BusinessAddress", businessAddress);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getBusinessAddressCity() {
		return Dispatch.get(this, "BusinessAddressCity").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param businessAddressCity an input-parameter of type String
	 */
	public void setBusinessAddressCity(String businessAddressCity) {
		Dispatch.put(this, "BusinessAddressCity", businessAddressCity);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getBusinessAddressCountry() {
		return Dispatch.get(this, "BusinessAddressCountry").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param businessAddressCountry an input-parameter of type String
	 */
	public void setBusinessAddressCountry(String businessAddressCountry) {
		Dispatch.put(this, "BusinessAddressCountry", businessAddressCountry);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getBusinessAddressPostalCode() {
		return Dispatch.get(this, "BusinessAddressPostalCode").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param businessAddressPostalCode an input-parameter of type String
	 */
	public void setBusinessAddressPostalCode(String businessAddressPostalCode) {
		Dispatch.put(this, "BusinessAddressPostalCode", businessAddressPostalCode);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getBusinessAddressPostOfficeBox() {
		return Dispatch.get(this, "BusinessAddressPostOfficeBox").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param businessAddressPostOfficeBox an input-parameter of type String
	 */
	public void setBusinessAddressPostOfficeBox(String businessAddressPostOfficeBox) {
		Dispatch.put(this, "BusinessAddressPostOfficeBox", businessAddressPostOfficeBox);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getBusinessAddressState() {
		return Dispatch.get(this, "BusinessAddressState").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param businessAddressState an input-parameter of type String
	 */
	public void setBusinessAddressState(String businessAddressState) {
		Dispatch.put(this, "BusinessAddressState", businessAddressState);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getBusinessAddressStreet() {
		return Dispatch.get(this, "BusinessAddressStreet").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param businessAddressStreet an input-parameter of type String
	 */
	public void setBusinessAddressStreet(String businessAddressStreet) {
		Dispatch.put(this, "BusinessAddressStreet", businessAddressStreet);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getBusinessFaxNumber() {
		return Dispatch.get(this, "BusinessFaxNumber").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param businessFaxNumber an input-parameter of type String
	 */
	public void setBusinessFaxNumber(String businessFaxNumber) {
		Dispatch.put(this, "BusinessFaxNumber", businessFaxNumber);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getBusinessHomePage() {
		return Dispatch.get(this, "BusinessHomePage").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param businessHomePage an input-parameter of type String
	 */
	public void setBusinessHomePage(String businessHomePage) {
		Dispatch.put(this, "BusinessHomePage", businessHomePage);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getBusinessTelephoneNumber() {
		return Dispatch.get(this, "BusinessTelephoneNumber").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param businessTelephoneNumber an input-parameter of type String
	 */
	public void setBusinessTelephoneNumber(String businessTelephoneNumber) {
		Dispatch.put(this, "BusinessTelephoneNumber", businessTelephoneNumber);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getCallbackTelephoneNumber() {
		return Dispatch.get(this, "CallbackTelephoneNumber").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param callbackTelephoneNumber an input-parameter of type String
	 */
	public void setCallbackTelephoneNumber(String callbackTelephoneNumber) {
		Dispatch.put(this, "CallbackTelephoneNumber", callbackTelephoneNumber);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getCarTelephoneNumber() {
		return Dispatch.get(this, "CarTelephoneNumber").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param carTelephoneNumber an input-parameter of type String
	 */
	public void setCarTelephoneNumber(String carTelephoneNumber) {
		Dispatch.put(this, "CarTelephoneNumber", carTelephoneNumber);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getChildren() {
		return Dispatch.get(this, "Children").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param children an input-parameter of type String
	 */
	public void setChildren(String children) {
		Dispatch.put(this, "Children", children);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getCompanyAndFullName() {
		return Dispatch.get(this, "CompanyAndFullName").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getCompanyLastFirstNoSpace() {
		return Dispatch.get(this, "CompanyLastFirstNoSpace").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getCompanyLastFirstSpaceOnly() {
		return Dispatch.get(this, "CompanyLastFirstSpaceOnly").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getCompanyMainTelephoneNumber() {
		return Dispatch.get(this, "CompanyMainTelephoneNumber").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param companyMainTelephoneNumber an input-parameter of type String
	 */
	public void setCompanyMainTelephoneNumber(String companyMainTelephoneNumber) {
		Dispatch.put(this, "CompanyMainTelephoneNumber", companyMainTelephoneNumber);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getCompanyName() {
		return Dispatch.get(this, "CompanyName").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param companyName an input-parameter of type String
	 */
	public void setCompanyName(String companyName) {
		Dispatch.put(this, "CompanyName", companyName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getComputerNetworkName() {
		return Dispatch.get(this, "ComputerNetworkName").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param computerNetworkName an input-parameter of type String
	 */
	public void setComputerNetworkName(String computerNetworkName) {
		Dispatch.put(this, "ComputerNetworkName", computerNetworkName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getCustomerID() {
		return Dispatch.get(this, "CustomerID").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param customerID an input-parameter of type String
	 */
	public void setCustomerID(String customerID) {
		Dispatch.put(this, "CustomerID", customerID);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getDepartment() {
		return Dispatch.get(this, "Department").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param department an input-parameter of type String
	 */
	public void setDepartment(String department) {
		Dispatch.put(this, "Department", department);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getEmail1Address() {
		return Dispatch.get(this, "Email1Address").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param email1Address an input-parameter of type String
	 */
	public void setEmail1Address(String email1Address) {
		Dispatch.put(this, "Email1Address", email1Address);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getEmail1AddressType() {
		return Dispatch.get(this, "Email1AddressType").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param email1AddressType an input-parameter of type String
	 */
	public void setEmail1AddressType(String email1AddressType) {
		Dispatch.put(this, "Email1AddressType", email1AddressType);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getEmail1DisplayName() {
		return Dispatch.get(this, "Email1DisplayName").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getEmail1EntryID() {
		return Dispatch.get(this, "Email1EntryID").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getEmail2Address() {
		return Dispatch.get(this, "Email2Address").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param email2Address an input-parameter of type String
	 */
	public void setEmail2Address(String email2Address) {
		Dispatch.put(this, "Email2Address", email2Address);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getEmail2AddressType() {
		return Dispatch.get(this, "Email2AddressType").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param email2AddressType an input-parameter of type String
	 */
	public void setEmail2AddressType(String email2AddressType) {
		Dispatch.put(this, "Email2AddressType", email2AddressType);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getEmail2DisplayName() {
		return Dispatch.get(this, "Email2DisplayName").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getEmail2EntryID() {
		return Dispatch.get(this, "Email2EntryID").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getEmail3Address() {
		return Dispatch.get(this, "Email3Address").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param email3Address an input-parameter of type String
	 */
	public void setEmail3Address(String email3Address) {
		Dispatch.put(this, "Email3Address", email3Address);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getEmail3AddressType() {
		return Dispatch.get(this, "Email3AddressType").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param email3AddressType an input-parameter of type String
	 */
	public void setEmail3AddressType(String email3AddressType) {
		Dispatch.put(this, "Email3AddressType", email3AddressType);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getEmail3DisplayName() {
		return Dispatch.get(this, "Email3DisplayName").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getEmail3EntryID() {
		return Dispatch.get(this, "Email3EntryID").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getFileAs() {
		return Dispatch.get(this, "FileAs").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param fileAs an input-parameter of type String
	 */
	public void setFileAs(String fileAs) {
		Dispatch.put(this, "FileAs", fileAs);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getFirstName() {
		return Dispatch.get(this, "FirstName").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param firstName an input-parameter of type String
	 */
	public void setFirstName(String firstName) {
		Dispatch.put(this, "FirstName", firstName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getFTPSite() {
		return Dispatch.get(this, "FTPSite").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param fTPSite an input-parameter of type String
	 */
	public void setFTPSite(String fTPSite) {
		Dispatch.put(this, "FTPSite", fTPSite);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getFullName() {
		return Dispatch.get(this, "FullName").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param fullName an input-parameter of type String
	 */
	public void setFullName(String fullName) {
		Dispatch.put(this, "FullName", fullName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getFullNameAndCompany() {
		return Dispatch.get(this, "FullNameAndCompany").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getGender() {
		return Dispatch.get(this, "Gender").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param gender an input-parameter of type int
	 */
	public void setGender(int gender) {
		Dispatch.put(this, "Gender", new Variant(gender));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getGovernmentIDNumber() {
		return Dispatch.get(this, "GovernmentIDNumber").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param governmentIDNumber an input-parameter of type String
	 */
	public void setGovernmentIDNumber(String governmentIDNumber) {
		Dispatch.put(this, "GovernmentIDNumber", governmentIDNumber);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getHobby() {
		return Dispatch.get(this, "Hobby").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param hobby an input-parameter of type String
	 */
	public void setHobby(String hobby) {
		Dispatch.put(this, "Hobby", hobby);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getHome2TelephoneNumber() {
		return Dispatch.get(this, "Home2TelephoneNumber").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param home2TelephoneNumber an input-parameter of type String
	 */
	public void setHome2TelephoneNumber(String home2TelephoneNumber) {
		Dispatch.put(this, "Home2TelephoneNumber", home2TelephoneNumber);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getHomeAddress() {
		return Dispatch.get(this, "HomeAddress").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param homeAddress an input-parameter of type String
	 */
	public void setHomeAddress(String homeAddress) {
		Dispatch.put(this, "HomeAddress", homeAddress);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getHomeAddressCity() {
		return Dispatch.get(this, "HomeAddressCity").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param homeAddressCity an input-parameter of type String
	 */
	public void setHomeAddressCity(String homeAddressCity) {
		Dispatch.put(this, "HomeAddressCity", homeAddressCity);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getHomeAddressCountry() {
		return Dispatch.get(this, "HomeAddressCountry").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param homeAddressCountry an input-parameter of type String
	 */
	public void setHomeAddressCountry(String homeAddressCountry) {
		Dispatch.put(this, "HomeAddressCountry", homeAddressCountry);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getHomeAddressPostalCode() {
		return Dispatch.get(this, "HomeAddressPostalCode").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param homeAddressPostalCode an input-parameter of type String
	 */
	public void setHomeAddressPostalCode(String homeAddressPostalCode) {
		Dispatch.put(this, "HomeAddressPostalCode", homeAddressPostalCode);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getHomeAddressPostOfficeBox() {
		return Dispatch.get(this, "HomeAddressPostOfficeBox").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param homeAddressPostOfficeBox an input-parameter of type String
	 */
	public void setHomeAddressPostOfficeBox(String homeAddressPostOfficeBox) {
		Dispatch.put(this, "HomeAddressPostOfficeBox", homeAddressPostOfficeBox);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getHomeAddressState() {
		return Dispatch.get(this, "HomeAddressState").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param homeAddressState an input-parameter of type String
	 */
	public void setHomeAddressState(String homeAddressState) {
		Dispatch.put(this, "HomeAddressState", homeAddressState);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getHomeAddressStreet() {
		return Dispatch.get(this, "HomeAddressStreet").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param homeAddressStreet an input-parameter of type String
	 */
	public void setHomeAddressStreet(String homeAddressStreet) {
		Dispatch.put(this, "HomeAddressStreet", homeAddressStreet);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getHomeFaxNumber() {
		return Dispatch.get(this, "HomeFaxNumber").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param homeFaxNumber an input-parameter of type String
	 */
	public void setHomeFaxNumber(String homeFaxNumber) {
		Dispatch.put(this, "HomeFaxNumber", homeFaxNumber);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getHomeTelephoneNumber() {
		return Dispatch.get(this, "HomeTelephoneNumber").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param homeTelephoneNumber an input-parameter of type String
	 */
	public void setHomeTelephoneNumber(String homeTelephoneNumber) {
		Dispatch.put(this, "HomeTelephoneNumber", homeTelephoneNumber);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getInitials() {
		return Dispatch.get(this, "Initials").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param initials an input-parameter of type String
	 */
	public void setInitials(String initials) {
		Dispatch.put(this, "Initials", initials);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getInternetFreeBusyAddress() {
		return Dispatch.get(this, "InternetFreeBusyAddress").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param internetFreeBusyAddress an input-parameter of type String
	 */
	public void setInternetFreeBusyAddress(String internetFreeBusyAddress) {
		Dispatch.put(this, "InternetFreeBusyAddress", internetFreeBusyAddress);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getISDNNumber() {
		return Dispatch.get(this, "ISDNNumber").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param iSDNNumber an input-parameter of type String
	 */
	public void setISDNNumber(String iSDNNumber) {
		Dispatch.put(this, "ISDNNumber", iSDNNumber);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getJobTitle() {
		return Dispatch.get(this, "JobTitle").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param jobTitle an input-parameter of type String
	 */
	public void setJobTitle(String jobTitle) {
		Dispatch.put(this, "JobTitle", jobTitle);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getJournal() {
		return Dispatch.get(this, "Journal").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param journal an input-parameter of type boolean
	 */
	public void setJournal(boolean journal) {
		Dispatch.put(this, "Journal", new Variant(journal));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getLanguage() {
		return Dispatch.get(this, "Language").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param language an input-parameter of type String
	 */
	public void setLanguage(String language) {
		Dispatch.put(this, "Language", language);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getLastFirstAndSuffix() {
		return Dispatch.get(this, "LastFirstAndSuffix").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getLastFirstNoSpace() {
		return Dispatch.get(this, "LastFirstNoSpace").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getLastFirstNoSpaceCompany() {
		return Dispatch.get(this, "LastFirstNoSpaceCompany").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getLastFirstSpaceOnly() {
		return Dispatch.get(this, "LastFirstSpaceOnly").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getLastFirstSpaceOnlyCompany() {
		return Dispatch.get(this, "LastFirstSpaceOnlyCompany").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getLastName() {
		return Dispatch.get(this, "LastName").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param lastName an input-parameter of type String
	 */
	public void setLastName(String lastName) {
		Dispatch.put(this, "LastName", lastName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getLastNameAndFirstName() {
		return Dispatch.get(this, "LastNameAndFirstName").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getMailingAddress() {
		return Dispatch.get(this, "MailingAddress").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param mailingAddress an input-parameter of type String
	 */
	public void setMailingAddress(String mailingAddress) {
		Dispatch.put(this, "MailingAddress", mailingAddress);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getMailingAddressCity() {
		return Dispatch.get(this, "MailingAddressCity").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param mailingAddressCity an input-parameter of type String
	 */
	public void setMailingAddressCity(String mailingAddressCity) {
		Dispatch.put(this, "MailingAddressCity", mailingAddressCity);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getMailingAddressCountry() {
		return Dispatch.get(this, "MailingAddressCountry").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param mailingAddressCountry an input-parameter of type String
	 */
	public void setMailingAddressCountry(String mailingAddressCountry) {
		Dispatch.put(this, "MailingAddressCountry", mailingAddressCountry);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getMailingAddressPostalCode() {
		return Dispatch.get(this, "MailingAddressPostalCode").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param mailingAddressPostalCode an input-parameter of type String
	 */
	public void setMailingAddressPostalCode(String mailingAddressPostalCode) {
		Dispatch.put(this, "MailingAddressPostalCode", mailingAddressPostalCode);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getMailingAddressPostOfficeBox() {
		return Dispatch.get(this, "MailingAddressPostOfficeBox").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param mailingAddressPostOfficeBox an input-parameter of type String
	 */
	public void setMailingAddressPostOfficeBox(String mailingAddressPostOfficeBox) {
		Dispatch.put(this, "MailingAddressPostOfficeBox", mailingAddressPostOfficeBox);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getMailingAddressState() {
		return Dispatch.get(this, "MailingAddressState").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param mailingAddressState an input-parameter of type String
	 */
	public void setMailingAddressState(String mailingAddressState) {
		Dispatch.put(this, "MailingAddressState", mailingAddressState);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getMailingAddressStreet() {
		return Dispatch.get(this, "MailingAddressStreet").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param mailingAddressStreet an input-parameter of type String
	 */
	public void setMailingAddressStreet(String mailingAddressStreet) {
		Dispatch.put(this, "MailingAddressStreet", mailingAddressStreet);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getManagerName() {
		return Dispatch.get(this, "ManagerName").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param managerName an input-parameter of type String
	 */
	public void setManagerName(String managerName) {
		Dispatch.put(this, "ManagerName", managerName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getMiddleName() {
		return Dispatch.get(this, "MiddleName").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param middleName an input-parameter of type String
	 */
	public void setMiddleName(String middleName) {
		Dispatch.put(this, "MiddleName", middleName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getMobileTelephoneNumber() {
		return Dispatch.get(this, "MobileTelephoneNumber").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param mobileTelephoneNumber an input-parameter of type String
	 */
	public void setMobileTelephoneNumber(String mobileTelephoneNumber) {
		Dispatch.put(this, "MobileTelephoneNumber", mobileTelephoneNumber);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getNetMeetingAlias() {
		return Dispatch.get(this, "NetMeetingAlias").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param netMeetingAlias an input-parameter of type String
	 */
	public void setNetMeetingAlias(String netMeetingAlias) {
		Dispatch.put(this, "NetMeetingAlias", netMeetingAlias);
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
	 * @return the result is of type String
	 */
	public String getNickName() {
		return Dispatch.get(this, "NickName").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param nickName an input-parameter of type String
	 */
	public void setNickName(String nickName) {
		Dispatch.put(this, "NickName", nickName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getOfficeLocation() {
		return Dispatch.get(this, "OfficeLocation").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param officeLocation an input-parameter of type String
	 */
	public void setOfficeLocation(String officeLocation) {
		Dispatch.put(this, "OfficeLocation", officeLocation);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getOrganizationalIDNumber() {
		return Dispatch.get(this, "OrganizationalIDNumber").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param organizationalIDNumber an input-parameter of type String
	 */
	public void setOrganizationalIDNumber(String organizationalIDNumber) {
		Dispatch.put(this, "OrganizationalIDNumber", organizationalIDNumber);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getOtherAddress() {
		return Dispatch.get(this, "OtherAddress").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param otherAddress an input-parameter of type String
	 */
	public void setOtherAddress(String otherAddress) {
		Dispatch.put(this, "OtherAddress", otherAddress);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getOtherAddressCity() {
		return Dispatch.get(this, "OtherAddressCity").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param otherAddressCity an input-parameter of type String
	 */
	public void setOtherAddressCity(String otherAddressCity) {
		Dispatch.put(this, "OtherAddressCity", otherAddressCity);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getOtherAddressCountry() {
		return Dispatch.get(this, "OtherAddressCountry").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param otherAddressCountry an input-parameter of type String
	 */
	public void setOtherAddressCountry(String otherAddressCountry) {
		Dispatch.put(this, "OtherAddressCountry", otherAddressCountry);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getOtherAddressPostalCode() {
		return Dispatch.get(this, "OtherAddressPostalCode").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param otherAddressPostalCode an input-parameter of type String
	 */
	public void setOtherAddressPostalCode(String otherAddressPostalCode) {
		Dispatch.put(this, "OtherAddressPostalCode", otherAddressPostalCode);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getOtherAddressPostOfficeBox() {
		return Dispatch.get(this, "OtherAddressPostOfficeBox").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param otherAddressPostOfficeBox an input-parameter of type String
	 */
	public void setOtherAddressPostOfficeBox(String otherAddressPostOfficeBox) {
		Dispatch.put(this, "OtherAddressPostOfficeBox", otherAddressPostOfficeBox);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getOtherAddressState() {
		return Dispatch.get(this, "OtherAddressState").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param otherAddressState an input-parameter of type String
	 */
	public void setOtherAddressState(String otherAddressState) {
		Dispatch.put(this, "OtherAddressState", otherAddressState);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getOtherAddressStreet() {
		return Dispatch.get(this, "OtherAddressStreet").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param otherAddressStreet an input-parameter of type String
	 */
	public void setOtherAddressStreet(String otherAddressStreet) {
		Dispatch.put(this, "OtherAddressStreet", otherAddressStreet);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getOtherFaxNumber() {
		return Dispatch.get(this, "OtherFaxNumber").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param otherFaxNumber an input-parameter of type String
	 */
	public void setOtherFaxNumber(String otherFaxNumber) {
		Dispatch.put(this, "OtherFaxNumber", otherFaxNumber);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getOtherTelephoneNumber() {
		return Dispatch.get(this, "OtherTelephoneNumber").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param otherTelephoneNumber an input-parameter of type String
	 */
	public void setOtherTelephoneNumber(String otherTelephoneNumber) {
		Dispatch.put(this, "OtherTelephoneNumber", otherTelephoneNumber);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getPagerNumber() {
		return Dispatch.get(this, "PagerNumber").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param pagerNumber an input-parameter of type String
	 */
	public void setPagerNumber(String pagerNumber) {
		Dispatch.put(this, "PagerNumber", pagerNumber);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getPersonalHomePage() {
		return Dispatch.get(this, "PersonalHomePage").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param personalHomePage an input-parameter of type String
	 */
	public void setPersonalHomePage(String personalHomePage) {
		Dispatch.put(this, "PersonalHomePage", personalHomePage);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getPrimaryTelephoneNumber() {
		return Dispatch.get(this, "PrimaryTelephoneNumber").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param primaryTelephoneNumber an input-parameter of type String
	 */
	public void setPrimaryTelephoneNumber(String primaryTelephoneNumber) {
		Dispatch.put(this, "PrimaryTelephoneNumber", primaryTelephoneNumber);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getProfession() {
		return Dispatch.get(this, "Profession").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param profession an input-parameter of type String
	 */
	public void setProfession(String profession) {
		Dispatch.put(this, "Profession", profession);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getRadioTelephoneNumber() {
		return Dispatch.get(this, "RadioTelephoneNumber").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param radioTelephoneNumber an input-parameter of type String
	 */
	public void setRadioTelephoneNumber(String radioTelephoneNumber) {
		Dispatch.put(this, "RadioTelephoneNumber", radioTelephoneNumber);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getReferredBy() {
		return Dispatch.get(this, "ReferredBy").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param referredBy an input-parameter of type String
	 */
	public void setReferredBy(String referredBy) {
		Dispatch.put(this, "ReferredBy", referredBy);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getSelectedMailingAddress() {
		return Dispatch.get(this, "SelectedMailingAddress").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param selectedMailingAddress an input-parameter of type int
	 */
	public void setSelectedMailingAddress(int selectedMailingAddress) {
		Dispatch.put(this, "SelectedMailingAddress", new Variant(selectedMailingAddress));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getSpouse() {
		return Dispatch.get(this, "Spouse").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param spouse an input-parameter of type String
	 */
	public void setSpouse(String spouse) {
		Dispatch.put(this, "Spouse", spouse);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getSuffix() {
		return Dispatch.get(this, "Suffix").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param suffix an input-parameter of type String
	 */
	public void setSuffix(String suffix) {
		Dispatch.put(this, "Suffix", suffix);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getTelexNumber() {
		return Dispatch.get(this, "TelexNumber").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param telexNumber an input-parameter of type String
	 */
	public void setTelexNumber(String telexNumber) {
		Dispatch.put(this, "TelexNumber", telexNumber);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getTitle() {
		return Dispatch.get(this, "Title").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param title an input-parameter of type String
	 */
	public void setTitle(String title) {
		Dispatch.put(this, "Title", title);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getTTYTDDTelephoneNumber() {
		return Dispatch.get(this, "TTYTDDTelephoneNumber").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param tTYTDDTelephoneNumber an input-parameter of type String
	 */
	public void setTTYTDDTelephoneNumber(String tTYTDDTelephoneNumber) {
		Dispatch.put(this, "TTYTDDTelephoneNumber", tTYTDDTelephoneNumber);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getUser1() {
		return Dispatch.get(this, "User1").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param user1 an input-parameter of type String
	 */
	public void setUser1(String user1) {
		Dispatch.put(this, "User1", user1);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getUser2() {
		return Dispatch.get(this, "User2").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param user2 an input-parameter of type String
	 */
	public void setUser2(String user2) {
		Dispatch.put(this, "User2", user2);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getUser3() {
		return Dispatch.get(this, "User3").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param user3 an input-parameter of type String
	 */
	public void setUser3(String user3) {
		Dispatch.put(this, "User3", user3);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getUser4() {
		return Dispatch.get(this, "User4").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param user4 an input-parameter of type String
	 */
	public void setUser4(String user4) {
		Dispatch.put(this, "User4", user4);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getUserCertificate() {
		return Dispatch.get(this, "UserCertificate").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param userCertificate an input-parameter of type String
	 */
	public void setUserCertificate(String userCertificate) {
		Dispatch.put(this, "UserCertificate", userCertificate);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getWebPage() {
		return Dispatch.get(this, "WebPage").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param webPage an input-parameter of type String
	 */
	public void setWebPage(String webPage) {
		Dispatch.put(this, "WebPage", webPage);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getYomiCompanyName() {
		return Dispatch.get(this, "YomiCompanyName").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param yomiCompanyName an input-parameter of type String
	 */
	public void setYomiCompanyName(String yomiCompanyName) {
		Dispatch.put(this, "YomiCompanyName", yomiCompanyName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getYomiFirstName() {
		return Dispatch.get(this, "YomiFirstName").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param yomiFirstName an input-parameter of type String
	 */
	public void setYomiFirstName(String yomiFirstName) {
		Dispatch.put(this, "YomiFirstName", yomiFirstName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getYomiLastName() {
		return Dispatch.get(this, "YomiLastName").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param yomiLastName an input-parameter of type String
	 */
	public void setYomiLastName(String yomiLastName) {
		Dispatch.put(this, "YomiLastName", yomiLastName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type MailItem
	 */
	public MailItem forwardAsVcard() {
		return new MailItem(Dispatch.call(this, "ForwardAsVcard").toDispatch());
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
	 * @return the result is of type ItemProperties
	 */
	public ItemProperties getItemProperties() {
		return new ItemProperties(Dispatch.get(this, "ItemProperties").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getLastFirstNoSpaceAndSuffix() {
		return Dispatch.get(this, "LastFirstNoSpaceAndSuffix").toString();
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
	 * @return the result is of type String
	 */
	public String getIMAddress() {
		return Dispatch.get(this, "IMAddress").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param iMAddress an input-parameter of type String
	 */
	public void setIMAddress(String iMAddress) {
		Dispatch.put(this, "IMAddress", iMAddress);
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
	 * @param email1DisplayName an input-parameter of type String
	 */
	public void setEmail1DisplayName(String email1DisplayName) {
		Dispatch.put(this, "Email1DisplayName", email1DisplayName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param email2DisplayName an input-parameter of type String
	 */
	public void setEmail2DisplayName(String email2DisplayName) {
		Dispatch.put(this, "Email2DisplayName", email2DisplayName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param email3DisplayName an input-parameter of type String
	 */
	public void setEmail3DisplayName(String email3DisplayName) {
		Dispatch.put(this, "Email3DisplayName", email3DisplayName);
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
	 * @param path an input-parameter of type String
	 */
	public void addPicture(String path) {
		Dispatch.call(this, "AddPicture", path);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void removePicture() {
		Dispatch.call(this, "RemovePicture");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getHasPicture() {
		return Dispatch.get(this, "HasPicture").changeType(Variant.VariantBoolean).getBoolean();
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
	 * @return the result is of type MailItem
	 */
	public MailItem forwardAsBusinessCard() {
		return new MailItem(Dispatch.call(this, "ForwardAsBusinessCard").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void showBusinessCardEditor() {
		Dispatch.call(this, "ShowBusinessCardEditor");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param path an input-parameter of type String
	 */
	public void saveBusinessCardImage(String path) {
		Dispatch.call(this, "SaveBusinessCardImage", path);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param phoneNumber an input-parameter of type int
	 */
	public void showCheckPhoneDialog(int phoneNumber) {
		Dispatch.call(this, "ShowCheckPhoneDialog", new Variant(phoneNumber));
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

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getBusinessCardLayoutXml() {
		return Dispatch.get(this, "BusinessCardLayoutXml").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param businessCardLayoutXml an input-parameter of type String
	 */
	public void setBusinessCardLayoutXml(String businessCardLayoutXml) {
		Dispatch.put(this, "BusinessCardLayoutXml", businessCardLayoutXml);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void resetBusinessCard() {
		Dispatch.call(this, "ResetBusinessCard");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param path an input-parameter of type String
	 */
	public void addBusinessCardLogoPicture(String path) {
		Dispatch.call(this, "AddBusinessCardLogoPicture", path);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getBusinessCardType() {
		return Dispatch.get(this, "BusinessCardType").changeType(Variant.VariantInt).getInt();
	}

}
