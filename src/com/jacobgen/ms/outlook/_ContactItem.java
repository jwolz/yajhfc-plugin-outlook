/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class _ContactItem extends CachingDispatch {

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
		return new _Application(Dispatch.get(this, getIDOfName("Application")).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getClass1() {
		return Dispatch.get(this, getIDOfName("Class")).changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type _NameSpace
	 */
	public _NameSpace getSession() {
		return new _NameSpace(Dispatch.get(this, getIDOfName("Session")).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Object
	 */
	public Object getParent() {
		return Dispatch.get(this, getIDOfName("Parent"));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Actions
	 */
	public Actions getActions() {
		return new Actions(Dispatch.get(this, getIDOfName("Actions")).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Attachments
	 */
	public Attachments getAttachments() {
		return new Attachments(Dispatch.get(this, getIDOfName("Attachments")).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getBillingInformation() {
		return Dispatch.get(this, getIDOfName("BillingInformation")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param billingInformation an input-parameter of type String
	 */
	public void setBillingInformation(String billingInformation) {
		Dispatch.put(this, getIDOfName("BillingInformation"), billingInformation);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getBody() {
		return Dispatch.get(this, getIDOfName("Body")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param body an input-parameter of type String
	 */
	public void setBody(String body) {
		Dispatch.put(this, getIDOfName("Body"), body);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getCategories() {
		return Dispatch.get(this, getIDOfName("Categories")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param categories an input-parameter of type String
	 */
	public void setCategories(String categories) {
		Dispatch.put(this, getIDOfName("Categories"), categories);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getCompanies() {
		return Dispatch.get(this, getIDOfName("Companies")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param companies an input-parameter of type String
	 */
	public void setCompanies(String companies) {
		Dispatch.put(this, getIDOfName("Companies"), companies);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getConversationIndex() {
		return Dispatch.get(this, getIDOfName("ConversationIndex")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getConversationTopic() {
		return Dispatch.get(this, getIDOfName("ConversationTopic")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type java.util.Date
	 */
	public java.util.Date getCreationTime() {
		return Dispatch.get(this, getIDOfName("CreationTime")).getJavaDate();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getEntryID() {
		return Dispatch.get(this, getIDOfName("EntryID")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type FormDescription
	 */
	public FormDescription getFormDescription() {
		return new FormDescription(Dispatch.get(this, getIDOfName("FormDescription")).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type _Inspector
	 */
	public _Inspector getGetInspector() {
		return new _Inspector(Dispatch.get(this, getIDOfName("GetInspector")).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getImportance() {
		return Dispatch.get(this, getIDOfName("Importance")).changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param importance an input-parameter of type int
	 */
	public void setImportance(int importance) {
		Dispatch.put(this, getIDOfName("Importance"), new Variant(importance));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type java.util.Date
	 */
	public java.util.Date getLastModificationTime() {
		return Dispatch.get(this, getIDOfName("LastModificationTime")).getJavaDate();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Variant
	 */
	public Variant getMAPIOBJECT() {
		return Dispatch.get(this, getIDOfName("MAPIOBJECT"));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getMessageClass() {
		return Dispatch.get(this, getIDOfName("MessageClass")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param messageClass an input-parameter of type String
	 */
	public void setMessageClass(String messageClass) {
		Dispatch.put(this, getIDOfName("MessageClass"), messageClass);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getMileage() {
		return Dispatch.get(this, getIDOfName("Mileage")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param mileage an input-parameter of type String
	 */
	public void setMileage(String mileage) {
		Dispatch.put(this, getIDOfName("Mileage"), mileage);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getNoAging() {
		return Dispatch.get(this, getIDOfName("NoAging")).changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param noAging an input-parameter of type boolean
	 */
	public void setNoAging(boolean noAging) {
		Dispatch.put(this, getIDOfName("NoAging"), new Variant(noAging));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getOutlookInternalVersion() {
		return Dispatch.get(this, getIDOfName("OutlookInternalVersion")).changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getOutlookVersion() {
		return Dispatch.get(this, getIDOfName("OutlookVersion")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getSaved() {
		return Dispatch.get(this, getIDOfName("Saved")).changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getSensitivity() {
		return Dispatch.get(this, getIDOfName("Sensitivity")).changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param sensitivity an input-parameter of type int
	 */
	public void setSensitivity(int sensitivity) {
		Dispatch.put(this, getIDOfName("Sensitivity"), new Variant(sensitivity));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getSize() {
		return Dispatch.get(this, getIDOfName("Size")).changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getSubject() {
		return Dispatch.get(this, getIDOfName("Subject")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param subject an input-parameter of type String
	 */
	public void setSubject(String subject) {
		Dispatch.put(this, getIDOfName("Subject"), subject);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getUnRead() {
		return Dispatch.get(this, getIDOfName("UnRead")).changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param unRead an input-parameter of type boolean
	 */
	public void setUnRead(boolean unRead) {
		Dispatch.put(this, getIDOfName("UnRead"), new Variant(unRead));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type UserProperties
	 */
	public UserProperties getUserProperties() {
		return new UserProperties(Dispatch.get(this, getIDOfName("UserProperties")).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param saveMode an input-parameter of type int
	 */
	public void close(int saveMode) {
		Dispatch.call(this, getIDOfName("Close"), new Variant(saveMode));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Object
	 */
	public Object copy() {
		return Dispatch.call(this, getIDOfName("Copy"));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void delete() {
		Dispatch.call(this, getIDOfName("Delete"));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param modal an input-parameter of type Variant
	 */
	public void display(Variant modal) {
		Dispatch.call(this, getIDOfName("Display"), modal);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void display() {
		Dispatch.call(this, getIDOfName("Display"));
	}

//	/**
//	 * Wrapper for calling the ActiveX-Method and receiving the output-parameter(s).
//	 * @param modal an input-parameter of type Variant
//	 */
//	public void display(Variant modal) {
//		Dispatch.call(this, getIDOfName("Display"), modal);
//
//	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param destFldr an input-parameter of type MAPIFolder
	 * @return the result is of type Object
	 */
	public Object move(MAPIFolder destFldr) {
		return Dispatch.call(this, getIDOfName("Move"), destFldr);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void printOut() {
		Dispatch.call(this, getIDOfName("PrintOut"));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void save() {
		Dispatch.call(this, getIDOfName("Save"));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param path an input-parameter of type String
	 * @param type an input-parameter of type Variant
	 */
	public void saveAs(String path, Variant type) {
		Dispatch.call(this, getIDOfName("SaveAs"), path, type);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param path an input-parameter of type String
	 */
	public void saveAs(String path) {
		Dispatch.call(this, getIDOfName("SaveAs"), path);
	}

//	/**
//	 * Wrapper for calling the ActiveX-Method and receiving the output-parameter(s).
//	 * @param path an input-parameter of type String
//	 * @param type an input-parameter of type Variant
//	 */
//	public void saveAs(String path, Variant type) {
//		Dispatch.call(this, getIDOfName("SaveAs"), path, type);
//
//	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getAccount() {
		return Dispatch.get(this, getIDOfName("Account")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param account an input-parameter of type String
	 */
	public void setAccount(String account) {
		Dispatch.put(this, getIDOfName("Account"), account);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type java.util.Date
	 */
	public java.util.Date getAnniversary() {
		return Dispatch.get(this, getIDOfName("Anniversary")).getJavaDate();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param anniversary an input-parameter of type java.util.Date
	 */
	public void setAnniversary(java.util.Date anniversary) {
		Dispatch.put(this, getIDOfName("Anniversary"), new Variant(anniversary));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getAssistantName() {
		return Dispatch.get(this, getIDOfName("AssistantName")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param assistantName an input-parameter of type String
	 */
	public void setAssistantName(String assistantName) {
		Dispatch.put(this, getIDOfName("AssistantName"), assistantName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getAssistantTelephoneNumber() {
		return Dispatch.get(this, getIDOfName("AssistantTelephoneNumber")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param assistantTelephoneNumber an input-parameter of type String
	 */
	public void setAssistantTelephoneNumber(String assistantTelephoneNumber) {
		Dispatch.put(this, getIDOfName("AssistantTelephoneNumber"), assistantTelephoneNumber);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type java.util.Date
	 */
	public java.util.Date getBirthday() {
		return Dispatch.get(this, getIDOfName("Birthday")).getJavaDate();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param birthday an input-parameter of type java.util.Date
	 */
	public void setBirthday(java.util.Date birthday) {
		Dispatch.put(this, getIDOfName("Birthday"), new Variant(birthday));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getBusiness2TelephoneNumber() {
		return Dispatch.get(this, getIDOfName("Business2TelephoneNumber")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param business2TelephoneNumber an input-parameter of type String
	 */
	public void setBusiness2TelephoneNumber(String business2TelephoneNumber) {
		Dispatch.put(this, getIDOfName("Business2TelephoneNumber"), business2TelephoneNumber);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getBusinessAddress() {
		return Dispatch.get(this, getIDOfName("BusinessAddress")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param businessAddress an input-parameter of type String
	 */
	public void setBusinessAddress(String businessAddress) {
		Dispatch.put(this, getIDOfName("BusinessAddress"), businessAddress);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getBusinessAddressCity() {
		return Dispatch.get(this, getIDOfName("BusinessAddressCity")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param businessAddressCity an input-parameter of type String
	 */
	public void setBusinessAddressCity(String businessAddressCity) {
		Dispatch.put(this, getIDOfName("BusinessAddressCity"), businessAddressCity);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getBusinessAddressCountry() {
		return Dispatch.get(this, getIDOfName("BusinessAddressCountry")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param businessAddressCountry an input-parameter of type String
	 */
	public void setBusinessAddressCountry(String businessAddressCountry) {
		Dispatch.put(this, getIDOfName("BusinessAddressCountry"), businessAddressCountry);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getBusinessAddressPostalCode() {
		return Dispatch.get(this, getIDOfName("BusinessAddressPostalCode")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param businessAddressPostalCode an input-parameter of type String
	 */
	public void setBusinessAddressPostalCode(String businessAddressPostalCode) {
		Dispatch.put(this, getIDOfName("BusinessAddressPostalCode"), businessAddressPostalCode);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getBusinessAddressPostOfficeBox() {
		return Dispatch.get(this, getIDOfName("BusinessAddressPostOfficeBox")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param businessAddressPostOfficeBox an input-parameter of type String
	 */
	public void setBusinessAddressPostOfficeBox(String businessAddressPostOfficeBox) {
		Dispatch.put(this, getIDOfName("BusinessAddressPostOfficeBox"), businessAddressPostOfficeBox);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getBusinessAddressState() {
		return Dispatch.get(this, getIDOfName("BusinessAddressState")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param businessAddressState an input-parameter of type String
	 */
	public void setBusinessAddressState(String businessAddressState) {
		Dispatch.put(this, getIDOfName("BusinessAddressState"), businessAddressState);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getBusinessAddressStreet() {
		return Dispatch.get(this, getIDOfName("BusinessAddressStreet")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param businessAddressStreet an input-parameter of type String
	 */
	public void setBusinessAddressStreet(String businessAddressStreet) {
		Dispatch.put(this, getIDOfName("BusinessAddressStreet"), businessAddressStreet);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getBusinessFaxNumber() {
		return Dispatch.get(this, getIDOfName("BusinessFaxNumber")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param businessFaxNumber an input-parameter of type String
	 */
	public void setBusinessFaxNumber(String businessFaxNumber) {
		Dispatch.put(this, getIDOfName("BusinessFaxNumber"), businessFaxNumber);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getBusinessHomePage() {
		return Dispatch.get(this, getIDOfName("BusinessHomePage")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param businessHomePage an input-parameter of type String
	 */
	public void setBusinessHomePage(String businessHomePage) {
		Dispatch.put(this, getIDOfName("BusinessHomePage"), businessHomePage);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getBusinessTelephoneNumber() {
		return Dispatch.get(this, getIDOfName("BusinessTelephoneNumber")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param businessTelephoneNumber an input-parameter of type String
	 */
	public void setBusinessTelephoneNumber(String businessTelephoneNumber) {
		Dispatch.put(this, getIDOfName("BusinessTelephoneNumber"), businessTelephoneNumber);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getCallbackTelephoneNumber() {
		return Dispatch.get(this, getIDOfName("CallbackTelephoneNumber")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param callbackTelephoneNumber an input-parameter of type String
	 */
	public void setCallbackTelephoneNumber(String callbackTelephoneNumber) {
		Dispatch.put(this, getIDOfName("CallbackTelephoneNumber"), callbackTelephoneNumber);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getCarTelephoneNumber() {
		return Dispatch.get(this, getIDOfName("CarTelephoneNumber")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param carTelephoneNumber an input-parameter of type String
	 */
	public void setCarTelephoneNumber(String carTelephoneNumber) {
		Dispatch.put(this, getIDOfName("CarTelephoneNumber"), carTelephoneNumber);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getChildren() {
		return Dispatch.get(this, getIDOfName("Children")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param children an input-parameter of type String
	 */
	public void setChildren(String children) {
		Dispatch.put(this, getIDOfName("Children"), children);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getCompanyAndFullName() {
		return Dispatch.get(this, getIDOfName("CompanyAndFullName")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getCompanyLastFirstNoSpace() {
		return Dispatch.get(this, getIDOfName("CompanyLastFirstNoSpace")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getCompanyLastFirstSpaceOnly() {
		return Dispatch.get(this, getIDOfName("CompanyLastFirstSpaceOnly")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getCompanyMainTelephoneNumber() {
		return Dispatch.get(this, getIDOfName("CompanyMainTelephoneNumber")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param companyMainTelephoneNumber an input-parameter of type String
	 */
	public void setCompanyMainTelephoneNumber(String companyMainTelephoneNumber) {
		Dispatch.put(this, getIDOfName("CompanyMainTelephoneNumber"), companyMainTelephoneNumber);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getCompanyName() {
		return Dispatch.get(this, getIDOfName("CompanyName")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param companyName an input-parameter of type String
	 */
	public void setCompanyName(String companyName) {
		Dispatch.put(this, getIDOfName("CompanyName"), companyName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getComputerNetworkName() {
		return Dispatch.get(this, getIDOfName("ComputerNetworkName")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param computerNetworkName an input-parameter of type String
	 */
	public void setComputerNetworkName(String computerNetworkName) {
		Dispatch.put(this, getIDOfName("ComputerNetworkName"), computerNetworkName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getCustomerID() {
		return Dispatch.get(this, getIDOfName("CustomerID")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param customerID an input-parameter of type String
	 */
	public void setCustomerID(String customerID) {
		Dispatch.put(this, getIDOfName("CustomerID"), customerID);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getDepartment() {
		return Dispatch.get(this, getIDOfName("Department")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param department an input-parameter of type String
	 */
	public void setDepartment(String department) {
		Dispatch.put(this, getIDOfName("Department"), department);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getEmail1Address() {
		return Dispatch.get(this, getIDOfName("Email1Address")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param email1Address an input-parameter of type String
	 */
	public void setEmail1Address(String email1Address) {
		Dispatch.put(this, getIDOfName("Email1Address"), email1Address);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getEmail1AddressType() {
		return Dispatch.get(this, getIDOfName("Email1AddressType")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param email1AddressType an input-parameter of type String
	 */
	public void setEmail1AddressType(String email1AddressType) {
		Dispatch.put(this, getIDOfName("Email1AddressType"), email1AddressType);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getEmail1DisplayName() {
		return Dispatch.get(this, getIDOfName("Email1DisplayName")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getEmail1EntryID() {
		return Dispatch.get(this, getIDOfName("Email1EntryID")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getEmail2Address() {
		return Dispatch.get(this, getIDOfName("Email2Address")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param email2Address an input-parameter of type String
	 */
	public void setEmail2Address(String email2Address) {
		Dispatch.put(this, getIDOfName("Email2Address"), email2Address);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getEmail2AddressType() {
		return Dispatch.get(this, getIDOfName("Email2AddressType")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param email2AddressType an input-parameter of type String
	 */
	public void setEmail2AddressType(String email2AddressType) {
		Dispatch.put(this, getIDOfName("Email2AddressType"), email2AddressType);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getEmail2DisplayName() {
		return Dispatch.get(this, getIDOfName("Email2DisplayName")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getEmail2EntryID() {
		return Dispatch.get(this, getIDOfName("Email2EntryID")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getEmail3Address() {
		return Dispatch.get(this, getIDOfName("Email3Address")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param email3Address an input-parameter of type String
	 */
	public void setEmail3Address(String email3Address) {
		Dispatch.put(this, getIDOfName("Email3Address"), email3Address);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getEmail3AddressType() {
		return Dispatch.get(this, getIDOfName("Email3AddressType")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param email3AddressType an input-parameter of type String
	 */
	public void setEmail3AddressType(String email3AddressType) {
		Dispatch.put(this, getIDOfName("Email3AddressType"), email3AddressType);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getEmail3DisplayName() {
		return Dispatch.get(this, getIDOfName("Email3DisplayName")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getEmail3EntryID() {
		return Dispatch.get(this, getIDOfName("Email3EntryID")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getFileAs() {
		return Dispatch.get(this, getIDOfName("FileAs")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param fileAs an input-parameter of type String
	 */
	public void setFileAs(String fileAs) {
		Dispatch.put(this, getIDOfName("FileAs"), fileAs);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getFirstName() {
		return Dispatch.get(this, getIDOfName("FirstName")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param firstName an input-parameter of type String
	 */
	public void setFirstName(String firstName) {
		Dispatch.put(this, getIDOfName("FirstName"), firstName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getFTPSite() {
		return Dispatch.get(this, getIDOfName("FTPSite")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param fTPSite an input-parameter of type String
	 */
	public void setFTPSite(String fTPSite) {
		Dispatch.put(this, getIDOfName("FTPSite"), fTPSite);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getFullName() {
		return Dispatch.get(this, getIDOfName("FullName")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param fullName an input-parameter of type String
	 */
	public void setFullName(String fullName) {
		Dispatch.put(this, getIDOfName("FullName"), fullName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getFullNameAndCompany() {
		return Dispatch.get(this, getIDOfName("FullNameAndCompany")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getGender() {
		return Dispatch.get(this, getIDOfName("Gender")).changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param gender an input-parameter of type int
	 */
	public void setGender(int gender) {
		Dispatch.put(this, getIDOfName("Gender"), new Variant(gender));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getGovernmentIDNumber() {
		return Dispatch.get(this, getIDOfName("GovernmentIDNumber")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param governmentIDNumber an input-parameter of type String
	 */
	public void setGovernmentIDNumber(String governmentIDNumber) {
		Dispatch.put(this, getIDOfName("GovernmentIDNumber"), governmentIDNumber);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getHobby() {
		return Dispatch.get(this, getIDOfName("Hobby")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param hobby an input-parameter of type String
	 */
	public void setHobby(String hobby) {
		Dispatch.put(this, getIDOfName("Hobby"), hobby);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getHome2TelephoneNumber() {
		return Dispatch.get(this, getIDOfName("Home2TelephoneNumber")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param home2TelephoneNumber an input-parameter of type String
	 */
	public void setHome2TelephoneNumber(String home2TelephoneNumber) {
		Dispatch.put(this, getIDOfName("Home2TelephoneNumber"), home2TelephoneNumber);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getHomeAddress() {
		return Dispatch.get(this, getIDOfName("HomeAddress")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param homeAddress an input-parameter of type String
	 */
	public void setHomeAddress(String homeAddress) {
		Dispatch.put(this, getIDOfName("HomeAddress"), homeAddress);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getHomeAddressCity() {
		return Dispatch.get(this, getIDOfName("HomeAddressCity")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param homeAddressCity an input-parameter of type String
	 */
	public void setHomeAddressCity(String homeAddressCity) {
		Dispatch.put(this, getIDOfName("HomeAddressCity"), homeAddressCity);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getHomeAddressCountry() {
		return Dispatch.get(this, getIDOfName("HomeAddressCountry")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param homeAddressCountry an input-parameter of type String
	 */
	public void setHomeAddressCountry(String homeAddressCountry) {
		Dispatch.put(this, getIDOfName("HomeAddressCountry"), homeAddressCountry);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getHomeAddressPostalCode() {
		return Dispatch.get(this, getIDOfName("HomeAddressPostalCode")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param homeAddressPostalCode an input-parameter of type String
	 */
	public void setHomeAddressPostalCode(String homeAddressPostalCode) {
		Dispatch.put(this, getIDOfName("HomeAddressPostalCode"), homeAddressPostalCode);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getHomeAddressPostOfficeBox() {
		return Dispatch.get(this, getIDOfName("HomeAddressPostOfficeBox")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param homeAddressPostOfficeBox an input-parameter of type String
	 */
	public void setHomeAddressPostOfficeBox(String homeAddressPostOfficeBox) {
		Dispatch.put(this, getIDOfName("HomeAddressPostOfficeBox"), homeAddressPostOfficeBox);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getHomeAddressState() {
		return Dispatch.get(this, getIDOfName("HomeAddressState")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param homeAddressState an input-parameter of type String
	 */
	public void setHomeAddressState(String homeAddressState) {
		Dispatch.put(this, getIDOfName("HomeAddressState"), homeAddressState);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getHomeAddressStreet() {
		return Dispatch.get(this, getIDOfName("HomeAddressStreet")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param homeAddressStreet an input-parameter of type String
	 */
	public void setHomeAddressStreet(String homeAddressStreet) {
		Dispatch.put(this, getIDOfName("HomeAddressStreet"), homeAddressStreet);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getHomeFaxNumber() {
		return Dispatch.get(this, getIDOfName("HomeFaxNumber")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param homeFaxNumber an input-parameter of type String
	 */
	public void setHomeFaxNumber(String homeFaxNumber) {
		Dispatch.put(this, getIDOfName("HomeFaxNumber"), homeFaxNumber);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getHomeTelephoneNumber() {
		return Dispatch.get(this, getIDOfName("HomeTelephoneNumber")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param homeTelephoneNumber an input-parameter of type String
	 */
	public void setHomeTelephoneNumber(String homeTelephoneNumber) {
		Dispatch.put(this, getIDOfName("HomeTelephoneNumber"), homeTelephoneNumber);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getInitials() {
		return Dispatch.get(this, getIDOfName("Initials")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param initials an input-parameter of type String
	 */
	public void setInitials(String initials) {
		Dispatch.put(this, getIDOfName("Initials"), initials);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getInternetFreeBusyAddress() {
		return Dispatch.get(this, getIDOfName("InternetFreeBusyAddress")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param internetFreeBusyAddress an input-parameter of type String
	 */
	public void setInternetFreeBusyAddress(String internetFreeBusyAddress) {
		Dispatch.put(this, getIDOfName("InternetFreeBusyAddress"), internetFreeBusyAddress);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getISDNNumber() {
		return Dispatch.get(this, getIDOfName("ISDNNumber")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param iSDNNumber an input-parameter of type String
	 */
	public void setISDNNumber(String iSDNNumber) {
		Dispatch.put(this, getIDOfName("ISDNNumber"), iSDNNumber);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getJobTitle() {
		return Dispatch.get(this, getIDOfName("JobTitle")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param jobTitle an input-parameter of type String
	 */
	public void setJobTitle(String jobTitle) {
		Dispatch.put(this, getIDOfName("JobTitle"), jobTitle);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getJournal() {
		return Dispatch.get(this, getIDOfName("Journal")).changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param journal an input-parameter of type boolean
	 */
	public void setJournal(boolean journal) {
		Dispatch.put(this, getIDOfName("Journal"), new Variant(journal));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getLanguage() {
		return Dispatch.get(this, getIDOfName("Language")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param language an input-parameter of type String
	 */
	public void setLanguage(String language) {
		Dispatch.put(this, getIDOfName("Language"), language);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getLastFirstAndSuffix() {
		return Dispatch.get(this, getIDOfName("LastFirstAndSuffix")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getLastFirstNoSpace() {
		return Dispatch.get(this, getIDOfName("LastFirstNoSpace")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getLastFirstNoSpaceCompany() {
		return Dispatch.get(this, getIDOfName("LastFirstNoSpaceCompany")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getLastFirstSpaceOnly() {
		return Dispatch.get(this, getIDOfName("LastFirstSpaceOnly")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getLastFirstSpaceOnlyCompany() {
		return Dispatch.get(this, getIDOfName("LastFirstSpaceOnlyCompany")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getLastName() {
		return Dispatch.get(this, getIDOfName("LastName")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param lastName an input-parameter of type String
	 */
	public void setLastName(String lastName) {
		Dispatch.put(this, getIDOfName("LastName"), lastName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getLastNameAndFirstName() {
		return Dispatch.get(this, getIDOfName("LastNameAndFirstName")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getMailingAddress() {
		return Dispatch.get(this, getIDOfName("MailingAddress")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param mailingAddress an input-parameter of type String
	 */
	public void setMailingAddress(String mailingAddress) {
		Dispatch.put(this, getIDOfName("MailingAddress"), mailingAddress);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getMailingAddressCity() {
		return Dispatch.get(this, getIDOfName("MailingAddressCity")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param mailingAddressCity an input-parameter of type String
	 */
	public void setMailingAddressCity(String mailingAddressCity) {
		Dispatch.put(this, getIDOfName("MailingAddressCity"), mailingAddressCity);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getMailingAddressCountry() {
		return Dispatch.get(this, getIDOfName("MailingAddressCountry")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param mailingAddressCountry an input-parameter of type String
	 */
	public void setMailingAddressCountry(String mailingAddressCountry) {
		Dispatch.put(this, getIDOfName("MailingAddressCountry"), mailingAddressCountry);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getMailingAddressPostalCode() {
		return Dispatch.get(this, getIDOfName("MailingAddressPostalCode")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param mailingAddressPostalCode an input-parameter of type String
	 */
	public void setMailingAddressPostalCode(String mailingAddressPostalCode) {
		Dispatch.put(this, getIDOfName("MailingAddressPostalCode"), mailingAddressPostalCode);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getMailingAddressPostOfficeBox() {
		return Dispatch.get(this, getIDOfName("MailingAddressPostOfficeBox")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param mailingAddressPostOfficeBox an input-parameter of type String
	 */
	public void setMailingAddressPostOfficeBox(String mailingAddressPostOfficeBox) {
		Dispatch.put(this, getIDOfName("MailingAddressPostOfficeBox"), mailingAddressPostOfficeBox);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getMailingAddressState() {
		return Dispatch.get(this, getIDOfName("MailingAddressState")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param mailingAddressState an input-parameter of type String
	 */
	public void setMailingAddressState(String mailingAddressState) {
		Dispatch.put(this, getIDOfName("MailingAddressState"), mailingAddressState);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getMailingAddressStreet() {
		return Dispatch.get(this, getIDOfName("MailingAddressStreet")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param mailingAddressStreet an input-parameter of type String
	 */
	public void setMailingAddressStreet(String mailingAddressStreet) {
		Dispatch.put(this, getIDOfName("MailingAddressStreet"), mailingAddressStreet);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getManagerName() {
		return Dispatch.get(this, getIDOfName("ManagerName")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param managerName an input-parameter of type String
	 */
	public void setManagerName(String managerName) {
		Dispatch.put(this, getIDOfName("ManagerName"), managerName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getMiddleName() {
		return Dispatch.get(this, getIDOfName("MiddleName")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param middleName an input-parameter of type String
	 */
	public void setMiddleName(String middleName) {
		Dispatch.put(this, getIDOfName("MiddleName"), middleName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getMobileTelephoneNumber() {
		return Dispatch.get(this, getIDOfName("MobileTelephoneNumber")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param mobileTelephoneNumber an input-parameter of type String
	 */
	public void setMobileTelephoneNumber(String mobileTelephoneNumber) {
		Dispatch.put(this, getIDOfName("MobileTelephoneNumber"), mobileTelephoneNumber);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getNetMeetingAlias() {
		return Dispatch.get(this, getIDOfName("NetMeetingAlias")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param netMeetingAlias an input-parameter of type String
	 */
	public void setNetMeetingAlias(String netMeetingAlias) {
		Dispatch.put(this, getIDOfName("NetMeetingAlias"), netMeetingAlias);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getNetMeetingServer() {
		return Dispatch.get(this, getIDOfName("NetMeetingServer")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param netMeetingServer an input-parameter of type String
	 */
	public void setNetMeetingServer(String netMeetingServer) {
		Dispatch.put(this, getIDOfName("NetMeetingServer"), netMeetingServer);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getNickName() {
		return Dispatch.get(this, getIDOfName("NickName")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param nickName an input-parameter of type String
	 */
	public void setNickName(String nickName) {
		Dispatch.put(this, getIDOfName("NickName"), nickName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getOfficeLocation() {
		return Dispatch.get(this, getIDOfName("OfficeLocation")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param officeLocation an input-parameter of type String
	 */
	public void setOfficeLocation(String officeLocation) {
		Dispatch.put(this, getIDOfName("OfficeLocation"), officeLocation);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getOrganizationalIDNumber() {
		return Dispatch.get(this, getIDOfName("OrganizationalIDNumber")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param organizationalIDNumber an input-parameter of type String
	 */
	public void setOrganizationalIDNumber(String organizationalIDNumber) {
		Dispatch.put(this, getIDOfName("OrganizationalIDNumber"), organizationalIDNumber);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getOtherAddress() {
		return Dispatch.get(this, getIDOfName("OtherAddress")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param otherAddress an input-parameter of type String
	 */
	public void setOtherAddress(String otherAddress) {
		Dispatch.put(this, getIDOfName("OtherAddress"), otherAddress);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getOtherAddressCity() {
		return Dispatch.get(this, getIDOfName("OtherAddressCity")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param otherAddressCity an input-parameter of type String
	 */
	public void setOtherAddressCity(String otherAddressCity) {
		Dispatch.put(this, getIDOfName("OtherAddressCity"), otherAddressCity);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getOtherAddressCountry() {
		return Dispatch.get(this, getIDOfName("OtherAddressCountry")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param otherAddressCountry an input-parameter of type String
	 */
	public void setOtherAddressCountry(String otherAddressCountry) {
		Dispatch.put(this, getIDOfName("OtherAddressCountry"), otherAddressCountry);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getOtherAddressPostalCode() {
		return Dispatch.get(this, getIDOfName("OtherAddressPostalCode")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param otherAddressPostalCode an input-parameter of type String
	 */
	public void setOtherAddressPostalCode(String otherAddressPostalCode) {
		Dispatch.put(this, getIDOfName("OtherAddressPostalCode"), otherAddressPostalCode);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getOtherAddressPostOfficeBox() {
		return Dispatch.get(this, getIDOfName("OtherAddressPostOfficeBox")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param otherAddressPostOfficeBox an input-parameter of type String
	 */
	public void setOtherAddressPostOfficeBox(String otherAddressPostOfficeBox) {
		Dispatch.put(this, getIDOfName("OtherAddressPostOfficeBox"), otherAddressPostOfficeBox);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getOtherAddressState() {
		return Dispatch.get(this, getIDOfName("OtherAddressState")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param otherAddressState an input-parameter of type String
	 */
	public void setOtherAddressState(String otherAddressState) {
		Dispatch.put(this, getIDOfName("OtherAddressState"), otherAddressState);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getOtherAddressStreet() {
		return Dispatch.get(this, getIDOfName("OtherAddressStreet")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param otherAddressStreet an input-parameter of type String
	 */
	public void setOtherAddressStreet(String otherAddressStreet) {
		Dispatch.put(this, getIDOfName("OtherAddressStreet"), otherAddressStreet);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getOtherFaxNumber() {
		return Dispatch.get(this, getIDOfName("OtherFaxNumber")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param otherFaxNumber an input-parameter of type String
	 */
	public void setOtherFaxNumber(String otherFaxNumber) {
		Dispatch.put(this, getIDOfName("OtherFaxNumber"), otherFaxNumber);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getOtherTelephoneNumber() {
		return Dispatch.get(this, getIDOfName("OtherTelephoneNumber")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param otherTelephoneNumber an input-parameter of type String
	 */
	public void setOtherTelephoneNumber(String otherTelephoneNumber) {
		Dispatch.put(this, getIDOfName("OtherTelephoneNumber"), otherTelephoneNumber);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getPagerNumber() {
		return Dispatch.get(this, getIDOfName("PagerNumber")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param pagerNumber an input-parameter of type String
	 */
	public void setPagerNumber(String pagerNumber) {
		Dispatch.put(this, getIDOfName("PagerNumber"), pagerNumber);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getPersonalHomePage() {
		return Dispatch.get(this, getIDOfName("PersonalHomePage")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param personalHomePage an input-parameter of type String
	 */
	public void setPersonalHomePage(String personalHomePage) {
		Dispatch.put(this, getIDOfName("PersonalHomePage"), personalHomePage);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getPrimaryTelephoneNumber() {
		return Dispatch.get(this, getIDOfName("PrimaryTelephoneNumber")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param primaryTelephoneNumber an input-parameter of type String
	 */
	public void setPrimaryTelephoneNumber(String primaryTelephoneNumber) {
		Dispatch.put(this, getIDOfName("PrimaryTelephoneNumber"), primaryTelephoneNumber);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getProfession() {
		return Dispatch.get(this, getIDOfName("Profession")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param profession an input-parameter of type String
	 */
	public void setProfession(String profession) {
		Dispatch.put(this, getIDOfName("Profession"), profession);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getRadioTelephoneNumber() {
		return Dispatch.get(this, getIDOfName("RadioTelephoneNumber")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param radioTelephoneNumber an input-parameter of type String
	 */
	public void setRadioTelephoneNumber(String radioTelephoneNumber) {
		Dispatch.put(this, getIDOfName("RadioTelephoneNumber"), radioTelephoneNumber);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getReferredBy() {
		return Dispatch.get(this, getIDOfName("ReferredBy")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param referredBy an input-parameter of type String
	 */
	public void setReferredBy(String referredBy) {
		Dispatch.put(this, getIDOfName("ReferredBy"), referredBy);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getSelectedMailingAddress() {
		return Dispatch.get(this, getIDOfName("SelectedMailingAddress")).changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param selectedMailingAddress an input-parameter of type int
	 */
	public void setSelectedMailingAddress(int selectedMailingAddress) {
		Dispatch.put(this, getIDOfName("SelectedMailingAddress"), new Variant(selectedMailingAddress));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getSpouse() {
		return Dispatch.get(this, getIDOfName("Spouse")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param spouse an input-parameter of type String
	 */
	public void setSpouse(String spouse) {
		Dispatch.put(this, getIDOfName("Spouse"), spouse);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getSuffix() {
		return Dispatch.get(this, getIDOfName("Suffix")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param suffix an input-parameter of type String
	 */
	public void setSuffix(String suffix) {
		Dispatch.put(this, getIDOfName("Suffix"), suffix);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getTelexNumber() {
		return Dispatch.get(this, getIDOfName("TelexNumber")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param telexNumber an input-parameter of type String
	 */
	public void setTelexNumber(String telexNumber) {
		Dispatch.put(this, getIDOfName("TelexNumber"), telexNumber);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getTitle() {
		return Dispatch.get(this, getIDOfName("Title")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param title an input-parameter of type String
	 */
	public void setTitle(String title) {
		Dispatch.put(this, getIDOfName("Title"), title);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getTTYTDDTelephoneNumber() {
		return Dispatch.get(this, getIDOfName("TTYTDDTelephoneNumber")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param tTYTDDTelephoneNumber an input-parameter of type String
	 */
	public void setTTYTDDTelephoneNumber(String tTYTDDTelephoneNumber) {
		Dispatch.put(this, getIDOfName("TTYTDDTelephoneNumber"), tTYTDDTelephoneNumber);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getUser1() {
		return Dispatch.get(this, getIDOfName("User1")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param user1 an input-parameter of type String
	 */
	public void setUser1(String user1) {
		Dispatch.put(this, getIDOfName("User1"), user1);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getUser2() {
		return Dispatch.get(this, getIDOfName("User2")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param user2 an input-parameter of type String
	 */
	public void setUser2(String user2) {
		Dispatch.put(this, getIDOfName("User2"), user2);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getUser3() {
		return Dispatch.get(this, getIDOfName("User3")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param user3 an input-parameter of type String
	 */
	public void setUser3(String user3) {
		Dispatch.put(this, getIDOfName("User3"), user3);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getUser4() {
		return Dispatch.get(this, getIDOfName("User4")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param user4 an input-parameter of type String
	 */
	public void setUser4(String user4) {
		Dispatch.put(this, getIDOfName("User4"), user4);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getUserCertificate() {
		return Dispatch.get(this, getIDOfName("UserCertificate")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param userCertificate an input-parameter of type String
	 */
	public void setUserCertificate(String userCertificate) {
		Dispatch.put(this, getIDOfName("UserCertificate"), userCertificate);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getWebPage() {
		return Dispatch.get(this, getIDOfName("WebPage")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param webPage an input-parameter of type String
	 */
	public void setWebPage(String webPage) {
		Dispatch.put(this, getIDOfName("WebPage"), webPage);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getYomiCompanyName() {
		return Dispatch.get(this, getIDOfName("YomiCompanyName")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param yomiCompanyName an input-parameter of type String
	 */
	public void setYomiCompanyName(String yomiCompanyName) {
		Dispatch.put(this, getIDOfName("YomiCompanyName"), yomiCompanyName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getYomiFirstName() {
		return Dispatch.get(this, getIDOfName("YomiFirstName")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param yomiFirstName an input-parameter of type String
	 */
	public void setYomiFirstName(String yomiFirstName) {
		Dispatch.put(this, getIDOfName("YomiFirstName"), yomiFirstName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getYomiLastName() {
		return Dispatch.get(this, getIDOfName("YomiLastName")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param yomiLastName an input-parameter of type String
	 */
	public void setYomiLastName(String yomiLastName) {
		Dispatch.put(this, getIDOfName("YomiLastName"), yomiLastName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type MailItem
	 */
	public MailItem forwardAsVcard() {
		return new MailItem(Dispatch.call(this, getIDOfName("ForwardAsVcard")).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Links
	 */
	public Links getLinks() {
		return new Links(Dispatch.get(this, getIDOfName("Links")).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type ItemProperties
	 */
	public ItemProperties getItemProperties() {
		return new ItemProperties(Dispatch.get(this, getIDOfName("ItemProperties")).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getLastFirstNoSpaceAndSuffix() {
		return Dispatch.get(this, getIDOfName("LastFirstNoSpaceAndSuffix")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getDownloadState() {
		return Dispatch.get(this, getIDOfName("DownloadState")).changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void showCategoriesDialog() {
		Dispatch.call(this, getIDOfName("ShowCategoriesDialog"));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getIMAddress() {
		return Dispatch.get(this, getIDOfName("IMAddress")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param iMAddress an input-parameter of type String
	 */
	public void setIMAddress(String iMAddress) {
		Dispatch.put(this, getIDOfName("IMAddress"), iMAddress);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getMarkForDownload() {
		return Dispatch.get(this, getIDOfName("MarkForDownload")).changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param markForDownload an input-parameter of type int
	 */
	public void setMarkForDownload(int markForDownload) {
		Dispatch.put(this, getIDOfName("MarkForDownload"), new Variant(markForDownload));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param email1DisplayName an input-parameter of type String
	 */
	public void setEmail1DisplayName(String email1DisplayName) {
		Dispatch.put(this, getIDOfName("Email1DisplayName"), email1DisplayName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param email2DisplayName an input-parameter of type String
	 */
	public void setEmail2DisplayName(String email2DisplayName) {
		Dispatch.put(this, getIDOfName("Email2DisplayName"), email2DisplayName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param email3DisplayName an input-parameter of type String
	 */
	public void setEmail3DisplayName(String email3DisplayName) {
		Dispatch.put(this, getIDOfName("Email3DisplayName"), email3DisplayName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getIsConflict() {
		return Dispatch.get(this, getIDOfName("IsConflict")).changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getAutoResolvedWinner() {
		return Dispatch.get(this, getIDOfName("AutoResolvedWinner")).changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Conflicts
	 */
	public Conflicts getConflicts() {
		return new Conflicts(Dispatch.get(this, getIDOfName("Conflicts")).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param path an input-parameter of type String
	 */
	public void addPicture(String path) {
		Dispatch.call(this, getIDOfName("AddPicture"), path);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void removePicture() {
		Dispatch.call(this, getIDOfName("RemovePicture"));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getHasPicture() {
		return Dispatch.get(this, getIDOfName("HasPicture")).changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type PropertyAccessor
	 */
	public PropertyAccessor getPropertyAccessor() {
		return new PropertyAccessor(Dispatch.get(this, getIDOfName("PropertyAccessor")).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type MailItem
	 */
	public MailItem forwardAsBusinessCard() {
		return new MailItem(Dispatch.call(this, getIDOfName("ForwardAsBusinessCard")).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void showBusinessCardEditor() {
		Dispatch.call(this, getIDOfName("ShowBusinessCardEditor"));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param path an input-parameter of type String
	 */
	public void saveBusinessCardImage(String path) {
		Dispatch.call(this, getIDOfName("SaveBusinessCardImage"), path);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param phoneNumber an input-parameter of type int
	 */
	public void showCheckPhoneDialog(int phoneNumber) {
		Dispatch.call(this, getIDOfName("ShowCheckPhoneDialog"), new Variant(phoneNumber));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getTaskSubject() {
		return Dispatch.get(this, getIDOfName("TaskSubject")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param taskSubject an input-parameter of type String
	 */
	public void setTaskSubject(String taskSubject) {
		Dispatch.put(this, getIDOfName("TaskSubject"), taskSubject);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type java.util.Date
	 */
	public java.util.Date getTaskDueDate() {
		return Dispatch.get(this, getIDOfName("TaskDueDate")).getJavaDate();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param taskDueDate an input-parameter of type java.util.Date
	 */
	public void setTaskDueDate(java.util.Date taskDueDate) {
		Dispatch.put(this, getIDOfName("TaskDueDate"), new Variant(taskDueDate));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type java.util.Date
	 */
	public java.util.Date getTaskStartDate() {
		return Dispatch.get(this, getIDOfName("TaskStartDate")).getJavaDate();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param taskStartDate an input-parameter of type java.util.Date
	 */
	public void setTaskStartDate(java.util.Date taskStartDate) {
		Dispatch.put(this, getIDOfName("TaskStartDate"), new Variant(taskStartDate));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type java.util.Date
	 */
	public java.util.Date getTaskCompletedDate() {
		return Dispatch.get(this, getIDOfName("TaskCompletedDate")).getJavaDate();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param taskCompletedDate an input-parameter of type java.util.Date
	 */
	public void setTaskCompletedDate(java.util.Date taskCompletedDate) {
		Dispatch.put(this, getIDOfName("TaskCompletedDate"), new Variant(taskCompletedDate));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type java.util.Date
	 */
	public java.util.Date getToDoTaskOrdinal() {
		return Dispatch.get(this, getIDOfName("ToDoTaskOrdinal")).getJavaDate();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param toDoTaskOrdinal an input-parameter of type java.util.Date
	 */
	public void setToDoTaskOrdinal(java.util.Date toDoTaskOrdinal) {
		Dispatch.put(this, getIDOfName("ToDoTaskOrdinal"), new Variant(toDoTaskOrdinal));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getReminderOverrideDefault() {
		return Dispatch.get(this, getIDOfName("ReminderOverrideDefault")).changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param reminderOverrideDefault an input-parameter of type boolean
	 */
	public void setReminderOverrideDefault(boolean reminderOverrideDefault) {
		Dispatch.put(this, getIDOfName("ReminderOverrideDefault"), new Variant(reminderOverrideDefault));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getReminderPlaySound() {
		return Dispatch.get(this, getIDOfName("ReminderPlaySound")).changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param reminderPlaySound an input-parameter of type boolean
	 */
	public void setReminderPlaySound(boolean reminderPlaySound) {
		Dispatch.put(this, getIDOfName("ReminderPlaySound"), new Variant(reminderPlaySound));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getReminderSet() {
		return Dispatch.get(this, getIDOfName("ReminderSet")).changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param reminderSet an input-parameter of type boolean
	 */
	public void setReminderSet(boolean reminderSet) {
		Dispatch.put(this, getIDOfName("ReminderSet"), new Variant(reminderSet));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getReminderSoundFile() {
		return Dispatch.get(this, getIDOfName("ReminderSoundFile")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param reminderSoundFile an input-parameter of type String
	 */
	public void setReminderSoundFile(String reminderSoundFile) {
		Dispatch.put(this, getIDOfName("ReminderSoundFile"), reminderSoundFile);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type java.util.Date
	 */
	public java.util.Date getReminderTime() {
		return Dispatch.get(this, getIDOfName("ReminderTime")).getJavaDate();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param reminderTime an input-parameter of type java.util.Date
	 */
	public void setReminderTime(java.util.Date reminderTime) {
		Dispatch.put(this, getIDOfName("ReminderTime"), new Variant(reminderTime));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param markInterval an input-parameter of type int
	 */
	public void markAsTask(int markInterval) {
		Dispatch.call(this, getIDOfName("MarkAsTask"), new Variant(markInterval));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void clearTaskFlag() {
		Dispatch.call(this, getIDOfName("ClearTaskFlag"));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getIsMarkedAsTask() {
		return Dispatch.get(this, getIDOfName("IsMarkedAsTask")).changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getBusinessCardLayoutXml() {
		return Dispatch.get(this, getIDOfName("BusinessCardLayoutXml")).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param businessCardLayoutXml an input-parameter of type String
	 */
	public void setBusinessCardLayoutXml(String businessCardLayoutXml) {
		Dispatch.put(this, getIDOfName("BusinessCardLayoutXml"), businessCardLayoutXml);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void resetBusinessCard() {
		Dispatch.call(this, getIDOfName("ResetBusinessCard"));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param path an input-parameter of type String
	 */
	public void addBusinessCardLogoPicture(String path) {
		Dispatch.call(this, getIDOfName("AddBusinessCardLogoPicture"), path);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getBusinessCardType() {
		return Dispatch.get(this, getIDOfName("BusinessCardType")).changeType(Variant.VariantInt).getInt();
	}

}
