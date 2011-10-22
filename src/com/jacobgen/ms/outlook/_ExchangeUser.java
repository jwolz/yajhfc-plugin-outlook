/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class _ExchangeUser extends Dispatch {

	public static final String componentName = "Outlook._ExchangeUser";

	public _ExchangeUser() {
		super(componentName);
	}

	/**
	* This constructor is used instead of a case operation to
	* turn a Dispatch object into a wider object - it must exist
	* in every wrapper class whose instances may be returned from
	* method calls wrapped in VT_DISPATCH Variants.
	*/
	public _ExchangeUser(Dispatch d) {
		// take over the IDispatch pointer
		m_pDispatch = d.m_pDispatch;
		// null out the input's pointer
		d.m_pDispatch = 0;
	}

	public _ExchangeUser(String compName) {
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
	 * @return the result is of type String
	 */
	public String getAddress() {
		return Dispatch.get(this, "Address").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param address an input-parameter of type String
	 */
	public void setAddress(String address) {
		Dispatch.put(this, "Address", address);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getDisplayType() {
		return Dispatch.get(this, "DisplayType").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getID() {
		return Dispatch.get(this, "ID").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type AddressEntry
	 */
	public AddressEntry getManager() {
		return new AddressEntry(Dispatch.get(this, "Manager").toDispatch());
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
	 * @param mAPIOBJECT an input-parameter of type Variant
	 */
	public void setMAPIOBJECT(Variant mAPIOBJECT) {
		Dispatch.put(this, "MAPIOBJECT", mAPIOBJECT);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type AddressEntries
	 */
	public AddressEntries getMembers() {
		return new AddressEntries(Dispatch.get(this, "Members").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getName() {
		return Dispatch.get(this, "Name").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param name an input-parameter of type String
	 */
	public void setName(String name) {
		Dispatch.put(this, "Name", name);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getType() {
		return Dispatch.get(this, "Type").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param type an input-parameter of type String
	 */
	public void setType(String type) {
		Dispatch.put(this, "Type", type);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void delete() {
		Dispatch.call(this, "Delete");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param hWnd an input-parameter of type Variant
	 */
	public void details(Variant hWnd) {
		Dispatch.call(this, "Details", hWnd);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void details() {
		Dispatch.call(this, "Details");
	}


	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param start an input-parameter of type java.util.Date
	 * @param minPerChar an input-parameter of type int
	 * @param completeFormat an input-parameter of type Variant
	 * @return the result is of type String
	 */
	public String getFreeBusy(java.util.Date start, int minPerChar, Variant completeFormat) {
		return Dispatch.call(this, "GetFreeBusy", new Variant(start), new Variant(minPerChar), completeFormat).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param start an input-parameter of type java.util.Date
	 * @param minPerChar an input-parameter of type int
	 * @return the result is of type String
	 */
	public String getFreeBusy(java.util.Date start, int minPerChar) {
		return Dispatch.call(this, "GetFreeBusy", new Variant(start), new Variant(minPerChar)).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param makePermanent an input-parameter of type Variant
	 * @param refresh an input-parameter of type Variant
	 */
	public void update(Variant makePermanent, Variant refresh) {
		Dispatch.call(this, "Update", makePermanent, refresh);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param makePermanent an input-parameter of type Variant
	 */
	public void update(Variant makePermanent) {
		Dispatch.call(this, "Update", makePermanent);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void update() {
		Dispatch.call(this, "Update");
	}


	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void updateFreeBusy() {
		Dispatch.call(this, "UpdateFreeBusy");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type _ContactItem
	 */
	public _ContactItem getContact() {
		return new _ContactItem(Dispatch.call(this, "GetContact").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type ExchangeUser
	 */
	public ExchangeUser getExchangeUser() {
		return new ExchangeUser(Dispatch.call(this, "GetExchangeUser").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getAddressEntryUserType() {
		return Dispatch.get(this, "AddressEntryUserType").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type ExchangeDistributionList
	 */
	public ExchangeDistributionList getExchangeDistributionList() {
		return new ExchangeDistributionList(Dispatch.call(this, "GetExchangeDistributionList").toDispatch());
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
	 * @return the result is of type AddressEntries
	 */
	public AddressEntries getDirectReports() {
		return new AddressEntries(Dispatch.call(this, "GetDirectReports").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type AddressEntries
	 */
	public AddressEntries getMemberOfList() {
		return new AddressEntries(Dispatch.call(this, "GetMemberOfList").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getAlias() {
		return Dispatch.get(this, "Alias").toString();
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
	public String getCity() {
		return Dispatch.get(this, "City").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param city an input-parameter of type String
	 */
	public void setCity(String city) {
		Dispatch.put(this, "City", city);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getComments() {
		return Dispatch.get(this, "Comments").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param comments an input-parameter of type String
	 */
	public void setComments(String comments) {
		Dispatch.put(this, "Comments", comments);
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
	public String getPostalCode() {
		return Dispatch.get(this, "PostalCode").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param postalCode an input-parameter of type String
	 */
	public void setPostalCode(String postalCode) {
		Dispatch.put(this, "PostalCode", postalCode);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getPrimarySmtpAddress() {
		return Dispatch.get(this, "PrimarySmtpAddress").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getStateOrProvince() {
		return Dispatch.get(this, "StateOrProvince").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param stateOrProvince an input-parameter of type String
	 */
	public void setStateOrProvince(String stateOrProvince) {
		Dispatch.put(this, "StateOrProvince", stateOrProvince);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getStreetAddress() {
		return Dispatch.get(this, "StreetAddress").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param streetAddress an input-parameter of type String
	 */
	public void setStreetAddress(String streetAddress) {
		Dispatch.put(this, "StreetAddress", streetAddress);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type ExchangeUser
	 */
	public ExchangeUser getExchangeUserManager() {
		return new ExchangeUser(Dispatch.call(this, "GetExchangeUserManager").toDispatch());
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
	 * @return the result is of type String
	 */
	public String getYomiDisplayName() {
		return Dispatch.get(this, "YomiDisplayName").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param yomiDisplayName an input-parameter of type String
	 */
	public void setYomiDisplayName(String yomiDisplayName) {
		Dispatch.put(this, "YomiDisplayName", yomiDisplayName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getYomiDepartment() {
		return Dispatch.get(this, "YomiDepartment").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param yomiDepartment an input-parameter of type String
	 */
	public void setYomiDepartment(String yomiDepartment) {
		Dispatch.put(this, "YomiDepartment", yomiDepartment);
	}

}
