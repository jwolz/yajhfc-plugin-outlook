/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class _ExchangeDistributionList extends Dispatch {

	public static final String componentName = "Outlook._ExchangeDistributionList";

	public _ExchangeDistributionList() {
		super(componentName);
	}

	/**
	* This constructor is used instead of a case operation to
	* turn a Dispatch object into a wider object - it must exist
	* in every wrapper class whose instances may be returned from
	* method calls wrapped in VT_DISPATCH Variants.
	*/
	public _ExchangeDistributionList(Dispatch d) {
		// take over the IDispatch pointer
		m_pDispatch = d.m_pDispatch;
		// null out the input's pointer
		d.m_pDispatch = 0;
	}

	public _ExchangeDistributionList(String compName) {
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
	public AddressEntries getMemberOfList() {
		return new AddressEntries(Dispatch.call(this, "GetMemberOfList").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type AddressEntries
	 */
	public AddressEntries getExchangeDistributionListMembers() {
		return new AddressEntries(Dispatch.call(this, "GetExchangeDistributionListMembers").toDispatch());
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
	public String getPrimarySmtpAddress() {
		return Dispatch.get(this, "PrimarySmtpAddress").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type AddressEntries
	 */
	public AddressEntries getOwners() {
		return new AddressEntries(Dispatch.call(this, "GetOwners").toDispatch());
	}

}
