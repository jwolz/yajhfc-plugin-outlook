/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class Recipient extends Dispatch {

	public static final String componentName = "Outlook.Recipient";

	public Recipient() {
		super(componentName);
	}

	/**
	* This constructor is used instead of a case operation to
	* turn a Dispatch object into a wider object - it must exist
	* in every wrapper class whose instances may be returned from
	* method calls wrapped in VT_DISPATCH Variants.
	*/
	public Recipient(Dispatch d) {
		// take over the IDispatch pointer
		m_pDispatch = d.m_pDispatch;
		// null out the input's pointer
		d.m_pDispatch = 0;
	}

	public Recipient(String compName) {
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
	 * @return the result is of type AddressEntry
	 */
	public AddressEntry getAddressEntry() {
		return new AddressEntry(Dispatch.get(this, "AddressEntry").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param addressEntry an input-parameter of type AddressEntry
	 */
	public void setAddressEntry(AddressEntry addressEntry) {
		Dispatch.put(this, "AddressEntry", addressEntry);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getAutoResponse() {
		return Dispatch.get(this, "AutoResponse").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param autoResponse an input-parameter of type String
	 */
	public void setAutoResponse(String autoResponse) {
		Dispatch.put(this, "AutoResponse", autoResponse);
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
	public String getEntryID() {
		return Dispatch.get(this, "EntryID").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getIndex() {
		return Dispatch.get(this, "Index").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getMeetingResponseStatus() {
		return Dispatch.get(this, "MeetingResponseStatus").changeType(Variant.VariantInt).getInt();
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
	 * @return the result is of type boolean
	 */
	public boolean getResolved() {
		return Dispatch.get(this, "Resolved").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getTrackingStatus() {
		return Dispatch.get(this, "TrackingStatus").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param trackingStatus an input-parameter of type int
	 */
	public void setTrackingStatus(int trackingStatus) {
		Dispatch.put(this, "TrackingStatus", new Variant(trackingStatus));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type java.util.Date
	 */
	public java.util.Date getTrackingStatusTime() {
		return Dispatch.get(this, "TrackingStatusTime").getJavaDate();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param trackingStatusTime an input-parameter of type java.util.Date
	 */
	public void setTrackingStatusTime(java.util.Date trackingStatusTime) {
		Dispatch.put(this, "TrackingStatusTime", new Variant(trackingStatusTime));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getType() {
		return Dispatch.get(this, "Type").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param type an input-parameter of type int
	 */
	public void setType(int type) {
		Dispatch.put(this, "Type", new Variant(type));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void delete() {
		Dispatch.call(this, "Delete");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param start an input-parameter of type java.util.Date
	 * @param minPerChar an input-parameter of type int
	 * @param completeFormat an input-parameter of type Variant
	 * @return the result is of type String
	 */
	public String freeBusy(java.util.Date start, int minPerChar, Variant completeFormat) {
		return Dispatch.call(this, "FreeBusy", new Variant(start), new Variant(minPerChar), completeFormat).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param start an input-parameter of type java.util.Date
	 * @param minPerChar an input-parameter of type int
	 * @return the result is of type String
	 */
	public String freeBusy(java.util.Date start, int minPerChar) {
		return Dispatch.call(this, "FreeBusy", new Variant(start), new Variant(minPerChar)).toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean resolve() {
		return Dispatch.call(this, "Resolve").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type PropertyAccessor
	 */
	public PropertyAccessor getPropertyAccessor() {
		return new PropertyAccessor(Dispatch.get(this, "PropertyAccessor").toDispatch());
	}

}
