/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class _SelectNamesDialog extends Dispatch {

	public static final String componentName = "Outlook._SelectNamesDialog";

	public _SelectNamesDialog() {
		super(componentName);
	}

	/**
	* This constructor is used instead of a case operation to
	* turn a Dispatch object into a wider object - it must exist
	* in every wrapper class whose instances may be returned from
	* method calls wrapped in VT_DISPATCH Variants.
	*/
	public _SelectNamesDialog(Dispatch d) {
		// take over the IDispatch pointer
		m_pDispatch = d.m_pDispatch;
		// null out the input's pointer
		d.m_pDispatch = 0;
	}

	public _SelectNamesDialog(String compName) {
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
	public String getCaption() {
		return Dispatch.get(this, "Caption").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param caption an input-parameter of type String
	 */
	public void setCaption(String caption) {
		Dispatch.put(this, "Caption", caption);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean display() {
		return Dispatch.call(this, "Display").changeType(Variant.VariantBoolean).getBoolean();
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
	 * @param recipients an input-parameter of type Recipients
	 */
	public void setRecipients(Recipients recipients) {
		Dispatch.put(this, "Recipients", recipients);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getBccLabel() {
		return Dispatch.get(this, "BccLabel").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param bccLabel an input-parameter of type String
	 */
	public void setBccLabel(String bccLabel) {
		Dispatch.put(this, "BccLabel", bccLabel);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getCcLabel() {
		return Dispatch.get(this, "CcLabel").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param ccLabel an input-parameter of type String
	 */
	public void setCcLabel(String ccLabel) {
		Dispatch.put(this, "CcLabel", ccLabel);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type String
	 */
	public String getToLabel() {
		return Dispatch.get(this, "ToLabel").toString();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param toLabel an input-parameter of type String
	 */
	public void setToLabel(String toLabel) {
		Dispatch.put(this, "ToLabel", toLabel);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getAllowMultipleSelection() {
		return Dispatch.get(this, "AllowMultipleSelection").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param allowMultipleSelection an input-parameter of type boolean
	 */
	public void setAllowMultipleSelection(boolean allowMultipleSelection) {
		Dispatch.put(this, "AllowMultipleSelection", new Variant(allowMultipleSelection));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getForceResolution() {
		return Dispatch.get(this, "ForceResolution").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param forceResolution an input-parameter of type boolean
	 */
	public void setForceResolution(boolean forceResolution) {
		Dispatch.put(this, "ForceResolution", new Variant(forceResolution));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getShowOnlyInitialAddressList() {
		return Dispatch.get(this, "ShowOnlyInitialAddressList").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param showOnlyInitialAddressList an input-parameter of type boolean
	 */
	public void setShowOnlyInitialAddressList(boolean showOnlyInitialAddressList) {
		Dispatch.put(this, "ShowOnlyInitialAddressList", new Variant(showOnlyInitialAddressList));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getNumberOfRecipientSelectors() {
		return Dispatch.get(this, "NumberOfRecipientSelectors").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param numberOfRecipientSelectors an input-parameter of type int
	 */
	public void setNumberOfRecipientSelectors(int numberOfRecipientSelectors) {
		Dispatch.put(this, "NumberOfRecipientSelectors", new Variant(numberOfRecipientSelectors));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type AddressList
	 */
	public AddressList getInitialAddressList() {
		return new AddressList(Dispatch.get(this, "InitialAddressList").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param initialAddressList an input-parameter of type AddressList
	 */
	public void setInitialAddressList(AddressList initialAddressList) {
		Dispatch.put(this, "InitialAddressList", initialAddressList);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param defaultMode an input-parameter of type int
	 */
	public void setDefaultDisplayMode(int defaultMode) {
		Dispatch.call(this, "SetDefaultDisplayMode", new Variant(defaultMode));
	}

}
