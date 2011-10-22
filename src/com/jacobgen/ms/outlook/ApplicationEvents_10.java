/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class ApplicationEvents_10 extends Dispatch {

	public static final String componentName = "Outlook.ApplicationEvents_10";

	public ApplicationEvents_10() {
		super(componentName);
	}

	/**
	* This constructor is used instead of a case operation to
	* turn a Dispatch object into a wider object - it must exist
	* in every wrapper class whose instances may be returned from
	* method calls wrapped in VT_DISPATCH Variants.
	*/
	public ApplicationEvents_10(Dispatch d) {
		// take over the IDispatch pointer
		m_pDispatch = d.m_pDispatch;
		// null out the input's pointer
		d.m_pDispatch = 0;
	}

	public ApplicationEvents_10(String compName) {
		super(compName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param item an input-parameter of type Object
	 * @param cancel an input-parameter of type boolean
	 * @return the result is of type int
	 */
	public int itemSend(Object item, boolean cancel) {
		return Dispatch.call(this, "ItemSend", item, new Variant(cancel)).changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method and receiving the output-parameter(s).
	 * @param item an input-parameter of type Object
	 * @param cancel is an one-element array which sends the input-parameter
	 *               to the ActiveX-Component and receives the output-parameter
	 * @return the result is of type int
	 */
	public int itemSend(Object item, boolean[] cancel) {
		Variant vnt_cancel = new Variant();
		if( cancel == null || cancel.length == 0 )
			vnt_cancel.putNoParam();
		else
			vnt_cancel.putBooleanRef(cancel[0]);

		int result_of_ItemSend = Dispatch.call(this, "ItemSend", item, vnt_cancel).changeType(Variant.VariantInt).getInt();

		if( cancel != null && cancel.length > 0 )
			cancel[0] = vnt_cancel.toBoolean();

		return result_of_ItemSend;
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int newMail() {
		return Dispatch.call(this, "NewMail").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param item an input-parameter of type Object
	 * @return the result is of type int
	 */
	public int reminder(Object item) {
		return Dispatch.call(this, "Reminder", item).changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param pages an input-parameter of type PropertyPages
	 * @return the result is of type int
	 */
	public int optionsPagesAdd(PropertyPages pages) {
		return Dispatch.call(this, "OptionsPagesAdd", pages).changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int startup() {
		return Dispatch.call(this, "Startup").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int quit() {
		return Dispatch.call(this, "Quit").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param searchObject an input-parameter of type Search
	 */
	public void advancedSearchComplete(Search searchObject) {
		Dispatch.call(this, "AdvancedSearchComplete", searchObject);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param searchObject an input-parameter of type Search
	 */
	public void advancedSearchStopped(Search searchObject) {
		Dispatch.call(this, "AdvancedSearchStopped", searchObject);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void mAPILogonComplete() {
		Dispatch.call(this, "MAPILogonComplete");
	}

}
