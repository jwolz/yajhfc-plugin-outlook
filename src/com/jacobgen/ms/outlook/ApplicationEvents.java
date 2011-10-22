/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class ApplicationEvents extends Dispatch {

	public static final String componentName = "Outlook.ApplicationEvents";

	public ApplicationEvents() {
		super(componentName);
	}

	/**
	* This constructor is used instead of a case operation to
	* turn a Dispatch object into a wider object - it must exist
	* in every wrapper class whose instances may be returned from
	* method calls wrapped in VT_DISPATCH Variants.
	*/
	public ApplicationEvents(Dispatch d) {
		// take over the IDispatch pointer
		m_pDispatch = d.m_pDispatch;
		// null out the input's pointer
		d.m_pDispatch = 0;
	}

	public ApplicationEvents(String compName) {
		super(compName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param item an input-parameter of type Object
	 * @param cancel an input-parameter of type boolean
	 */
	public void itemSend(Object item, boolean cancel) {
		Dispatch.call(this, "ItemSend", item, new Variant(cancel));
	}

	/**
	 * Wrapper for calling the ActiveX-Method and receiving the output-parameter(s).
	 * @param item an input-parameter of type Object
	 * @param cancel is an one-element array which sends the input-parameter
	 *               to the ActiveX-Component and receives the output-parameter
	 */
	public void itemSend(Object item, boolean[] cancel) {
		Variant vnt_cancel = new Variant();
		if( cancel == null || cancel.length == 0 )
			vnt_cancel.putNoParam();
		else
			vnt_cancel.putBooleanRef(cancel[0]);

		Dispatch.call(this, "ItemSend", item, vnt_cancel);

		if( cancel != null && cancel.length > 0 )
			cancel[0] = vnt_cancel.toBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void newMail() {
		Dispatch.call(this, "NewMail");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param item an input-parameter of type Object
	 */
	public void reminder(Object item) {
		Dispatch.call(this, "Reminder", item);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param pages an input-parameter of type PropertyPages
	 */
	public void optionsPagesAdd(PropertyPages pages) {
		Dispatch.call(this, "OptionsPagesAdd", pages);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void startup() {
		Dispatch.call(this, "Startup");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void quit() {
		Dispatch.call(this, "Quit");
	}

}
