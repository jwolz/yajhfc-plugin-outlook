/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class ItemsEvents extends Dispatch {

	public static final String componentName = "Outlook.ItemsEvents";

	public ItemsEvents() {
		super(componentName);
	}

	/**
	* This constructor is used instead of a case operation to
	* turn a Dispatch object into a wider object - it must exist
	* in every wrapper class whose instances may be returned from
	* method calls wrapped in VT_DISPATCH Variants.
	*/
	public ItemsEvents(Dispatch d) {
		// take over the IDispatch pointer
		m_pDispatch = d.m_pDispatch;
		// null out the input's pointer
		d.m_pDispatch = 0;
	}

	public ItemsEvents(String compName) {
		super(compName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param item an input-parameter of type Object
	 */
	public void itemAdd(Object item) {
		Dispatch.call(this, "ItemAdd", item);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param item an input-parameter of type Object
	 */
	public void itemChange(Object item) {
		Dispatch.call(this, "ItemChange", item);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void itemRemove() {
		Dispatch.call(this, "ItemRemove");
	}

}
