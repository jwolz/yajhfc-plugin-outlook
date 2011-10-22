/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class FormRegionEvents extends Dispatch {

	public static final String componentName = "Outlook.FormRegionEvents";

	public FormRegionEvents() {
		super(componentName);
	}

	/**
	* This constructor is used instead of a case operation to
	* turn a Dispatch object into a wider object - it must exist
	* in every wrapper class whose instances may be returned from
	* method calls wrapped in VT_DISPATCH Variants.
	*/
	public FormRegionEvents(Dispatch d) {
		// take over the IDispatch pointer
		m_pDispatch = d.m_pDispatch;
		// null out the input's pointer
		d.m_pDispatch = 0;
	}

	public FormRegionEvents(String compName) {
		super(compName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param expand an input-parameter of type boolean
	 */
	public void expanded(boolean expand) {
		Dispatch.call(this, "Expanded", new Variant(expand));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void close() {
		Dispatch.call(this, "Close");
	}

}
