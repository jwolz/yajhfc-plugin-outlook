/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class _ViewsEvents extends Dispatch {

	public static final String componentName = "Outlook._ViewsEvents";

	public _ViewsEvents() {
		super(componentName);
	}

	/**
	* This constructor is used instead of a case operation to
	* turn a Dispatch object into a wider object - it must exist
	* in every wrapper class whose instances may be returned from
	* method calls wrapped in VT_DISPATCH Variants.
	*/
	public _ViewsEvents(Dispatch d) {
		// take over the IDispatch pointer
		m_pDispatch = d.m_pDispatch;
		// null out the input's pointer
		d.m_pDispatch = 0;
	}

	public _ViewsEvents(String compName) {
		super(compName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param view an input-parameter of type View
	 */
	public void viewAdd(View view) {
		Dispatch.call(this, "ViewAdd", view);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param view an input-parameter of type View
	 */
	public void viewRemove(View view) {
		Dispatch.call(this, "ViewRemove", view);
	}

}
