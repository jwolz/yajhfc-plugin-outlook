/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class SyncObjectEvents extends Dispatch {

	public static final String componentName = "Outlook.SyncObjectEvents";

	public SyncObjectEvents() {
		super(componentName);
	}

	/**
	* This constructor is used instead of a case operation to
	* turn a Dispatch object into a wider object - it must exist
	* in every wrapper class whose instances may be returned from
	* method calls wrapped in VT_DISPATCH Variants.
	*/
	public SyncObjectEvents(Dispatch d) {
		// take over the IDispatch pointer
		m_pDispatch = d.m_pDispatch;
		// null out the input's pointer
		d.m_pDispatch = 0;
	}

	public SyncObjectEvents(String compName) {
		super(compName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void syncStart() {
		Dispatch.call(this, "SyncStart");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param state an input-parameter of type int
	 * @param description an input-parameter of type String
	 * @param value an input-parameter of type int
	 * @param max an input-parameter of type int
	 */
	public void progress(int state, String description, int value, int max) {
		Dispatch.call(this, "Progress", new Variant(state), description, new Variant(value), new Variant(max));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param code an input-parameter of type int
	 * @param description an input-parameter of type String
	 */
	public void onError(int code, String description) {
		Dispatch.call(this, "OnError", new Variant(code), description);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void syncEnd() {
		Dispatch.call(this, "SyncEnd");
	}

}
