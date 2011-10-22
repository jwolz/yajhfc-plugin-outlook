/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class OutlookBarPaneEvents extends Dispatch {

	public static final String componentName = "Outlook.OutlookBarPaneEvents";

	public OutlookBarPaneEvents() {
		super(componentName);
	}

	/**
	* This constructor is used instead of a case operation to
	* turn a Dispatch object into a wider object - it must exist
	* in every wrapper class whose instances may be returned from
	* method calls wrapped in VT_DISPATCH Variants.
	*/
	public OutlookBarPaneEvents(Dispatch d) {
		// take over the IDispatch pointer
		m_pDispatch = d.m_pDispatch;
		// null out the input's pointer
		d.m_pDispatch = 0;
	}

	public OutlookBarPaneEvents(String compName) {
		super(compName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param shortcut an input-parameter of type OutlookBarShortcut
	 * @param cancel an input-parameter of type boolean
	 */
	public void beforeNavigate(OutlookBarShortcut shortcut, boolean cancel) {
		Dispatch.call(this, "BeforeNavigate", shortcut, new Variant(cancel));
	}

	/**
	 * Wrapper for calling the ActiveX-Method and receiving the output-parameter(s).
	 * @param shortcut an input-parameter of type OutlookBarShortcut
	 * @param cancel is an one-element array which sends the input-parameter
	 *               to the ActiveX-Component and receives the output-parameter
	 */
	public void beforeNavigate(OutlookBarShortcut shortcut, boolean[] cancel) {
		Variant vnt_cancel = new Variant();
		if( cancel == null || cancel.length == 0 )
			vnt_cancel.putNoParam();
		else
			vnt_cancel.putBooleanRef(cancel[0]);

		Dispatch.call(this, "BeforeNavigate", shortcut, vnt_cancel);

		if( cancel != null && cancel.length > 0 )
			cancel[0] = vnt_cancel.toBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param toGroup an input-parameter of type OutlookBarGroup
	 * @param cancel an input-parameter of type boolean
	 */
	public void beforeGroupSwitch(OutlookBarGroup toGroup, boolean cancel) {
		Dispatch.call(this, "BeforeGroupSwitch", toGroup, new Variant(cancel));
	}

	/**
	 * Wrapper for calling the ActiveX-Method and receiving the output-parameter(s).
	 * @param toGroup an input-parameter of type OutlookBarGroup
	 * @param cancel is an one-element array which sends the input-parameter
	 *               to the ActiveX-Component and receives the output-parameter
	 */
	public void beforeGroupSwitch(OutlookBarGroup toGroup, boolean[] cancel) {
		Variant vnt_cancel = new Variant();
		if( cancel == null || cancel.length == 0 )
			vnt_cancel.putNoParam();
		else
			vnt_cancel.putBooleanRef(cancel[0]);

		Dispatch.call(this, "BeforeGroupSwitch", toGroup, vnt_cancel);

		if( cancel != null && cancel.length > 0 )
			cancel[0] = vnt_cancel.toBoolean();
	}

}
