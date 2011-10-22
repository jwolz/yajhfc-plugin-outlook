/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class ReminderCollectionEvents extends Dispatch {

	public static final String componentName = "Outlook.ReminderCollectionEvents";

	public ReminderCollectionEvents() {
		super(componentName);
	}

	/**
	* This constructor is used instead of a case operation to
	* turn a Dispatch object into a wider object - it must exist
	* in every wrapper class whose instances may be returned from
	* method calls wrapped in VT_DISPATCH Variants.
	*/
	public ReminderCollectionEvents(Dispatch d) {
		// take over the IDispatch pointer
		m_pDispatch = d.m_pDispatch;
		// null out the input's pointer
		d.m_pDispatch = 0;
	}

	public ReminderCollectionEvents(String compName) {
		super(compName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param cancel an input-parameter of type boolean
	 */
	public void beforeReminderShow(boolean cancel) {
		Dispatch.call(this, "BeforeReminderShow", new Variant(cancel));
	}

	/**
	 * Wrapper for calling the ActiveX-Method and receiving the output-parameter(s).
	 * @param cancel is an one-element array which sends the input-parameter
	 *               to the ActiveX-Component and receives the output-parameter
	 */
	public void beforeReminderShow(boolean[] cancel) {
		Variant vnt_cancel = new Variant();
		if( cancel == null || cancel.length == 0 )
			vnt_cancel.putNoParam();
		else
			vnt_cancel.putBooleanRef(cancel[0]);

		Dispatch.call(this, "BeforeReminderShow", vnt_cancel);

		if( cancel != null && cancel.length > 0 )
			cancel[0] = vnt_cancel.toBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param reminderObject an input-parameter of type _Reminder
	 */
	public void reminderAdd(_Reminder reminderObject) {
		Dispatch.call(this, "ReminderAdd", reminderObject);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param reminderObject an input-parameter of type _Reminder
	 */
	public void reminderChange(_Reminder reminderObject) {
		Dispatch.call(this, "ReminderChange", reminderObject);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param reminderObject an input-parameter of type _Reminder
	 */
	public void reminderFire(_Reminder reminderObject) {
		Dispatch.call(this, "ReminderFire", reminderObject);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void reminderRemove() {
		Dispatch.call(this, "ReminderRemove");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param reminderObject an input-parameter of type _Reminder
	 */
	public void snooze(_Reminder reminderObject) {
		Dispatch.call(this, "Snooze", reminderObject);
	}

}
