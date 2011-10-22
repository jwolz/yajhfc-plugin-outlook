/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class MAPIFolderEvents_12 extends Dispatch {

	public static final String componentName = "Outlook.MAPIFolderEvents_12";

	public MAPIFolderEvents_12() {
		super(componentName);
	}

	/**
	* This constructor is used instead of a case operation to
	* turn a Dispatch object into a wider object - it must exist
	* in every wrapper class whose instances may be returned from
	* method calls wrapped in VT_DISPATCH Variants.
	*/
	public MAPIFolderEvents_12(Dispatch d) {
		// take over the IDispatch pointer
		m_pDispatch = d.m_pDispatch;
		// null out the input's pointer
		d.m_pDispatch = 0;
	}

	public MAPIFolderEvents_12(String compName) {
		super(compName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param moveTo an input-parameter of type MAPIFolder
	 * @param cancel an input-parameter of type boolean
	 */
	public void beforeFolderMove(MAPIFolder moveTo, boolean cancel) {
		Dispatch.call(this, "BeforeFolderMove", moveTo, new Variant(cancel));
	}

	/**
	 * Wrapper for calling the ActiveX-Method and receiving the output-parameter(s).
	 * @param moveTo an input-parameter of type MAPIFolder
	 * @param cancel is an one-element array which sends the input-parameter
	 *               to the ActiveX-Component and receives the output-parameter
	 */
	public void beforeFolderMove(MAPIFolder moveTo, boolean[] cancel) {
		Variant vnt_cancel = new Variant();
		if( cancel == null || cancel.length == 0 )
			vnt_cancel.putNoParam();
		else
			vnt_cancel.putBooleanRef(cancel[0]);

		Dispatch.call(this, "BeforeFolderMove", moveTo, vnt_cancel);

		if( cancel != null && cancel.length > 0 )
			cancel[0] = vnt_cancel.toBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param item an input-parameter of type Object
	 * @param moveTo an input-parameter of type MAPIFolder
	 * @param cancel an input-parameter of type boolean
	 */
	public void beforeItemMove(Object item, MAPIFolder moveTo, boolean cancel) {
		Dispatch.call(this, "BeforeItemMove", item, moveTo, new Variant(cancel));
	}

	/**
	 * Wrapper for calling the ActiveX-Method and receiving the output-parameter(s).
	 * @param item an input-parameter of type Object
	 * @param moveTo an input-parameter of type MAPIFolder
	 * @param cancel is an one-element array which sends the input-parameter
	 *               to the ActiveX-Component and receives the output-parameter
	 */
	public void beforeItemMove(Object item, MAPIFolder moveTo, boolean[] cancel) {
		Variant vnt_cancel = new Variant();
		if( cancel == null || cancel.length == 0 )
			vnt_cancel.putNoParam();
		else
			vnt_cancel.putBooleanRef(cancel[0]);

		Dispatch.call(this, "BeforeItemMove", item, moveTo, vnt_cancel);

		if( cancel != null && cancel.length > 0 )
			cancel[0] = vnt_cancel.toBoolean();
	}

}
