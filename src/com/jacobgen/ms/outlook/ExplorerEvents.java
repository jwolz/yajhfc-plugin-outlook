/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class ExplorerEvents extends Dispatch {

	public static final String componentName = "Outlook.ExplorerEvents";

	public ExplorerEvents() {
		super(componentName);
	}

	/**
	* This constructor is used instead of a case operation to
	* turn a Dispatch object into a wider object - it must exist
	* in every wrapper class whose instances may be returned from
	* method calls wrapped in VT_DISPATCH Variants.
	*/
	public ExplorerEvents(Dispatch d) {
		// take over the IDispatch pointer
		m_pDispatch = d.m_pDispatch;
		// null out the input's pointer
		d.m_pDispatch = 0;
	}

	public ExplorerEvents(String compName) {
		super(compName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void activate() {
		Dispatch.call(this, "Activate");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void folderSwitch() {
		Dispatch.call(this, "FolderSwitch");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param newFolder an input-parameter of type Object
	 * @param cancel an input-parameter of type boolean
	 */
	public void beforeFolderSwitch(Object newFolder, boolean cancel) {
		Dispatch.call(this, "BeforeFolderSwitch", newFolder, new Variant(cancel));
	}

	/**
	 * Wrapper for calling the ActiveX-Method and receiving the output-parameter(s).
	 * @param newFolder an input-parameter of type Object
	 * @param cancel is an one-element array which sends the input-parameter
	 *               to the ActiveX-Component and receives the output-parameter
	 */
	public void beforeFolderSwitch(Object newFolder, boolean[] cancel) {
		Variant vnt_cancel = new Variant();
		if( cancel == null || cancel.length == 0 )
			vnt_cancel.putNoParam();
		else
			vnt_cancel.putBooleanRef(cancel[0]);

		Dispatch.call(this, "BeforeFolderSwitch", newFolder, vnt_cancel);

		if( cancel != null && cancel.length > 0 )
			cancel[0] = vnt_cancel.toBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void viewSwitch() {
		Dispatch.call(this, "ViewSwitch");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param newView an input-parameter of type Variant
	 * @param cancel an input-parameter of type boolean
	 */
	public void beforeViewSwitch(Variant newView, boolean cancel) {
		Dispatch.call(this, "BeforeViewSwitch", newView, new Variant(cancel));
	}

	/**
	 * Wrapper for calling the ActiveX-Method and receiving the output-parameter(s).
	 * @param newView an input-parameter of type Variant
	 * @param cancel is an one-element array which sends the input-parameter
	 *               to the ActiveX-Component and receives the output-parameter
	 */
	public void beforeViewSwitch(Variant newView, boolean[] cancel) {
		Variant vnt_cancel = new Variant();
		if( cancel == null || cancel.length == 0 )
			vnt_cancel.putNoParam();
		else
			vnt_cancel.putBooleanRef(cancel[0]);

		Dispatch.call(this, "BeforeViewSwitch", newView, vnt_cancel);

		if( cancel != null && cancel.length > 0 )
			cancel[0] = vnt_cancel.toBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void deactivate() {
		Dispatch.call(this, "Deactivate");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void selectionChange() {
		Dispatch.call(this, "SelectionChange");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void close() {
		Dispatch.call(this, "Close");
	}

}
