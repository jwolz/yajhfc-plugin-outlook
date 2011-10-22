/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class NavigationGroupsEvents_12 extends Dispatch {

	public static final String componentName = "Outlook.NavigationGroupsEvents_12";

	public NavigationGroupsEvents_12() {
		super(componentName);
	}

	/**
	* This constructor is used instead of a case operation to
	* turn a Dispatch object into a wider object - it must exist
	* in every wrapper class whose instances may be returned from
	* method calls wrapped in VT_DISPATCH Variants.
	*/
	public NavigationGroupsEvents_12(Dispatch d) {
		// take over the IDispatch pointer
		m_pDispatch = d.m_pDispatch;
		// null out the input's pointer
		d.m_pDispatch = 0;
	}

	public NavigationGroupsEvents_12(String compName) {
		super(compName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param navigationFolder an input-parameter of type NavigationFolder
	 */
	public void selectedChange(NavigationFolder navigationFolder) {
		Dispatch.call(this, "SelectedChange", navigationFolder);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param navigationFolder an input-parameter of type NavigationFolder
	 */
	public void navigationFolderAdd(NavigationFolder navigationFolder) {
		Dispatch.call(this, "NavigationFolderAdd", navigationFolder);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void navigationFolderRemove() {
		Dispatch.call(this, "NavigationFolderRemove");
	}

}
