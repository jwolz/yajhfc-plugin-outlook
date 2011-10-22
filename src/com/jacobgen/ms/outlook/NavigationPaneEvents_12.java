/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class NavigationPaneEvents_12 extends Dispatch {

	public static final String componentName = "Outlook.NavigationPaneEvents_12";

	public NavigationPaneEvents_12() {
		super(componentName);
	}

	/**
	* This constructor is used instead of a case operation to
	* turn a Dispatch object into a wider object - it must exist
	* in every wrapper class whose instances may be returned from
	* method calls wrapped in VT_DISPATCH Variants.
	*/
	public NavigationPaneEvents_12(Dispatch d) {
		// take over the IDispatch pointer
		m_pDispatch = d.m_pDispatch;
		// null out the input's pointer
		d.m_pDispatch = 0;
	}

	public NavigationPaneEvents_12(String compName) {
		super(compName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param currentModule an input-parameter of type NavigationModule
	 */
	public void moduleSwitch(NavigationModule currentModule) {
		Dispatch.call(this, "ModuleSwitch", currentModule);
	}

}
