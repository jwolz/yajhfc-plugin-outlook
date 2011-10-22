/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class _NavigationPane extends Dispatch {

	public static final String componentName = "Outlook._NavigationPane";

	public _NavigationPane() {
		super(componentName);
	}

	/**
	* This constructor is used instead of a case operation to
	* turn a Dispatch object into a wider object - it must exist
	* in every wrapper class whose instances may be returned from
	* method calls wrapped in VT_DISPATCH Variants.
	*/
	public _NavigationPane(Dispatch d) {
		// take over the IDispatch pointer
		m_pDispatch = d.m_pDispatch;
		// null out the input's pointer
		d.m_pDispatch = 0;
	}

	public _NavigationPane(String compName) {
		super(compName);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type _Application
	 */
	public _Application getApplication() {
		return new _Application(Dispatch.get(this, "Application").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getClass1() {
		return Dispatch.get(this, "Class").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type _NameSpace
	 */
	public _NameSpace getSession() {
		return new _NameSpace(Dispatch.get(this, "Session").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Object
	 */
	public Object getParent() {
		return Dispatch.get(this, "Parent");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getIsCollapsed() {
		return Dispatch.get(this, "IsCollapsed").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param isCollapsed an input-parameter of type boolean
	 */
	public void setIsCollapsed(boolean isCollapsed) {
		Dispatch.put(this, "IsCollapsed", new Variant(isCollapsed));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type NavigationModule
	 */
	public NavigationModule getCurrentModule() {
		return new NavigationModule(Dispatch.get(this, "CurrentModule").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param currentModule an input-parameter of type NavigationModule
	 */
	public void setCurrentModule(NavigationModule currentModule) {
		Dispatch.put(this, "CurrentModule", currentModule);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getDisplayedModuleCount() {
		return Dispatch.get(this, "DisplayedModuleCount").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param displayedModuleCount an input-parameter of type int
	 */
	public void setDisplayedModuleCount(int displayedModuleCount) {
		Dispatch.put(this, "DisplayedModuleCount", new Variant(displayedModuleCount));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type NavigationModules
	 */
	public NavigationModules getModules() {
		return new NavigationModules(Dispatch.get(this, "Modules").toDispatch());
	}

}
