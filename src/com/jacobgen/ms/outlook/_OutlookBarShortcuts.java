/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class _OutlookBarShortcuts extends Dispatch {

	public static final String componentName = "Outlook._OutlookBarShortcuts";

	public _OutlookBarShortcuts() {
		super(componentName);
	}

	/**
	* This constructor is used instead of a case operation to
	* turn a Dispatch object into a wider object - it must exist
	* in every wrapper class whose instances may be returned from
	* method calls wrapped in VT_DISPATCH Variants.
	*/
	public _OutlookBarShortcuts(Dispatch d) {
		// take over the IDispatch pointer
		m_pDispatch = d.m_pDispatch;
		// null out the input's pointer
		d.m_pDispatch = 0;
	}

	public _OutlookBarShortcuts(String compName) {
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
	 * @return the result is of type int
	 */
	public int getCount() {
		return Dispatch.get(this, "Count").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param index an input-parameter of type Variant
	 * @return the result is of type OutlookBarShortcut
	 */
	public OutlookBarShortcut item(Variant index) {
		return new OutlookBarShortcut(Dispatch.call(this, "Item", index).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param target an input-parameter of type Variant
	 * @param name an input-parameter of type String
	 * @param index an input-parameter of type Variant
	 * @return the result is of type OutlookBarShortcut
	 */
	public OutlookBarShortcut add(Variant target, String name, Variant index) {
		return new OutlookBarShortcut(Dispatch.call(this, "Add", target, name, index).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param target an input-parameter of type Variant
	 * @param name an input-parameter of type String
	 * @return the result is of type OutlookBarShortcut
	 */
	public OutlookBarShortcut add(Variant target, String name) {
		return new OutlookBarShortcut(Dispatch.call(this, "Add", target, name).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param index an input-parameter of type Variant
	 */
	public void remove(Variant index) {
		Dispatch.call(this, "Remove", index);
	}

}
