/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class PropertyPages extends Dispatch {

	public static final String componentName = "Outlook.PropertyPages";

	public PropertyPages() {
		super(componentName);
	}

	/**
	* This constructor is used instead of a case operation to
	* turn a Dispatch object into a wider object - it must exist
	* in every wrapper class whose instances may be returned from
	* method calls wrapped in VT_DISPATCH Variants.
	*/
	public PropertyPages(Dispatch d) {
		// take over the IDispatch pointer
		m_pDispatch = d.m_pDispatch;
		// null out the input's pointer
		d.m_pDispatch = 0;
	}

	public PropertyPages(String compName) {
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
	 * @return the result is of type Object
	 */
	public Object item(Variant index) {
		return Dispatch.call(this, "Item", index);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param page an input-parameter of type Variant
	 * @param title an input-parameter of type String
	 */
	public void add(Variant page, String title) {
		Dispatch.call(this, "Add", page, title);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param page an input-parameter of type Variant
	 */
	public void add(Variant page) {
		Dispatch.call(this, "Add", page);
	}


	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param index an input-parameter of type Variant
	 */
	public void remove(Variant index) {
		Dispatch.call(this, "Remove", index);
	}

}
