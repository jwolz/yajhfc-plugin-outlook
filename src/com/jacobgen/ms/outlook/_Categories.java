/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class _Categories extends Dispatch {

	public static final String componentName = "Outlook._Categories";

	public _Categories() {
		super(componentName);
	}

	/**
	* This constructor is used instead of a case operation to
	* turn a Dispatch object into a wider object - it must exist
	* in every wrapper class whose instances may be returned from
	* method calls wrapped in VT_DISPATCH Variants.
	*/
	public _Categories(Dispatch d) {
		// take over the IDispatch pointer
		m_pDispatch = d.m_pDispatch;
		// null out the input's pointer
		d.m_pDispatch = 0;
	}

	public _Categories(String compName) {
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
	 * @return the result is of type _Category
	 */
	public _Category item(Variant index) {
		return new _Category(Dispatch.call(this, "Item", index).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param name an input-parameter of type String
	 * @param color an input-parameter of type Variant
	 * @param shortcutKey an input-parameter of type Variant
	 * @return the result is of type Category
	 */
	public Category add(String name, Variant color, Variant shortcutKey) {
		return new Category(Dispatch.call(this, "Add", name, color, shortcutKey).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param name an input-parameter of type String
	 * @param color an input-parameter of type Variant
	 * @return the result is of type Category
	 */
	public Category add(String name, Variant color) {
		return new Category(Dispatch.call(this, "Add", name, color).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param name an input-parameter of type String
	 * @return the result is of type Category
	 */
	public Category add(String name) {
		return new Category(Dispatch.call(this, "Add", name).toDispatch());
	}

//	/**
//	 * Wrapper for calling the ActiveX-Method and receiving the output-parameter(s).
//	 * @param name an input-parameter of type String
//	 * @param color an input-parameter of type Variant
//	 * @param shortcutKey an input-parameter of type Variant
//	 * @return the result is of type Category
//	 */
//	public Category add(String name, Variant color, Variant shortcutKey) {
//		Category result_of_Add = new Category(Dispatch.call(this, "Add", name, color, shortcutKey).toDispatch());
//
//
//		return result_of_Add;
//	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param index an input-parameter of type Variant
	 */
	public void remove(Variant index) {
		Dispatch.call(this, "Remove", index);
	}

}
