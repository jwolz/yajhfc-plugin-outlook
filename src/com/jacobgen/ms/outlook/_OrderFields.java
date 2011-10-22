/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class _OrderFields extends Dispatch {

	public static final String componentName = "Outlook._OrderFields";

	public _OrderFields() {
		super(componentName);
	}

	/**
	* This constructor is used instead of a case operation to
	* turn a Dispatch object into a wider object - it must exist
	* in every wrapper class whose instances may be returned from
	* method calls wrapped in VT_DISPATCH Variants.
	*/
	public _OrderFields(Dispatch d) {
		// take over the IDispatch pointer
		m_pDispatch = d.m_pDispatch;
		// null out the input's pointer
		d.m_pDispatch = 0;
	}

	public _OrderFields(String compName) {
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
	 * @return the result is of type _OrderField
	 */
	public _OrderField item(Variant index) {
		return new _OrderField(Dispatch.call(this, "Item", index).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param propertyName an input-parameter of type String
	 * @param isDescending an input-parameter of type Variant
	 * @return the result is of type OrderField
	 */
	public OrderField add(String propertyName, Variant isDescending) {
		return new OrderField(Dispatch.call(this, "Add", propertyName, isDescending).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param propertyName an input-parameter of type String
	 * @return the result is of type OrderField
	 */
	public OrderField add(String propertyName) {
		return new OrderField(Dispatch.call(this, "Add", propertyName).toDispatch());
	}


	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param index an input-parameter of type Variant
	 */
	public void remove(Variant index) {
		Dispatch.call(this, "Remove", index);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void removeAll() {
		Dispatch.call(this, "RemoveAll");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param propertyName an input-parameter of type String
	 * @param index an input-parameter of type Variant
	 * @param isDescending an input-parameter of type Variant
	 * @return the result is of type OrderField
	 */
	public OrderField insert(String propertyName, Variant index, Variant isDescending) {
		return new OrderField(Dispatch.call(this, "Insert", propertyName, index, isDescending).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param propertyName an input-parameter of type String
	 * @param index an input-parameter of type Variant
	 * @return the result is of type OrderField
	 */
	public OrderField insert(String propertyName, Variant index) {
		return new OrderField(Dispatch.call(this, "Insert", propertyName, index).toDispatch());
	}



}
