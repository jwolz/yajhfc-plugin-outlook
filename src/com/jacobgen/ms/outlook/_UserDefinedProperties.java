/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class _UserDefinedProperties extends Dispatch {

	public static final String componentName = "Outlook._UserDefinedProperties";

	public _UserDefinedProperties() {
		super(componentName);
	}

	/**
	* This constructor is used instead of a case operation to
	* turn a Dispatch object into a wider object - it must exist
	* in every wrapper class whose instances may be returned from
	* method calls wrapped in VT_DISPATCH Variants.
	*/
	public _UserDefinedProperties(Dispatch d) {
		// take over the IDispatch pointer
		m_pDispatch = d.m_pDispatch;
		// null out the input's pointer
		d.m_pDispatch = 0;
	}

	public _UserDefinedProperties(String compName) {
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
	 * @return the result is of type _UserDefinedProperty
	 */
	public _UserDefinedProperty item(Variant index) {
		return new _UserDefinedProperty(Dispatch.call(this, "Item", index).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param name an input-parameter of type String
	 * @param type an input-parameter of type int
	 * @param displayFormat an input-parameter of type Variant
	 * @param formula an input-parameter of type Variant
	 * @return the result is of type UserDefinedProperty
	 */
	public UserDefinedProperty add(String name, int type, Variant displayFormat, Variant formula) {
		return new UserDefinedProperty(Dispatch.call(this, "Add", name, new Variant(type), displayFormat, formula).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param name an input-parameter of type String
	 * @param type an input-parameter of type int
	 * @param displayFormat an input-parameter of type Variant
	 * @return the result is of type UserDefinedProperty
	 */
	public UserDefinedProperty add(String name, int type, Variant displayFormat) {
		return new UserDefinedProperty(Dispatch.call(this, "Add", name, new Variant(type), displayFormat).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param name an input-parameter of type String
	 * @param type an input-parameter of type int
	 * @return the result is of type UserDefinedProperty
	 */
	public UserDefinedProperty add(String name, int type) {
		return new UserDefinedProperty(Dispatch.call(this, "Add", name, new Variant(type)).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param name an input-parameter of type String
	 * @return the result is of type UserDefinedProperty
	 */
	public UserDefinedProperty find(String name) {
		return new UserDefinedProperty(Dispatch.call(this, "Find", name).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param index an input-parameter of type int
	 */
	public void remove(int index) {
		Dispatch.call(this, "Remove", new Variant(index));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void refresh() {
		Dispatch.call(this, "Refresh");
	}

}
