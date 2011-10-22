/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class AddressEntries extends Dispatch {

	public static final String componentName = "Outlook.AddressEntries";

	public AddressEntries() {
		super(componentName);
	}

	/**
	* This constructor is used instead of a case operation to
	* turn a Dispatch object into a wider object - it must exist
	* in every wrapper class whose instances may be returned from
	* method calls wrapped in VT_DISPATCH Variants.
	*/
	public AddressEntries(Dispatch d) {
		// take over the IDispatch pointer
		m_pDispatch = d.m_pDispatch;
		// null out the input's pointer
		d.m_pDispatch = 0;
	}

	public AddressEntries(String compName) {
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
	 * @return the result is of type AddressEntry
	 */
	public AddressEntry item(Variant index) {
		return new AddressEntry(Dispatch.call(this, "Item", index).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Variant
	 */
	public Variant getRawTable() {
		return Dispatch.get(this, "RawTable");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param type an input-parameter of type String
	 * @param name an input-parameter of type Variant
	 * @param address an input-parameter of type Variant
	 * @return the result is of type AddressEntry
	 */
	public AddressEntry add(String type, Variant name, Variant address) {
		return new AddressEntry(Dispatch.call(this, "Add", type, name, address).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param type an input-parameter of type String
	 * @param name an input-parameter of type Variant
	 * @return the result is of type AddressEntry
	 */
	public AddressEntry add(String type, Variant name) {
		return new AddressEntry(Dispatch.call(this, "Add", type, name).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param type an input-parameter of type String
	 * @return the result is of type AddressEntry
	 */
	public AddressEntry add(String type) {
		return new AddressEntry(Dispatch.call(this, "Add", type).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type AddressEntry
	 */
	public AddressEntry getFirst() {
		return new AddressEntry(Dispatch.call(this, "GetFirst").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type AddressEntry
	 */
	public AddressEntry getLast() {
		return new AddressEntry(Dispatch.call(this, "GetLast").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type AddressEntry
	 */
	public AddressEntry getNext() {
		return new AddressEntry(Dispatch.call(this, "GetNext").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type AddressEntry
	 */
	public AddressEntry getPrevious() {
		return new AddressEntry(Dispatch.call(this, "GetPrevious").toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param property an input-parameter of type Variant
	 * @param order an input-parameter of type Variant
	 */
	public void sort(Variant property, Variant order) {
		Dispatch.call(this, "Sort", property, order);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param property an input-parameter of type Variant
	 */
	public void sort(Variant property) {
		Dispatch.call(this, "Sort", property);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void sort() {
		Dispatch.call(this, "Sort");
	}

}
