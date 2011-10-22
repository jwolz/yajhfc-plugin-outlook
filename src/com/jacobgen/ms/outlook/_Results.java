/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class _Results extends Dispatch {

	public static final String componentName = "Outlook._Results";

	public _Results() {
		super(componentName);
	}

	/**
	* This constructor is used instead of a case operation to
	* turn a Dispatch object into a wider object - it must exist
	* in every wrapper class whose instances may be returned from
	* method calls wrapped in VT_DISPATCH Variants.
	*/
	public _Results(Dispatch d) {
		// take over the IDispatch pointer
		m_pDispatch = d.m_pDispatch;
		// null out the input's pointer
		d.m_pDispatch = 0;
	}

	public _Results(String compName) {
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
	 * @return the result is of type Variant
	 */
	public Variant getRawTable() {
		return Dispatch.get(this, "RawTable");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Object
	 */
	public Object getFirst() {
		return Dispatch.call(this, "GetFirst");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Object
	 */
	public Object getLast() {
		return Dispatch.call(this, "GetLast");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Object
	 */
	public Object getNext() {
		return Dispatch.call(this, "GetNext");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Object
	 */
	public Object getPrevious() {
		return Dispatch.call(this, "GetPrevious");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void resetColumns() {
		Dispatch.call(this, "ResetColumns");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param columns an input-parameter of type String
	 */
	public void setColumns(String columns) {
		Dispatch.call(this, "SetColumns", columns);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param property an input-parameter of type String
	 * @param descending an input-parameter of type Variant
	 */
	public void sort(String property, Variant descending) {
		Dispatch.call(this, "Sort", property, descending);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param property an input-parameter of type String
	 */
	public void sort(String property) {
		Dispatch.call(this, "Sort", property);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getDefaultItemType() {
		return Dispatch.get(this, "DefaultItemType").changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param defaultItemType an input-parameter of type int
	 */
	public void setDefaultItemType(int defaultItemType) {
		Dispatch.put(this, "DefaultItemType", new Variant(defaultItemType));
	}

}
