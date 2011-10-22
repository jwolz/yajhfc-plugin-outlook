/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class _Items extends Dispatch {

	public static final String componentName = "Outlook._Items";

	public _Items() {
		super(componentName);
	}

	/**
	* This constructor is used instead of a case operation to
	* turn a Dispatch object into a wider object - it must exist
	* in every wrapper class whose instances may be returned from
	* method calls wrapped in VT_DISPATCH Variants.
	*/
	public _Items(Dispatch d) {
		// take over the IDispatch pointer
		m_pDispatch = d.m_pDispatch;
		// null out the input's pointer
		d.m_pDispatch = 0;
	}

	public _Items(String compName) {
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
	public Variant getParent() {
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
	public Variant item(Integer index) {
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
	 * @return the result is of type boolean
	 */
	public boolean getIncludeRecurrences() {
		return Dispatch.get(this, "IncludeRecurrences").changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param includeRecurrences an input-parameter of type boolean
	 */
	public void setIncludeRecurrences(boolean includeRecurrences) {
		Dispatch.put(this, "IncludeRecurrences", new Variant(includeRecurrences));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param type an input-parameter of type Variant
	 * @return the result is of type Object
	 */
	public Variant add(Variant type) {
		return Dispatch.call(this, "Add", type);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Object
	 */
	public Variant add() {
		return Dispatch.call(this, "Add");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param filter an input-parameter of type String
	 * @return the result is of type Object
	 */
	public Variant find(String filter) {
		return Dispatch.call(this, "Find", filter);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Object
	 */
	public Variant findNext() {
		return Dispatch.call(this, "FindNext");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Object
	 */
	public Variant getFirst() {
		return Dispatch.call(this, "GetFirst");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Object
	 */
	public Variant getLast() {
		return Dispatch.call(this, "GetLast");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Object
	 */
	public Variant getNext() {
		return Dispatch.call(this, "GetNext");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Object
	 */
	public Variant getPrevious() {
		return Dispatch.call(this, "GetPrevious");
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
	public void resetColumns() {
		Dispatch.call(this, "ResetColumns");
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param filter an input-parameter of type String
	 * @return the result is of type _Items
	 */
	public _Items restrict(String filter) {
		return new _Items(Dispatch.call(this, "Restrict", filter).toDispatch());
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

}
