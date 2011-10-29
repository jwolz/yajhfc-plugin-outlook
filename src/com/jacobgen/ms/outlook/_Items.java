/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class _Items extends CachingDispatch {

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
		return new _Application(Dispatch.get(this, getIDOfName("Application")).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getClass1() {
		return Dispatch.get(this, getIDOfName("Class")).changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type _NameSpace
	 */
	public _NameSpace getSession() {
		return new _NameSpace(Dispatch.get(this, getIDOfName("Session")).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Object
	 */
	public Variant getParent() {
		return Dispatch.get(this, getIDOfName("Parent"));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getCount() {
		return Dispatch.get(this, getIDOfName("Count")).changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param index an input-parameter of type Variant
	 * @return the result is of type Object
	 */
	public Variant item(Integer index) {
		return Dispatch.call(this, getIDOfName("Item"), index);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Variant
	 */
	public Variant getRawTable() {
		return Dispatch.get(this, getIDOfName("RawTable"));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getIncludeRecurrences() {
		return Dispatch.get(this, getIDOfName("IncludeRecurrences")).changeType(Variant.VariantBoolean).getBoolean();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param includeRecurrences an input-parameter of type boolean
	 */
	public void setIncludeRecurrences(boolean includeRecurrences) {
		Dispatch.put(this, getIDOfName("IncludeRecurrences"), new Variant(includeRecurrences));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param type an input-parameter of type Variant
	 * @return the result is of type Object
	 */
	public Variant add(Variant type) {
		return Dispatch.call(this, getIDOfName("Add"), type);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Object
	 */
	public Variant add() {
		return Dispatch.call(this, getIDOfName("Add"));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param filter an input-parameter of type String
	 * @return the result is of type Object
	 */
	public Variant find(String filter) {
		return Dispatch.call(this, getIDOfName("Find"), filter);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Object
	 */
	public Variant findNext() {
		return Dispatch.call(this, getIDOfName("FindNext"));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Object
	 */
	public Variant getFirst() {
		return Dispatch.call(this, getIDOfName("GetFirst"));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Object
	 */
	public Variant getLast() {
		return Dispatch.call(this, getIDOfName("GetLast"));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Object
	 */
	public Variant getNext() {
		return Dispatch.call(this, getIDOfName("GetNext"));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Object
	 */
	public Variant getPrevious() {
		return Dispatch.call(this, getIDOfName("GetPrevious"));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param index an input-parameter of type int
	 */
	public void remove(int index) {
		Dispatch.call(this, getIDOfName("Remove"), new Variant(index));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void resetColumns() {
		Dispatch.call(this, getIDOfName("ResetColumns"));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param filter an input-parameter of type String
	 * @return the result is of type _Items
	 */
	public _Items restrict(String filter) {
		return new _Items(Dispatch.call(this, getIDOfName("Restrict"), filter).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param columns an input-parameter of type String
	 */
	public void setColumns(String columns) {
		Dispatch.call(this, getIDOfName("SetColumns"), columns);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param property an input-parameter of type String
	 * @param descending an input-parameter of type Variant
	 */
	public void sort(String property, Variant descending) {
		Dispatch.call(this, getIDOfName("Sort"), property, descending);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param property an input-parameter of type String
	 */
	public void sort(String property) {
		Dispatch.call(this, getIDOfName("Sort"), property);
	}

}
