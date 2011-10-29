/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class _Table extends CachingDispatch {

	public static final String componentName = "Outlook._Table";

	public _Table() {
		super(componentName);
	}

	/**
	* This constructor is used instead of a case operation to
	* turn a Dispatch object into a wider object - it must exist
	* in every wrapper class whose instances may be returned from
	* method calls wrapped in VT_DISPATCH Variants.
	*/
	public _Table(Dispatch d) {
		// take over the IDispatch pointer
		m_pDispatch = d.m_pDispatch;
		// null out the input's pointer
		d.m_pDispatch = 0;
	}

	public _Table(String compName) {
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
	public Object getParent() {
		return Dispatch.get(this, getIDOfName("Parent"));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param filter an input-parameter of type String
	 * @return the result is of type Row
	 */
	public Row findRow(String filter) {
		return new Row(Dispatch.call(this, getIDOfName("FindRow"), filter).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Row
	 */
	public Row findNextRow() {
		return new Row(Dispatch.call(this, getIDOfName("FindNextRow")).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param maxRows an input-parameter of type int
	 * @return the result is of type Variant
	 */
	public Variant getArray(int maxRows) {
		return Dispatch.call(this, getIDOfName("GetArray"), new Variant(maxRows));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Row
	 */
	public Row getNextRow() {
		return new Row(Dispatch.call(this, getIDOfName("GetNextRow")).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type int
	 */
	public int getRowCount() {
		return Dispatch.call(this, getIDOfName("GetRowCount")).changeType(Variant.VariantInt).getInt();
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 */
	public void moveToStart() {
		Dispatch.call(this, getIDOfName("MoveToStart"));
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param filter an input-parameter of type String
	 * @return the result is of type Table
	 */
	public Table restrict(String filter) {
		return new Table(Dispatch.call(this, getIDOfName("Restrict"), filter).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param sortProperty an input-parameter of type String
	 * @param descending an input-parameter of type Variant
	 */
	public void sort(String sortProperty, Variant descending) {
		Dispatch.call(this, getIDOfName("Sort"), sortProperty, descending);
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param sortProperty an input-parameter of type String
	 */
	public void sort(String sortProperty) {
		Dispatch.call(this, getIDOfName("Sort"), sortProperty);
	}


	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type Columns
	 */
	public Columns getColumns() {
		return new Columns(Dispatch.get(this, getIDOfName("Columns")).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @return the result is of type boolean
	 */
	public boolean getEndOfTable() {
		return Dispatch.get(this, getIDOfName("EndOfTable")).changeType(Variant.VariantBoolean).getBoolean();
	}

}
