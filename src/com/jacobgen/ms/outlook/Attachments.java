/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public class Attachments extends Dispatch {

	public static final String componentName = "Outlook.Attachments";

	public Attachments() {
		super(componentName);
	}

	/**
	* This constructor is used instead of a case operation to
	* turn a Dispatch object into a wider object - it must exist
	* in every wrapper class whose instances may be returned from
	* method calls wrapped in VT_DISPATCH Variants.
	*/
	public Attachments(Dispatch d) {
		// take over the IDispatch pointer
		m_pDispatch = d.m_pDispatch;
		// null out the input's pointer
		d.m_pDispatch = 0;
	}

	public Attachments(String compName) {
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
	 * @return the result is of type Attachment
	 */
	public Attachment item(Variant index) {
		return new Attachment(Dispatch.call(this, "Item", index).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param source an input-parameter of type Variant
	 * @param type an input-parameter of type Variant
	 * @param position an input-parameter of type Variant
	 * @param displayName an input-parameter of type Variant
	 * @return the result is of type Attachment
	 */
	public Attachment add(Variant source, Variant type, Variant position, Variant displayName) {
		return new Attachment(Dispatch.call(this, "Add", source, type, position, displayName).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param source an input-parameter of type Variant
	 * @param type an input-parameter of type Variant
	 * @param position an input-parameter of type Variant
	 * @return the result is of type Attachment
	 */
	public Attachment add(Variant source, Variant type, Variant position) {
		return new Attachment(Dispatch.call(this, "Add", source, type, position).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param source an input-parameter of type Variant
	 * @param type an input-parameter of type Variant
	 * @return the result is of type Attachment
	 */
	public Attachment add(Variant source, Variant type) {
		return new Attachment(Dispatch.call(this, "Add", source, type).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param source an input-parameter of type Variant
	 * @return the result is of type Attachment
	 */
	public Attachment add(Variant source) {
		return new Attachment(Dispatch.call(this, "Add", source).toDispatch());
	}

	/**
	 * Wrapper for calling the ActiveX-Method with input-parameter(s).
	 * @param index an input-parameter of type int
	 */
	public void remove(int index) {
		Dispatch.call(this, "Remove", new Variant(index));
	}

}
